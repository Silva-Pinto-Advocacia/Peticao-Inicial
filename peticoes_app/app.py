"""
Silva Pinto Advocacia — Gerador Automático de Petições
Backend Flask com integração Claude API
Estratégia: extração de texto de todos os arquivos — zero PDFs nativos na API
"""

import os, sys, json, uuid, zipfile, shutil, logging, re, random, base64, gc
from pathlib import Path
from datetime import datetime

from flask import Flask, request, jsonify, send_file, render_template
import anthropic

# ── Config ──────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB

# ── Token budget ─────────────────────────────────────────────────────────────
MAX_CHARS_PER_PDF      = 1_500   # reduced for memory
MAX_CHARS_PER_XLSX     = 2_500
MAX_CHARS_DOCX_MODEL   = 50_000  # send full modelo for find/replace
MAX_CHARS_DOCX_OTHER   = 1_500
MAX_PDFS               = 3       # reduced for memory
TOTAL_CHAR_BUDGET      = 130_000  # raised: full modelo + ficha + relatórios

# ── Helpers ──────────────────────────────────────────────────────────────────

def xe(text: str) -> str:
    return (text.replace("&","&amp;").replace("<","&lt;")
                .replace(">","&gt;").replace('"',"&quot;").replace("'","&apos;"))

def random_para_id() -> str:
    return f"{random.randint(0x10000000, 0x7FFFFFFE):08X}"

def extract_xml_text(xml: str) -> str:
    text = re.sub(r"<[^>]+>", " ", xml)
    return re.sub(r"\s+", " ", text).strip()

def truncate(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + f"\n...[TRUNCADO — {len(text)-max_chars} chars omitidos]"

def extract_pdf_text(fpath: Path, max_chars: int = MAX_CHARS_PER_PDF) -> str:
    """Extract text from PDF using pypdf (lightweight, Python 3.14 safe)."""
    try:
        import pypdf
        text_parts = []
        with open(fpath, "rb") as f:
            reader = pypdf.PdfReader(f, strict=False)
            for page in reader.pages[:8]:
                try:
                    t = page.extract_text()
                    if t:
                        text_parts.append(t)
                except Exception:
                    pass
        full = "\n".join(text_parts)
        if not full.strip():
            return f"[PDF sem texto extraível: {fpath.name}]"
        return truncate(full, max_chars)
    except Exception as e:
        log.warning("pypdf failed for %s: %s", fpath.name, e)
        return f"[Não foi possível extrair texto: {fpath.name}]"

def read_docx_text(path: Path, max_chars: int = MAX_CHARS_DOCX_OTHER) -> str:
    """Read text from docx using python-docx (no subprocess)."""
    try:
        import docx as _docx
        doc = _docx.Document(str(path))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        parts.append(cell.text.strip())
        return truncate("\n".join(parts), max_chars)
    except Exception as e:
        log.warning("python-docx read error %s: %s", path.name, e)
        try:
            with zipfile.ZipFile(str(path)) as zf:
                xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
            return truncate(extract_xml_text(xml), max_chars)
        except Exception:
            return ""

def read_text_file(path: Path, max_chars: int = MAX_CHARS_PER_XLSX) -> str:
    try:
        # Try reading xlsx with openpyxl
        if path.suffix.lower() in (".xlsx", ".xls"):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(str(path), data_only=True)
                parts = []
                for sheet in wb.sheetnames:
                    ws = wb[sheet]
                    parts.append(f"[Planilha: {sheet}]")
                    for row in ws.iter_rows(values_only=True):
                        cells = [str(c) if c is not None else "" for c in row]
                        line = " | ".join(cells)
                        if line.strip(" |"):
                            parts.append(line)
                return truncate("\n".join(parts), max_chars)
            except Exception as e:
                log.warning("openpyxl failed: %s", e)
        with open(path, encoding="utf-8", errors="replace") as f:
            return truncate(f.read(), max_chars)
    except Exception:
        return ""

def unpack_docx(docx_path: Path, out_dir: Path) -> bool:
    """Unpack docx (zip) into directory."""
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(str(docx_path), 'r') as zf:
            zf.extractall(str(out_dir))
        return True
    except Exception as e:
        log.error("unpack error: %s", e)
        return False

def pack_docx(unpacked_dir: Path, out_path: Path, original: Path) -> bool:
    """Repack directory into docx (zip), preserving original file list order."""
    try:
        # Get file list from original to preserve content types order
        orig_names = set()
        with zipfile.ZipFile(str(original), 'r') as orig_zf:
            orig_names = set(orig_zf.namelist())

        with zipfile.ZipFile(str(out_path), 'w', zipfile.ZIP_DEFLATED) as zf:
            # Write [Content_Types].xml first (required by OOXML spec)
            ct = unpacked_dir / "[Content_Types].xml"
            if ct.exists():
                zf.write(str(ct), "[Content_Types].xml")
            # Write all other files
            for fpath in sorted(unpacked_dir.rglob("*")):
                if fpath.is_file() and fpath.name != "[Content_Types].xml":
                    arcname = str(fpath.relative_to(unpacked_dir))
                    zf.write(str(fpath), arcname)
        return True
    except Exception as e:
        log.error("pack error: %s", e)
        return False

# ── System Prompt ─────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """Você é o Assistente Jurídico Especialista do escritório Silva Pinto Advocacia.
Advogado responsável: Dr. Casil da Silva Pinto — OAB/RJ nº 189.781.
Especialidade: Ações anulatórias de atos administrativos em concursos públicos.

SUA FUNÇÃO:
Analisar os arquivos do cliente, cruzar com o resumo do caso e com a ficha do cliente,
e gerar um JSON estruturado para preenchimento da Petição Inicial em modelo .docx.

═══ INSTRUÇÕES DE EXECUÇÃO ═══

1. ANÁLISE E EXTRAÇÃO DE DADOS
- Leia todos os arquivos extraídos do ZIP do cliente.
- Identifique os documentos de qualificação (procuração, RG, CPF, comprovante de residência)
  e extraia: nome completo, nacionalidade, estado civil, profissão, RG, CPF, e-mail e
  endereço completo com CEP.
- Para qualquer dado essencial não encontrado, use a marcação curta "[DADO AUSENTE]" e
  inclua no array "dados_ausentes" a descrição do que está faltando.

2. REDAÇÃO JURÍDICA (Fatos, Direito, Pontuação)
- Linguagem altamente técnica, formal, persuasiva, adequada a Petição Inicial.
- Capítulo "Da Pontuação": MÁXIMO 3 parágrafos, sempre direto, indicando a pontuação
  que o candidato alcançará após anulação e dando ÊNFASE em que atingirá a nota de corte
  para avançar de fase.
- Capítulo "Da Probabilidade do Direito do Autor": insira na íntegra o relatório técnico
  de UMA questão (escolha aleatória dentre os relatórios disponíveis no ZIP),
  no campo "probabilidade_direito".
- Se o cliente NÃO tiver direito à gratuidade de justiça conforme a ficha, marque
  "gratuidade": false — o sistema removerá automaticamente o capítulo correspondente
  e o pedido de gratuidade do rol.

3. EXTRAÇÃO E RESUMO DAS QUESTÕES A ANULAR
- Identifique no arquivo "ficha do candidato" quais questões o autor pretende anular.
- Para cada questão indicada, leia o relatório técnico correspondente (arquivos
  nomeados como "Questões", "Relatório Técnico", "Parecer" etc.).
- Redija um parágrafo único, técnico e objetivo no campo "resumo_peticao", contendo:
  (a) o vício identificado (erro grosseiro, extrapolação do edital, duplicidade de
  gabarito, etc.), (b) a análise matemática ou jurídica que comprova o vício,
  (c) a consequência para o candidato.
- Se o relatório técnico já trouxer um resumo, utilize-o; senão, sintetize a
  fundamentação em um parágrafo claro, direto e convincente.
- Para cada questão, identifique o vício e nomeie no campo "vicio" (ex: "ERRO GROSSEIRO",
  "EXTRAPOLAÇÃO DO EDITAL", "DUPLICIDADE DE GABARITO", "AUSÊNCIA DE RESPOSTA CORRETA").
- O comando da questão e o enunciado devem ficar juntos em um único parágrafo no campo
  "enunciado". As alternativas devem ser listadas no array "alternativas",
  uma por elemento (ex: "A) texto", "B) texto"...).

4. ROL DE DOCUMENTOS
- Liste em "rol_documentos" todos os documentos que devem instruir a petição,
  numerados sequencialmente. Para cada item, indique o nome do arquivo correspondente
  no ZIP (campo "arquivo_correspondente") para que o sistema renomeie automaticamente.
- Inclua RG, CPF, Procuração, Comprovante de Residência, Cartão Resposta,
  Edital, e os Pareceres/Relatórios Técnicos de cada questão a ser anulada.

5. CONFERÊNCIAS OBRIGATÓRIAS
- O número de questões em "questoes" deve coincidir exatamente com as questões
  pedidas para anulação na ficha do candidato.
- A comarca em "processo.comarca" deve ser a do domicílio do cliente.
- Use o gênero correto do cliente em toda a redação (concordância nominal e verbal).
- Não invente jurisprudências — só use as que constam no modelo enviado.

═══ REGRAS INVIOLÁVEIS ═══

- Use somente dados reais extraídos dos documentos anexados — nunca invente.
- Não use jurisprudências ou citações genéricas; use apenas as do modelo.
- Diferencie petições por procedimento:
  * Procedimento Comum: o rol completo de questões já vai na inicial.
  * Tutela Cautelar Antecedente: pode informar que o rol completo virá em emenda.
- Todas as informações do candidato devem refletir os documentos do cliente atual,
  em substituição completa aos dados que constam no modelo.

═══ SUBSTITUIÇÃO DOS DADOS DO CLIENTE NO MODELO ═══

ATENÇÃO MÁXIMA: Esta é a tarefa MAIS CRÍTICA da sua função.
O modelo da petição vem com dados de um CLIENTE ANTIGO escritos diretamente no texto.
Você DEVE identificar TODOS esses trechos sem exceção e gerar pares
"buscar → substituir" para que o sistema faça a troca automaticamente.

TIPOS DE DADOS QUE PRECISAM SER SUBSTITUÍDOS (lista NÃO exaustiva — identifique tudo):

1. QUALIFICAÇÃO PESSOAL DO CLIENTE
   - Nome completo (em TODAS as ocorrências, inclusive em letras maiúsculas)
   - Estado civil (ex: "Casado", "Solteira")
   - Profissão (ex: "desempregado", "policial militar")
   - Número do RG (ex: "339924-4")
   - Número do CPF (ex: "140.948.387-83")
   - E-mail
   - Endereço completo (rua, número, bairro, complemento)
   - Cidade e UF
   - CEP

2. DADOS DO CONCURSO
   - Nome do concurso e edital (ex: "Edital nº 01/2025 — PCES")
   - Cargo pretendido (ex: "Oficial Investigador de Polícia")
   - Banca examinadora (ex: "IBADE", "FGV", "AOCP")
   - Tipo de prova realizada (ex: "Prova Tipo 2")
   - Datas das etapas

3. PONTUAÇÃO E CLASSIFICAÇÃO
   - Pontuação obtida pelo cliente antigo (ex: "55 pontos", "55 pts")
   - Pontuação após anulação (ex: "59 pts", "59 pontos")
   - Nota de corte (ex: "58 pontos") — se diferente para o novo cliente
   - Quaisquer outros números relacionados à pontuação

4. JURISDIÇÃO E PARTES
   - Comarca de endereçamento (ex: "BARRA DE SÃO FRANCISCO/ES")
   - Estado/UF que compõe o polo passivo
   - Tipo de ação (se diferente)

5. QUESTÕES IMPUGNADAS
   - Lista de números das questões antigas (ex: "questões 10, 25, 31 e 34")
   - Substituir pela lista das questões reais do novo cliente

REGRAS OBRIGATÓRIAS PARA OS PARES:

✅ PARES GRANULARES: Crie um par por dado individual em vez de um par gigante.
   - BOM: {"buscar": "ADLER MARQUES DE LIMA", "substituir": "JOÃO DA SILVA"}
   - BOM: {"buscar": "140.948.387-83", "substituir": "111.222.333-44"}
   - BOM: {"buscar": "55 pontos", "substituir": "60 pontos"}
   - RUIM: {"buscar": "ADLER MARQUES DE LIMA, Casado, desempregado, RG 339924-4...",
            "substituir": "JOÃO DA SILVA, Solteiro, professor, RG 555..."}

✅ TEXTO LITERAL: copie EXATAMENTE como aparece no modelo (acentos, maiúsculas,
   pontuação, espaços). Se aparece "ADLER MARQUES DE LIMA" em maiúsculas,
   crie um par com texto em maiúsculas. Se aparece "Adler Marques de Lima"
   em outro lugar, crie OUTRO par para essa variação.

✅ BUSQUE TODAS AS OCORRÊNCIAS: percorra o modelo inteiro mentalmente e
   procure cada dado em CADA seção (cabeçalho, qualificação, fatos, pontuação,
   pedidos, etc.). NUNCA assuma que substituir uma vez resolve tudo —
   se o sistema vai trocar todas as ocorrências de uma string, mas se o nome
   aparece em formatos diferentes (maiúsculo, minúsculo, abreviado), cada
   formato precisa de seu próprio par.

✅ NÃO INCLUA DUPLICATAS: se a string idêntica aparece 5 vezes, basta UM par.

✅ MÍNIMO 15 PARES: para um modelo típico de petição de concurso, você deve
   gerar AO MENOS 15-25 pares de substituição. Se gerar menos de 10, é sinal
   de que está deixando dados antigos passar.

JSON PURO (sem markdown, sem backticks, sem texto antes ou depois).
REGRAS CRÍTICAS PARA O JSON:
- Use APENAS aspas duplas para strings, nunca aspas simples ou aspas tipográficas.
- Dentro de strings, escape quebras de linha como \\n (nunca quebra de linha real).
- Dentro de strings, escape aspas duplas como \\".
- NÃO inclua vírgulas após o último item de objetos ou arrays.
- Não use comentários no JSON.
- Mantenha valores curtos e objetivos para evitar exceder o limite de tokens.
- Antes de responder, valide mentalmente a sintaxe do JSON.

Schema:
{
  "cliente": {
    "nome_completo": "",
    "nacionalidade": "",
    "estado_civil": "",
    "profissao": "",
    "rg": "",
    "cpf": "",
    "email": "",
    "endereco": "",
    "cidade": "",
    "uf": "",
    "cep": "",
    "genero": "M ou F"
  },
  "processo": {
    "tipo_acao": "",
    "comarca": "",
    "banca": "",
    "cargo": "",
    "concurso": "",
    "pontuacao_obtida": "",
    "pontuacao_corte": "",
    "pontuacao_apos_anulacao": "",
    "categoria": "ampla_concorrencia ou cotas",
    "gratuidade": true,
    "procedimento": "comum ou cautelar_antecedente"
  },
  "questoes": [
    {
      "numero": 0,
      "vicio": "",
      "resumo_peticao": "",
      "enunciado": "",
      "alternativas": ["A) ...", "B) ...", "C) ...", "D) ...", "E) ..."],
      "gabarito_banca": "",
      "resposta_correta": "",
      "relatorio_integra": ""
    }
  ],
  "questao_destaque_idx": 0,
  "substituicoes": [
    {"buscar": "ADLER MARQUES DE LIMA", "substituir": "NOME DO NOVO CLIENTE"},
    {"buscar": "55 pontos", "substituir": "60 pontos"}
  ],
  "textos": {
    "fatos": "",
    "pontuacao": "",
    "probabilidade_direito": "",
    "fundamentos_juridicos": ""
  },
  "rol_documentos": [
    {"numero": 1, "descricao": "", "arquivo_correspondente": ""}
  ],
  "dados_ausentes": [],
  "relatorio_alteracoes": []
}"""

# ── Claude call ───────────────────────────────────────────────────────────────

def _parse_claude_json(raw: str) -> dict:
    """Robustly parse JSON from Claude's response, handling common malformations."""
    # Strip markdown fences
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    raw = raw.strip()

    # Extract just the JSON object — find the outermost { ... }
    start = raw.find("{")
    end   = raw.rfind("}")
    if start >= 0 and end > start:
        raw = raw[start:end+1]

    # First attempt: parse as-is
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e1:
        log.warning("JSON decode failed at first try: %s", e1)

    # Second attempt: fix common issues
    fixed = raw
    # Remove trailing commas before } or ]
    fixed = re.sub(r",(\s*[}\]])", r"\1", fixed)
    # Replace smart quotes that may have leaked in
    fixed = fixed.replace("\u201c", '"').replace("\u201d", '"')
    fixed = fixed.replace("\u2018", "'").replace("\u2019", "'")
    # Newlines inside strings — most common cause of "Expecting ',' delimiter"
    # is unescaped newlines in JSON string values. Try escaping them.
    try:
        return json.loads(fixed)
    except json.JSONDecodeError as e2:
        log.warning("JSON decode failed at second try: %s", e2)

    # Third attempt: aggressive — try to fix unescaped newlines inside strings
    # by parsing line-by-line and joining quoted strings
    try:
        # Replace literal newlines inside quoted strings with \\n
        result = []
        in_string = False
        escape = False
        for ch in fixed:
            if escape:
                result.append(ch)
                escape = False
                continue
            if ch == "\\":
                result.append(ch)
                escape = True
                continue
            if ch == '"':
                in_string = not in_string
                result.append(ch)
                continue
            if in_string and ch == "\n":
                result.append("\\n")
                continue
            if in_string and ch == "\r":
                result.append("\\r")
                continue
            if in_string and ch == "\t":
                result.append("\\t")
                continue
            result.append(ch)
        fixed = "".join(result)
        return json.loads(fixed)
    except json.JSONDecodeError as e3:
        log.error("All JSON parse attempts failed. Last error: %s", e3)
        log.error("Raw content (first 500 chars): %s", raw[:500])
        log.error("Raw content (last 500 chars): %s", raw[-500:])
        raise


def call_claude(api_key: str, full_text: str) -> dict:
    """Single text-only call to Claude. No PDFs, no binary — pure text."""
    # 5-minute timeout: longer than gunicorn's, so Claude has time to respond
    client = anthropic.Anthropic(api_key=api_key, timeout=300.0, max_retries=2)
    log.info("Sending %d chars to Claude", len(full_text))
    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=16000,  # increased to avoid JSON truncation
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": full_text}]
    )
    log.info("Claude response: %d input / %d output tokens. Stop reason: %s",
             message.usage.input_tokens if message.usage else 0,
             message.usage.output_tokens if message.usage else 0,
             message.stop_reason)

    # Detect truncation
    if message.stop_reason == "max_tokens":
        raise ValueError(
            "A resposta do Claude foi truncada por exceder o limite de tokens. "
            "Tente um ZIP com menos arquivos ou peça em duas etapas."
        )

    raw = "".join(b.text for b in message.content if hasattr(b, "text"))
    return _parse_claude_json(raw)

# ── DOCX editing ──────────────────────────────────────────────────────────────

def apply_substitutions(unpacked_dir: Path, data: dict) -> list[str]:
    """Apply find/replace pairs from Claude across XML and headers/footers."""
    changes = []

    # Files to edit: document.xml + headers + footers
    word_dir = unpacked_dir / "word"
    files_to_edit = []
    if (word_dir / "document.xml").exists():
        files_to_edit.append(word_dir / "document.xml")
    for f in word_dir.glob("header*.xml"):
        files_to_edit.append(f)
    for f in word_dir.glob("footer*.xml"):
        files_to_edit.append(f)

    if not files_to_edit:
        return ["ERRO: nenhum XML editável encontrado"]

    # Get find/replace pairs from Claude response
    pairs = data.get("substituicoes", [])
    log.info("Total de substituições a aplicar: %d", len(pairs))
    if not pairs:
        changes.append("⚠️ Nenhum par de substituição retornado pelo Claude")

    # Apply each pair across all XML files
    for pair_idx, pair in enumerate(pairs):
        old = pair.get("buscar", "").strip()
        new = pair.get("substituir", "").strip()
        if not old:
            log.warning("Par #%d: 'buscar' vazio, pulando", pair_idx)
            continue
        if old == new:
            log.info("Par #%d: buscar==substituir, pulando", pair_idx)
            continue

        log.info("Par #%d: buscando '%s' (len=%d)", pair_idx, old[:80], len(old))

        old_xe = xe(old)
        new_xe = xe(new)
        total_count = 0
        method_used = ""

        for xml_path in files_to_edit:
            xml = xml_path.read_text(encoding="utf-8")
            file_count = 0

            # Method 1: literal XML-escaped match
            count = xml.count(old_xe)
            if count > 0:
                xml = xml.replace(old_xe, new_xe)
                file_count += count
                method_used = "literal"
                xml_path.write_text(xml, encoding="utf-8")

            # Method 2: cross-run replacement (loop until no more matches)
            else:
                attempts = 0
                while attempts < 20:  # safety limit
                    new_xml = _replace_across_runs(xml, old, new)
                    if new_xml == xml:
                        break
                    xml = new_xml
                    file_count += 1
                    attempts += 1
                    method_used = "cross-run"
                if file_count > 0:
                    xml_path.write_text(xml, encoding="utf-8")

            total_count += file_count
            if file_count > 0:
                log.info("  %s: %d replacement(s) via %s",
                         xml_path.name, file_count, method_used)

        if total_count > 0:
            changes.append(f"✅ [{method_used}] '{old[:50]}' → '{new[:50]}' ({total_count}x)")
        else:
            log.warning("  Não encontrado em nenhum arquivo: '%s'", old[:80])
            changes.append(f"⚠️ Não encontrado: '{old[:60]}'")

    # Remove gratuidade chapter if not applicable
    p = data.get("processo", {})
    if not p.get("gratuidade", True):
        for xml_path in files_to_edit:
            xml = xml_path.read_text(encoding="utf-8")
            new_xml, removed = _remove_gratuidade_chapter(xml)
            if removed:
                xml_path.write_text(new_xml, encoding="utf-8")
                changes.append("🗑️ Capítulo de gratuidade removido")
                break

    return changes


def _replace_across_runs(xml: str, old: str, new: str) -> str:
    """Try to replace text that may be split across multiple <w:t> elements."""
    if not old:
        return xml

    # Extract all visible text spans
    pattern = re.compile(r'<w:t[^>]*>([^<]*)</w:t>', re.DOTALL)
    matches = list(pattern.finditer(xml))
    if not matches:
        return xml

    # Build concatenated text and position map
    concat = ""
    positions = []
    for i, m in enumerate(matches):
        text = m.group(1)
        start = len(concat)
        concat += text
        positions.append((start, start + len(text), i))

    # Try direct find first
    idx = concat.find(old)

    # Whitespace-tolerant fallback: collapse whitespace in both
    if idx < 0:
        norm_old = re.sub(r"\s+", " ", old).strip()
        norm_concat = re.sub(r"\s+", " ", concat).strip()
        if norm_old in norm_concat:
            # Find the actual range in concat by walking char by char
            # building a normalized version while tracking original positions
            def find_norm(haystack: str, needle: str):
                """Returns (start_in_haystack, end_in_haystack) for a whitespace-normalised match."""
                # Build mapping of original_index -> normalized_index
                norm_chars = []
                orig_indices = []
                last_was_space = True
                for i, ch in enumerate(haystack):
                    if ch.isspace():
                        if not last_was_space:
                            norm_chars.append(" ")
                            orig_indices.append(i)
                            last_was_space = True
                    else:
                        norm_chars.append(ch)
                        orig_indices.append(i)
                        last_was_space = False
                norm_str = "".join(norm_chars).strip()
                # Find needle (also normalised)
                n_idx = norm_str.find(needle)
                if n_idx < 0:
                    return None
                # Account for leading whitespace stripped
                leading_strip = len(norm_str) - len(norm_str.lstrip()) if False else 0
                # Map back to original
                start_orig = orig_indices[n_idx] if n_idx < len(orig_indices) else None
                end_norm   = n_idx + len(needle) - 1
                end_orig   = orig_indices[end_norm] + 1 if end_norm < len(orig_indices) else len(haystack)
                return (start_orig, end_orig)

            result = find_norm(concat, norm_old)
            if result:
                idx, end_idx_calc = result
            else:
                return xml
        else:
            return xml
    else:
        end_idx_calc = idx + len(old)

    end_idx = end_idx_calc

    # Identify runs that overlap [idx, end_idx)
    affected = [(s, e, i) for (s, e, i) in positions if not (e <= idx or s >= end_idx)]
    if not affected:
        return xml

    affected_idxs = [a[2] for a in affected]
    first_affected = affected[0][2]

    # Build new XML
    new_xml_parts = []
    last_pos = 0
    for j, m in enumerate(matches):
        new_xml_parts.append(xml[last_pos:m.start()])
        if j == first_affected:
            run_start, run_end, _ = positions[j]
            prefix = m.group(1)[:max(0, idx - run_start)]
            suffix = m.group(1)[max(0, end_idx - run_start):] if run_end > end_idx else ""
            new_text = prefix + new + suffix
            new_xml_parts.append(f'<w:t xml:space="preserve">{xe(new_text)}</w:t>')
        elif j in affected_idxs:
            run_start, run_end, _ = positions[j]
            if run_end <= end_idx:
                new_xml_parts.append('<w:t xml:space="preserve"></w:t>')
            else:
                suffix = m.group(1)[end_idx - run_start:]
                new_xml_parts.append(f'<w:t xml:space="preserve">{xe(suffix)}</w:t>')
        else:
            new_xml_parts.append(m.group(0))
        last_pos = m.end()
    new_xml_parts.append(xml[last_pos:])
    return "".join(new_xml_parts)


def _build_qualificacao(c: dict) -> str:
    parts = []
    if c.get("nacionalidade"): parts.append(c["nacionalidade"])
    if c.get("estado_civil"):  parts.append(c["estado_civil"])
    if c.get("profissao"):     parts.append(c["profissao"])
    if c.get("rg"):            parts.append(f"portador do RG nº {c['rg']}")
    if c.get("cpf"):           parts.append(f"inscrito no CPF/MF sob o nº {c['cpf']}")
    if c.get("email"):         parts.append(f"e-mail {c['email']}")
    addr = _build_endereco(c)
    if addr:                   parts.append(f"residente e domiciliado em {addr}")
    return ", ".join(parts)

def _build_endereco(c: dict) -> str:
    return ", ".join(p for p in [c.get("endereco",""), c.get("cidade",""), c.get("uf",""), c.get("cep","")] if p)

def _replace_block(xml: str, marker: str, new_text: str, changes: list) -> str:
    if marker not in xml or not new_text:
        return xml
    xml = xml.replace(marker, _text_to_paragraphs(new_text))
    changes.append(f"✅ Bloco: {marker}")
    return xml

def _text_to_paragraphs(text: str) -> str:
    paras = [p.strip() for p in text.split("\n\n") if p.strip()]
    result = []
    for para in paras:
        pid = random_para_id()
        result.append(
            f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000">'
            f'<w:pPr><w:ind w:left="0"/><w:jc w:val="both"/></w:pPr>'
            f'<w:r><w:t xml:space="preserve">{xe(para.replace(chr(10)," "))}</w:t></w:r>'
            f'</w:p>'
        )
    return "\n".join(result)

def _insert_questoes(xml: str, questoes: list, changes: list) -> str:
    if not questoes:
        return xml
    marker = next((m for m in ["BLOCO_QUESTOES_ILEGAIS","DAS_QUESTOES_ILEGAIS","QUESTOES_ANULAVEIS"] if m in xml), None)
    if not marker:
        changes.append("⚠️ Marcador de questões não encontrado no modelo")
        return xml

    blocks = []
    for q in sorted(questoes, key=lambda x: x.get("numero", 0)):
        num, vicio = q.get("numero","?"), q.get("vicio","VÍCIO").upper()
        pid = random_para_id()
        blocks.append(
            f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000">'
            f'<w:pPr><w:jc w:val="both"/><w:rPr><w:b/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:b/></w:rPr><w:t>{xe(f"QUESTÃO {num} — {vicio}")}</w:t></w:r></w:p>'
        )
        if q.get("enunciado"):
            pid = random_para_id()
            blocks.append(
                f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000">'
                f'<w:pPr><w:jc w:val="both"/></w:pPr>'
                f'<w:r><w:t xml:space="preserve">{xe(q["enunciado"])}</w:t></w:r></w:p>'
            )
        for alt in q.get("alternativas", []):
            pid = random_para_id()
            blocks.append(
                f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000">'
                f'<w:pPr><w:ind w:left="720"/><w:jc w:val="both"/></w:pPr>'
                f'<w:r><w:t xml:space="preserve">{xe(alt)}</w:t></w:r></w:p>'
            )
        if q.get("resumo_peticao"):
            pid = random_para_id()
            blocks.append(
                f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000">'
                f'<w:pPr><w:jc w:val="both"/></w:pPr>'
                f'<w:r><w:t xml:space="preserve">{xe(q["resumo_peticao"])}</w:t></w:r></w:p>'
            )
        pid = random_para_id()
        blocks.append(f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000"><w:pPr/></w:p>')

    xml = xml.replace(marker, "\n".join(blocks))
    changes.append(f"✅ {len(questoes)} questão(ões) inserida(s)")
    return xml

def _remove_gratuidade_chapter(xml: str):
    pat = r'<w:p[^>]*>.*?gratuidade.*?</w:p>\s*(<w:p[^>]*>.*?</w:p>\s*)*?(?=<w:p[^>]*>[^<]*(?:DA PROBABILIDADE|DO DIREITO|DOS PEDIDOS))'
    new_xml, n = re.subn(pat, "", xml, flags=re.IGNORECASE | re.DOTALL)
    return (new_xml, True) if n else (xml, False)

# ── File renaming ─────────────────────────────────────────────────────────────

def rename_files_by_rol(work_dir: Path, rol: list) -> dict:
    mapping = {}
    files_in_dir = {f.name.lower(): f for f in work_dir.iterdir() if f.is_file()}
    for item in rol:
        num  = item.get("numero", "")
        desc = item.get("descricao", "")
        orig = item.get("arquivo_correspondente", "")
        safe = re.sub(r"\s+", "_", re.sub(r"[^\w\s\-]", "", desc).strip())
        src  = None
        if orig:
            candidate = work_dir / orig
            if candidate.exists():
                src = candidate
        if not src:
            orig_l = orig.lower()
            for fl, fp in files_in_dir.items():
                if orig_l in fl or fl in orig_l:
                    src = fp
                    break
        if src and src.exists():
            new_name = f"{num}. {safe}{src.suffix}"
            src.rename(work_dir / new_name)
            mapping[src.name] = new_name
        else:
            mapping[f"[AUSENTE] {orig}"] = f"{num}. {desc} — NÃO ENCONTRADO"
    return mapping

# ── Pipeline ──────────────────────────────────────────────────────────────────

def process_zip(zip_path: Path, session_dir: Path, form_data: dict, api_key: str) -> dict:
    log.info("Pipeline start: %s", zip_path.name)

    # 1. Extract — handle Brazilian Portuguese filenames with cp437/latin-1 encoding
    extract_dir = session_dir / "extracted"
    extract_dir.mkdir()
    with zipfile.ZipFile(zip_path) as zf:
        for info in zf.infolist():
            # ZIP spec uses cp437 by default; many tools encode filenames as latin-1 or utf-8
            try:
                # Try to decode the raw bytes as utf-8 first, then fall back
                raw = info.filename.encode("cp437", errors="replace")
                try:
                    fixed_name = raw.decode("utf-8")
                except UnicodeDecodeError:
                    fixed_name = raw.decode("latin-1", errors="replace")
            except Exception:
                fixed_name = info.filename
            # Sanitize: ASCII-safe filename for filesystem
            safe_name = re.sub(r"[^\w\-./ ]", "_", fixed_name)
            target = extract_dir / safe_name
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                with zf.open(info) as src, open(target, "wb") as dst:
                    shutil.copyfileobj(src, dst)

    all_files = [f for f in extract_dir.rglob("*") if f.is_file()]
    file_list = [str(f.relative_to(extract_dir)) for f in all_files]
    log.info("Files: %s", file_list)

    # 2. Classify
    model_keywords  = ("modelo", "peticao", "petição", "inicial")
    report_keywords = ("relatorio", "relatório", "questao", "questão", "tecnico", "parecer", "fundamentacao")
    id_keywords     = ("rg", "cpf", "procuracao", "identidade", "comprovante", "cnh", "habilitacao")

    docx_model = None
    other_docx = []
    pdfs_report = []
    pdfs_id = []
    text_files = []

    for fpath in all_files:
        nl = fpath.name.lower()
        if nl.endswith(".docx"):
            if any(k in nl for k in model_keywords):
                docx_model = fpath
            else:
                other_docx.append(fpath)
        elif nl.endswith(".pdf"):
            if any(k in nl for k in id_keywords):
                pdfs_id.append(fpath)
            else:
                pdfs_report.append(fpath)
        elif nl.endswith((".xlsx", ".xls", ".csv", ".txt")):
            text_files.append(fpath)

    if not docx_model and other_docx:
        docx_model = other_docx.pop(0)
    if not docx_model:
        return {"error": "Nenhum arquivo .docx modelo encontrado no ZIP."}

    # 3. Build text block — ALL text, NO binary PDFs
    parts = []
    char_used = 0

    def add(section: str, content: str, cap: int):
        nonlocal char_used
        remaining = TOTAL_CHAR_BUDGET - char_used
        if remaining <= 100:
            return
        snip = truncate(content, min(cap, remaining))
        parts.append(section + "\n" + snip)
        char_used += len(snip)

    parts.append("=== INVENTÁRIO DO ZIP ===\n" + "\n".join(f"  • {f}" for f in file_list))
    parts.append(f"""=== DADOS FORNECIDOS PELO USUÁRIO ===
Tipo de Ação: {form_data.get('tipo_acao','')}
Comarca: {form_data.get('comarca','')}
Resumo dos Fatos: {form_data.get('fatos','')}
Pedidos Adicionais: {form_data.get('pedidos','')}
Observações: {form_data.get('obs','')}""")
    char_used = sum(len(p) for p in parts)

    # Text files first (ficha xlsx — most critical)
    for fpath in text_files:
        rname = fpath.relative_to(extract_dir)
        add(f"\n=== TEXTO: {rname} ===", read_text_file(fpath), MAX_CHARS_PER_XLSX)

    # Other DOCX (relatórios em docx, fichas)
    for fpath in other_docx:
        rname = fpath.relative_to(extract_dir)
        add(f"\n=== DOCX: {rname} ===", read_docx_text(fpath, MAX_CHARS_DOCX_OTHER), MAX_CHARS_DOCX_OTHER)

    # Report PDFs — text extraction only, no binary upload
    for fpath in (pdfs_report + pdfs_id)[:MAX_PDFS]:
        rname = fpath.relative_to(extract_dir)
        add(f"\n=== PDF: {rname} ===", extract_pdf_text(fpath), MAX_CHARS_PER_PDF)

    # Modelo docx — send the FULL TEXT so Claude can identify what to replace
    modelo_text = read_docx_text(docx_model, max_chars=50_000)  # full modelo
    parts.append(
        f"\n=== MODELO DA PETIÇÃO (texto integral) — arquivo: {docx_model.name} ===\n"
        f"ATENÇÃO: Identifique TODOS os trechos do modelo abaixo que se referem ao CLIENTE ANTIGO\n"
        f"(nome, RG, CPF, endereço, pontuação, número do edital, banca, cargo, comarca, etc.)\n"
        f"e gere pares de 'buscar' (texto antigo) → 'substituir' (texto novo do cliente atual)\n"
        f"no campo 'substituicoes' do JSON. Inclua TAMBÉM o tipo de ação se for diferente.\n\n"
        f"{modelo_text}"
    )

    full_text = "\n\n".join(parts)
    log.info("Text block: %d chars (budget: %d)", len(full_text), TOTAL_CHAR_BUDGET)

    # Free memory before API call
    del parts
    gc.collect()

    # 4. Call Claude (text only)
    data = call_claude(api_key, full_text)
    log.info("Claude OK")
    del full_text
    gc.collect()

    # 5. Unpack modelo
    unpacked_dir = session_dir / "unpacked"
    if not unpack_docx(docx_model, unpacked_dir):
        return {"error": "Falha ao desempacotar o modelo DOCX."}

    # 6. Apply edits
    changes = apply_substitutions(unpacked_dir, data)

    # 7. Repack
    safe_nome  = re.sub(r"[^\w\s]","",data.get("cliente",{}).get("nome_completo","Cliente")).replace(" ","_")
    tipo_acao  = re.sub(r"[^\w\s]","",data.get("processo",{}).get("tipo_acao","Peticao")).replace(" ","_")
    data_hoje  = datetime.now().strftime("%Y-%m-%d")
    out_name   = f"Peticao_{tipo_acao}_{safe_nome}_{data_hoje}.docx"
    out_path   = session_dir / out_name

    if not pack_docx(unpacked_dir, out_path, docx_model):
        return {"error": "Falha ao reempacotar o DOCX final."}

    # 8. Delivery dir
    delivery = session_dir / "entrega"
    delivery.mkdir()
    shutil.copy(out_path, delivery / out_name)
    for fpath in all_files:
        if fpath != docx_model:
            dest = delivery / fpath.name
            if not dest.exists():
                shutil.copy(fpath, dest)

    # 9. Rename
    rol = data.get("rol_documentos", [])
    rename_map = rename_files_by_rol(delivery, rol) if rol else {}

    # 10. Final zip
    zip_name = f"Entrega_{safe_nome}_{data_hoje}.zip"
    zip_out  = OUTPUT_DIR / zip_name
    with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as zf:
        for fpath in delivery.iterdir():
            if fpath.is_file():
                zf.write(fpath, fpath.name)

    log.info("Done: %s", zip_name)
    return {
        "success":        True,
        "zip_filename":   zip_name,
        "docx_filename":  out_name,
        "cliente":        data.get("cliente", {}),
        "processo":       data.get("processo", {}),
        "questoes":       [{"numero": q.get("numero"), "vicio": q.get("vicio")} for q in data.get("questoes", [])],
        "changes":        changes,
        "rename_map":     rename_map,
        "dados_ausentes": data.get("dados_ausentes", []),
        "relatorio":      data.get("relatorio_alteracoes", []),
    }

# ── Global error handlers ─────────────────────────────────────────────────────

@app.errorhandler(Exception)
def handle_exception(e):
    log.exception("Unhandled exception")
    return jsonify({"error": f"Erro interno: {str(e)}"}), 500

@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "Arquivo muito grande. Limite: 100 MB."}), 413

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/health")
def health():
    """Healthcheck + diagnostic endpoint."""
    import sys
    checks = {}
    try:
                checks["pypdf"] = "ok"
    except ImportError as e:
        checks["pypdf"] = f"MISSING: {e}"
    try:
        import openpyxl
        checks["openpyxl"] = "ok"
    except ImportError as e:
        checks["openpyxl"] = f"MISSING: {e}"
    try:
        import anthropic
        checks["anthropic"] = "ok"
    except ImportError as e:
        checks["anthropic"] = f"MISSING: {e}"
    checks["unpack_py"]  = "ok" if (SCRIPTS / "unpack.py").exists() else "MISSING"
    checks["pack_py"]    = "ok" if (SCRIPTS / "pack.py").exists() else "MISSING"
    checks["upload_dir"] = str(UPLOAD_DIR)
    checks["output_dir"] = str(OUTPUT_DIR)
    checks["python"]     = sys.version
    return jsonify(checks)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/gerar", methods=["POST"])
def gerar():
    import traceback
    try:
        api_key = request.form.get("api_key", "").strip()
        if not api_key:
            return jsonify({"error": "Chave de API não fornecida."}), 400
        if "zip_file" not in request.files or request.files["zip_file"].filename == "":
            return jsonify({"error": "Arquivo ZIP não enviado."}), 400
        zip_file = request.files["zip_file"]
        if not zip_file.filename.lower().endswith(".zip"):
            return jsonify({"error": "Apenas arquivos .zip são aceitos."}), 400

        session_id  = str(uuid.uuid4())
        session_dir = UPLOAD_DIR / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        zip_path = session_dir / "input.zip"

        log.info("Saving uploaded ZIP...")
        zip_file.save(str(zip_path))
        log.info("ZIP saved: %s bytes", zip_path.stat().st_size)

        form_data = {k: request.form.get(k, "") for k in ("tipo_acao","comarca","fatos","pedidos","obs")}

        try:
            result = process_zip(zip_path, session_dir, form_data, api_key)
        except anthropic.AuthenticationError:
            return jsonify({"error": "Chave de API inválida. Verifique suas credenciais em console.anthropic.com"}), 401
        except anthropic.RateLimitError:
            return jsonify({"error": "Limite de uso da API Anthropic atingido. Aguarde alguns minutos."}), 429
        except anthropic.APIStatusError as e:
            log.exception("Anthropic API error")
            return jsonify({"error": f"Erro da API Anthropic: {e.status_code} — {e.message}"}), 502
        except json.JSONDecodeError as e:
            log.exception("JSON decode error")
            return jsonify({"error": f"Claude não retornou JSON válido: {e}"}), 500
        except Exception as e:
            log.exception("Pipeline error")
            tb = traceback.format_exc()
            return jsonify({"error": str(e), "traceback": tb[-2000:]}), 500

        if "error" in result:
            return jsonify(result), 500

        result["session_id"] = session_id
        log.info("Request completed OK: %s", result.get("zip_filename"))
        return jsonify(result)

    except Exception as e:
        log.exception("Outer exception in /gerar")
        tb = traceback.format_exc()
        return jsonify({"error": f"Erro inesperado: {str(e)}", "traceback": tb[-2000:]}), 500

@app.route("/download/<session_id>/<filename>")
def download(session_id, filename):
    if not re.match(r"^[\w\-]+$", session_id):
        return "ID inválido", 400
    for base in [OUTPUT_DIR, UPLOAD_DIR / session_id]:
        path = base / filename
        if path.exists():
            return send_file(str(path), as_attachment=True, download_name=filename)
    return "Arquivo não encontrado", 404

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
