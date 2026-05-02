"""
Silva Pinto Advocacia — Gerador Automático de Petições
Backend Flask com integração Claude API
Estratégia: extração de texto de todos os arquivos — zero PDFs nativos na API
"""

import os, sys, json, uuid, zipfile, shutil, logging, re, random, base64
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
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

# ── Token budget ─────────────────────────────────────────────────────────────
MAX_CHARS_PER_PDF      = 3_000   # ~750 tokens per PDF
MAX_CHARS_PER_XLSX     = 4_000
MAX_CHARS_DOCX_MODEL   = 4_000
MAX_CHARS_DOCX_OTHER   = 2_500
MAX_PDFS               = 6       # max number of PDFs to include
TOTAL_CHAR_BUDGET      = 60_000  # hard ceiling for entire text block (~15k tokens)

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

SUAS RESPONSABILIDADES:
1. Analisar os textos extraídos dos documentos do cliente (RG, CPF, procuração, ficha XLSX, relatórios).
2. Identificar quais questões serão contestadas (consta na ficha do candidato).
3. Redigir blocos de Fatos e Fundamentação em linguagem jurídica técnica, formal e persuasiva.
4. No capítulo "Da Pontuação": máximo 3 parágrafos, indicar pontuação que o candidato alcançará e destacar que atingirá a nota de corte.
5. No capítulo "Da Probabilidade do Direito": inserir na íntegra o relatório técnico de UMA questão (escolha aleatória dentre as disponíveis).
6. Se o cliente NÃO tiver direito à gratuidade (conforme ficha), indicar gratuidade: false.
7. Verificar concordância de gênero em todo o texto gerado.
8. Verificar se a comarca está correta conforme domicílio do cliente.
9. Sinalizar campos ausentes com [DADO AUSENTE — descrição].

FORMATO DE RESPOSTA — JSON PURO (sem markdown, sem backticks, sem texto antes ou depois):
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
    "gratuidade": true
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

def call_claude(api_key: str, full_text: str) -> dict:
    """Single text-only call to Claude. No PDFs, no binary — pure text."""
    client = anthropic.Anthropic(api_key=api_key)
    log.info("Sending %d chars to Claude", len(full_text))
    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": full_text}]
    )
    raw = "".join(b.text for b in message.content if hasattr(b, "text"))
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    return json.loads(raw.strip())

# ── DOCX editing ──────────────────────────────────────────────────────────────

def apply_substitutions(unpacked_dir: Path, data: dict) -> list[str]:
    doc_xml_path = unpacked_dir / "word" / "document.xml"
    if not doc_xml_path.exists():
        return ["ERRO: document.xml não encontrado"]

    xml = doc_xml_path.read_text(encoding="utf-8")
    changes = []
    c = data.get("cliente", {})
    p = data.get("processo", {})
    t = data.get("textos", {})

    substitutions = [
        ("NOME_COMPLETO_CLIENTE",   c.get("nome_completo", "[DADO AUSENTE — Nome]")),
        ("QUALIFICACAO_CLIENTE",    _build_qualificacao(c)),
        ("CPF_CLIENTE",             c.get("cpf", "[DADO AUSENTE — CPF]")),
        ("RG_CLIENTE",              c.get("rg", "[DADO AUSENTE — RG]")),
        ("EMAIL_CLIENTE",           c.get("email", "[DADO AUSENTE — E-mail]")),
        ("ENDERECO_CLIENTE",        _build_endereco(c)),
        ("CIDADE_CLIENTE",          c.get("cidade", "[DADO AUSENTE — Cidade]")),
        ("UF_CLIENTE",              c.get("uf", "[DADO AUSENTE — UF]")),
        ("COMARCA_JUIZO",           p.get("comarca", "[DADO AUSENTE — Comarca]")),
        ("TIPO_ACAO",               p.get("tipo_acao", "")),
        ("NOME_BANCA",              p.get("banca", "[DADO AUSENTE — Banca]")),
        ("NOME_CONCURSO",           p.get("concurso", "[DADO AUSENTE — Concurso]")),
        ("CARGO_PRETENDIDO",        p.get("cargo", "[DADO AUSENTE — Cargo]")),
        ("PONTUACAO_OBTIDA",        str(p.get("pontuacao_obtida", "[DADO AUSENTE]"))),
        ("PONTUACAO_CORTE",         str(p.get("pontuacao_corte", "[DADO AUSENTE]"))),
        ("PONTUACAO_APOS_ANULACAO", str(p.get("pontuacao_apos_anulacao", "[DADO AUSENTE]"))),
    ]

    for placeholder, value in substitutions:
        if placeholder in xml:
            xml = xml.replace(placeholder, xe(value))
            changes.append(f"✅ {placeholder}")

    xml = _replace_block(xml, "BLOCO_FATOS",                 t.get("fatos", ""), changes)
    xml = _replace_block(xml, "BLOCO_PONTUACAO",             t.get("pontuacao", ""), changes)
    xml = _replace_block(xml, "BLOCO_PROBABILIDADE_DIREITO", t.get("probabilidade_direito", ""), changes)
    xml = _replace_block(xml, "BLOCO_FUNDAMENTOS",           t.get("fundamentos_juridicos", ""), changes)
    xml = _insert_questoes(xml, data.get("questoes", []), changes)

    if not p.get("gratuidade", True):
        xml, removed = _remove_gratuidade_chapter(xml)
        if removed:
            changes.append("🗑️ Capítulo de gratuidade removido")

    doc_xml_path.write_text(xml, encoding="utf-8")
    return changes

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

    # 1. Extract
    extract_dir = session_dir / "extracted"
    extract_dir.mkdir()
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(extract_dir)

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

    # Modelo docx — skeleton only (placeholders)
    add(f"\n=== MODELO DOCX: {docx_model.name} (esqueleto) ===",
        read_docx_text(docx_model, MAX_CHARS_DOCX_MODEL), MAX_CHARS_DOCX_MODEL)

    full_text = "\n\n".join(parts)
    log.info("Text block: %d chars (budget: %d)", len(full_text), TOTAL_CHAR_BUDGET)

    # 4. Call Claude (text only)
    data = call_claude(api_key, full_text)
    log.info("Claude OK")

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
