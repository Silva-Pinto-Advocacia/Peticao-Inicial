"""
Silva Pinto Advocacia — Gerador Automático de Petições
Backend Flask com integração Claude API
"""

import os
import sys
import json
import uuid
import zipfile
import shutil
import logging
import re
import random
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime

from flask import Flask, request, jsonify, send_file, render_template
import anthropic

# ── Config ──────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
SCRIPTS    = BASE_DIR / "scripts" / "office"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

# ── Helpers ──────────────────────────────────────────────────────────────────

def xe(text: str) -> str:
    """XML-escape a string."""
    return (text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;"))


def random_para_id() -> str:
    """Valid w14:paraId hex (0x10000000–0x7FFFFFFE)."""
    return f"{random.randint(0x10000000, 0x7FFFFFFE):08X}"


def extract_xml_text(xml: str) -> str:
    """Strip XML tags and normalise whitespace."""
    text = re.sub(r"<[^>]+>", " ", xml)
    return re.sub(r"\s+", " ", text).strip()


def extract_wt_values(xml: str) -> list[str]:
    """Extract all <w:t> text values."""
    return re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml)


def unpack_docx(docx_path: Path, out_dir: Path) -> bool:
    result = subprocess.run(
        [sys.executable, str(SCRIPTS / "unpack.py"), str(docx_path), str(out_dir)],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        log.error("unpack error: %s", result.stderr)
        return False
    return True


def pack_docx(unpacked_dir: Path, out_path: Path, original: Path) -> bool:
    result = subprocess.run(
        [sys.executable, str(SCRIPTS / "pack.py"),
         str(unpacked_dir), str(out_path),
         "--original", str(original)],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        log.error("pack error: %s", result.stderr)
        return False
    return True


def read_file_text(path: Path, max_chars: int = 8000) -> str:
    """Read text from .txt / .csv / basic XML. Returns empty string on failure."""
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            return f.read(max_chars)
    except Exception:
        return ""


def read_docx_text(path: Path) -> str:
    """Quick text extraction from docx via unpack."""
    tmp = Path(tempfile.mkdtemp())
    try:
        if unpack_docx(path, tmp):
            doc_xml = (tmp / "word" / "document.xml").read_text(encoding="utf-8")
            return extract_xml_text(doc_xml)[:6000]
    except Exception as e:
        log.warning("docx read error: %s", e)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return ""


# ── Claude interaction ────────────────────────────────────────────────────────

SYSTEM_PROMPT = """Você é o Assistente Jurídico Especialista do escritório Silva Pinto Advocacia.
Advogado responsável: Dr. Casil da Silva Pinto — OAB/RJ nº 189.781.
Especialidade: Ações anulatórias de atos administrativos em concursos públicos.

SUAS RESPONSABILIDADES:
1. Analisar os documentos do cliente (RG, CPF, procuração, comprovante de residência).
2. Identificar exatamente quais questões serão contestadas (conforme ficha do candidato).
3. Redigir os blocos de Fatos e Fundamentação em linguagem jurídica técnica, formal e persuasiva.
4. Gerar substituições precisas para cada campo do modelo de petição.
5. No capítulo "Da Pontuação": máximo 3 parágrafos, citar a pontuação que o candidato alcançará e destacar que atingirá a nota de corte.
6. No capítulo "Da Probabilidade do Direito": inserir na íntegra o relatório técnico de UMA questão (escolha aleatória).
7. Se o cliente NÃO tiver direito à gratuidade de justiça (conforme ficha), indicar claramente que o capítulo de gratuidade deve ser REMOVIDO.
8. Verificar concordância de gênero em todo o texto.
9. Verificar se a comarca está correta conforme o domicílio do cliente.
10. Sinalizar campos ausentes com [DADO AUSENTE — descrição].

FORMATO DE RESPOSTA — JSON PURO (sem markdown, sem backticks):
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


def call_claude(api_key: str, user_message: str, documents_content: list) -> dict:
    """Call Claude API and parse JSON response."""
    client = anthropic.Anthropic(api_key=api_key)

    content = []
    for doc in documents_content:
        if doc["type"] == "text":
            content.append({"type": "text", "text": doc["text"]})
        elif doc["type"] == "document":
            content.append({
                "type": "document",
                "source": {"type": "base64", "media_type": "application/pdf",
                           "data": doc["data"]}
            })
    content.append({"type": "text", "text": user_message})

    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=8000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": content}]
    )

    raw = "".join(b.text for b in message.content if hasattr(b, "text"))
    # Strip accidental markdown fences
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    return json.loads(raw.strip())


# ── DOCX editing ─────────────────────────────────────────────────────────────

def apply_substitutions(unpacked_dir: Path, data: dict) -> list[str]:
    """Apply all substitutions to word/document.xml. Returns change log."""
    doc_xml_path = unpacked_dir / "word" / "document.xml"
    if not doc_xml_path.exists():
        return ["ERRO: document.xml não encontrado"]

    xml = doc_xml_path.read_text(encoding="utf-8")
    changes = []
    c = data.get("cliente", {})
    p = data.get("processo", {})
    t = data.get("textos", {})

    # ── Simple field substitutions (ordered longest-match first)
    substitutions = [
        # Client data
        ("NOME_COMPLETO_CLIENTE",    c.get("nome_completo", "[DADO AUSENTE — Nome]")),
        ("QUALIFICACAO_CLIENTE",     _build_qualificacao(c)),
        ("CPF_CLIENTE",              c.get("cpf", "[DADO AUSENTE — CPF]")),
        ("RG_CLIENTE",               c.get("rg", "[DADO AUSENTE — RG]")),
        ("EMAIL_CLIENTE",            c.get("email", "[DADO AUSENTE — E-mail]")),
        ("ENDERECO_CLIENTE",         _build_endereco(c)),
        ("CIDADE_CLIENTE",           c.get("cidade", "[DADO AUSENTE — Cidade]")),
        ("UF_CLIENTE",               c.get("uf", "[DADO AUSENTE — UF]")),
        # Process data
        ("COMARCA_JUIZO",            p.get("comarca", "[DADO AUSENTE — Comarca]")),
        ("TIPO_ACAO",                p.get("tipo_acao", "")),
        ("NOME_BANCA",               p.get("banca", "[DADO AUSENTE — Banca]")),
        ("NOME_CONCURSO",            p.get("concurso", "[DADO AUSENTE — Concurso]")),
        ("CARGO_PRETENDIDO",         p.get("cargo", "[DADO AUSENTE — Cargo]")),
        ("PONTUACAO_OBTIDA",         str(p.get("pontuacao_obtida", "[DADO AUSENTE]"))),
        ("PONTUACAO_CORTE",          str(p.get("pontuacao_corte", "[DADO AUSENTE]"))),
        ("PONTUACAO_APOS_ANULACAO",  str(p.get("pontuacao_apos_anulacao", "[DADO AUSENTE]"))),
    ]

    for placeholder, value in substitutions:
        if placeholder in xml:
            xml = xml.replace(placeholder, xe(value))
            changes.append(f"✅ Substituído: {placeholder}")

    # ── Paragraph-level block replacements (fatos, pontuação, etc.)
    xml = _replace_block(xml, "BLOCO_FATOS",                t.get("fatos", ""), changes)
    xml = _replace_block(xml, "BLOCO_PONTUACAO",            t.get("pontuacao", ""), changes)
    xml = _replace_block(xml, "BLOCO_PROBABILIDADE_DIREITO", t.get("probabilidade_direito", ""), changes)
    xml = _replace_block(xml, "BLOCO_FUNDAMENTOS",          t.get("fundamentos_juridicos", ""), changes)

    # ── Insert question summaries
    xml = _insert_questoes(xml, data.get("questoes", []), changes)

    # ── Remove gratuidade chapter if not applicable
    if not p.get("gratuidade", True):
        xml, removed = _remove_gratuidade_chapter(xml)
        if removed:
            changes.append("🗑️ Capítulo de gratuidade de justiça removido")

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
    parts = [c.get("endereco",""), c.get("cidade",""), c.get("uf",""), c.get("cep","")]
    return ", ".join(p for p in parts if p)


def _replace_block(xml: str, marker: str, new_text: str, changes: list) -> str:
    """Replace a single-line XML placeholder with formatted paragraphs."""
    if marker not in xml or not new_text:
        return xml
    para_xml = _text_to_paragraphs(new_text)
    xml = xml.replace(marker, para_xml)
    changes.append(f"✅ Bloco inserido: {marker}")
    return xml


def _text_to_paragraphs(text: str) -> str:
    """Convert plain text (with \n\n) into w:p XML elements."""
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    result = []
    for para in paragraphs:
        pid = random_para_id()
        lines = para.replace("\n", " ")
        result.append(
            f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000">'
            f'<w:pPr><w:ind w:left="0"/><w:jc w:val="both"/></w:pPr>'
            f'<w:r><w:t xml:space="preserve">{xe(lines)}</w:t></w:r>'
            f'</w:p>'
        )
    return "\n".join(result)


def _insert_questoes(xml: str, questoes: list, changes: list) -> str:
    """Insert question blocks into the illegal questions chapter."""
    if not questoes:
        return xml

    markers = [
        "BLOCO_QUESTOES_ILEGAIS",
        "DAS_QUESTOES_ILEGAIS",
        "QUESTOES_ANULAVEIS",
    ]
    marker_found = None
    for m in markers:
        if m in xml:
            marker_found = m
            break

    if not marker_found:
        changes.append("⚠️ Marcador de questões ilegais não encontrado no XML — blocos não inseridos")
        return xml

    blocks = []
    for q in sorted(questoes, key=lambda x: x.get("numero", 0)):
        num   = q.get("numero", "?")
        vicio = q.get("vicio", "VÍCIO NÃO IDENTIFICADO").upper()
        resumo = q.get("resumo_peticao", "")
        enunc  = q.get("enunciado", "")
        alts   = q.get("alternativas", [])

        # Title paragraph
        pid = random_para_id()
        titulo = f"QUESTÃO {num} — {vicio}"
        blocks.append(
            f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000">'
            f'<w:pPr><w:jc w:val="both"/><w:rPr><w:b/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:b/></w:rPr><w:t>{xe(titulo)}</w:t></w:r>'
            f'</w:p>'
        )
        # Enunciado
        if enunc:
            pid = random_para_id()
            blocks.append(
                f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000">'
                f'<w:pPr><w:ind w:left="0"/><w:jc w:val="both"/></w:pPr>'
                f'<w:r><w:t xml:space="preserve">{xe(enunc)}</w:t></w:r>'
                f'</w:p>'
            )
        # Alternativas — one per line
        for alt in alts:
            pid = random_para_id()
            blocks.append(
                f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000">'
                f'<w:pPr><w:ind w:left="720"/><w:jc w:val="both"/></w:pPr>'
                f'<w:r><w:t xml:space="preserve">{xe(alt)}</w:t></w:r>'
                f'</w:p>'
            )
        # Resumo
        if resumo:
            pid = random_para_id()
            blocks.append(
                f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000">'
                f'<w:pPr><w:ind w:left="0"/><w:jc w:val="both"/></w:pPr>'
                f'<w:r><w:t xml:space="preserve">{xe(resumo)}</w:t></w:r>'
                f'</w:p>'
            )
        # spacer
        pid = random_para_id()
        blocks.append(f'<w:p w14:paraId="{pid}" w14:textId="FFFFFFFF" w:rsidR="00000000"><w:pPr/></w:p>')

    xml = xml.replace(marker_found, "\n".join(blocks))
    changes.append(f"✅ {len(questoes)} questão(ões) inserida(s) no capítulo Das Questões Ilegais")
    return xml


def _remove_gratuidade_chapter(xml: str):
    """Remove the gratuidade chapter and its heading from XML."""
    patterns = [
        r'<w:p[^>]*>.*?gratuidade.*?</w:p>\s*(<w:p[^>]*>.*?</w:p>\s*)*?(?=<w:p[^>]*>[^<]*(?:DA PROBABILIDADE|DO DIREITO|DOS PEDIDOS))',
    ]
    for pat in patterns:
        new_xml, n = re.subn(pat, "", xml, flags=re.IGNORECASE | re.DOTALL)
        if n:
            return new_xml, True
    return xml, False


# ── File renaming ─────────────────────────────────────────────────────────────

def rename_files_by_rol(work_dir: Path, rol: list) -> dict:
    """Rename files in work_dir according to the Rol de Documentos mapping."""
    mapping = {}
    files_in_dir = {f.name.lower(): f for f in work_dir.iterdir() if f.is_file()}

    for item in rol:
        num   = item.get("numero", "")
        desc  = item.get("descricao", "")
        orig  = item.get("arquivo_correspondente", "")

        # Sanitise new name
        safe_desc = re.sub(r"[^\w\s\-]", "", desc).strip()
        safe_desc = re.sub(r"\s+", "_", safe_desc)

        src_path = None
        # Try exact match first
        if orig:
            candidate = work_dir / orig
            if candidate.exists():
                src_path = candidate
        # Fallback: fuzzy match
        if not src_path:
            orig_lower = orig.lower()
            for fname_lower, fpath in files_in_dir.items():
                if orig_lower in fname_lower or fname_lower in orig_lower:
                    src_path = fpath
                    break

        if src_path and src_path.exists():
            suffix = src_path.suffix
            new_name = f"{num}. {safe_desc}{suffix}"
            new_path = work_dir / new_name
            src_path.rename(new_path)
            mapping[str(src_path.name)] = new_name
        else:
            mapping[f"[AUSENTE] {orig}"] = f"{num}. {desc} — NÃO ENCONTRADO"

    return mapping


# ── Token budget constants ────────────────────────────────────────────────────
# Claude limit = 200k tokens. We stay well under with these caps:
MAX_CHARS_PER_TEXT_FILE = 4_000   # ~1 000 tokens each
MAX_CHARS_DOCX_MODEL    = 5_000   # modelo is big but we only need placeholders
MAX_CHARS_DOCX_OTHER    = 3_000   # outros docx (fichas, relatórios)
MAX_PDFS_PER_CALL       = 4       # PDFs are expensive; cap at 4 per API call
MAX_CHARS_TEXT_BLOCK    = 80_000  # total text block hard cap (~20k tokens)


def _truncate(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + f"\n... [TRUNCADO — {len(text)-max_chars} caracteres omitidos]"


def _classify_files(all_files: list[Path], extract_dir: Path):
    """Split files into: modelo_docx, identity_pdfs, report_pdfs, text_files, other_docx."""
    modelo      = None
    id_pdfs     = []   # RG, CPF, procuração, comprovante
    report_pdfs = []   # relatórios técnicos de questões
    text_files  = []   # xlsx, csv, txt
    other_docx  = []   # outros docx

    id_keywords     = ("rg", "cpf", "procuracao", "procuração", "identidade",
                       "comprovante", "residencia", "residência", "cnh", "habilitacao")
    report_keywords = ("relatorio", "relatório", "questao", "questão", "tecnico",
                       "técnico", "parecer", "fundamentacao", "fundamentação", "anulacao")
    model_keywords  = ("modelo", "peticao", "petição", "inicial")

    for fpath in all_files:
        if not fpath.is_file():
            continue
        name_low = fpath.name.lower()

        if name_low.endswith(".docx"):
            if any(k in name_low for k in model_keywords):
                modelo = fpath
            else:
                other_docx.append(fpath)
        elif name_low.endswith(".pdf"):
            if any(k in name_low for k in id_keywords):
                id_pdfs.append(fpath)
            elif any(k in name_low for k in report_keywords):
                report_pdfs.append(fpath)
            else:
                # Ambiguous: treat as report (more relevant for Claude)
                report_pdfs.append(fpath)
        elif name_low.endswith((".xlsx", ".xls", ".csv", ".txt")):
            text_files.append(fpath)

    # Fallback modelo
    if not modelo and other_docx:
        modelo = other_docx.pop(0)

    return modelo, id_pdfs, report_pdfs, text_files, other_docx


def _build_text_block(
    file_list, extract_dir, modelo_path,
    text_files, other_docx, form_data
) -> str:
    """Build the text portion of the Claude message within char budget."""
    parts = []

    # Inventory
    parts.append("=== INVENTÁRIO DO ZIP ===")
    for f in file_list:
        parts.append(f"  • {f}")

    # Form data (always included — small)
    parts.append(f"""
=== DADOS FORNECIDOS PELO USUÁRIO ===
Tipo de Ação: {form_data.get('tipo_acao', '')}
Comarca: {form_data.get('comarca', '')}
Resumo dos Fatos: {form_data.get('fatos', '')}
Pedidos Adicionais: {form_data.get('pedidos', '')}
Observações: {form_data.get('obs', '')}
""")

    # Text files (ficha xlsx etc.) — highest priority after form data
    parts.append("=== DOCUMENTOS DE TEXTO ===")
    for fpath in text_files:
        rname = str(fpath.relative_to(extract_dir))
        text  = _truncate(read_file_text(fpath), MAX_CHARS_PER_TEXT_FILE)
        parts.append(f"\n--- {rname} ---\n{text}")

    # Other DOCX (fichas, relatórios em docx)
    for fpath in other_docx:
        rname = str(fpath.relative_to(extract_dir))
        text  = _truncate(read_docx_text(fpath), MAX_CHARS_DOCX_OTHER)
        parts.append(f"\n--- DOCX: {rname} ---\n{text}")

    # Modelo DOCX — only send placeholder skeleton, not full text
    if modelo_path:
        rname = str(modelo_path.relative_to(extract_dir))
        text  = _truncate(read_docx_text(modelo_path), MAX_CHARS_DOCX_MODEL)
        parts.append(f"\n--- MODELO DOCX (esqueleto de placeholders): {rname} ---\n{text}")

    full = "\n".join(parts)
    return _truncate(full, MAX_CHARS_TEXT_BLOCK)


def _load_pdf_b64(fpath: Path) -> str:
    import base64
    with open(fpath, "rb") as f:
        return base64.b64encode(f.read()).decode()


# ── Main processing pipeline ──────────────────────────────────────────────────

def process_zip(zip_path: Path, session_dir: Path, form_data: dict, api_key: str) -> dict:
    """Full pipeline: unzip → analyse → edit docx → repack → zip output."""
    log.info("Starting pipeline for %s", zip_path.name)

    # 1. Extract ZIP
    extract_dir = session_dir / "extracted"
    extract_dir.mkdir()
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(extract_dir)

    all_files = list(extract_dir.rglob("*"))
    file_list = [f.relative_to(extract_dir) for f in all_files if f.is_file()]
    log.info("Files in ZIP: %s", [str(f) for f in file_list])

    # 2. Classify files by type
    docx_model_path, id_pdfs, report_pdfs, text_files, other_docx = \
        _classify_files(all_files, extract_dir)

    if not docx_model_path:
        return {"error": "Nenhum arquivo .docx modelo encontrado no ZIP."}

    log.info("Modelo: %s | ID PDFs: %d | Relatórios: %d | Textos: %d",
             docx_model_path.name, len(id_pdfs), len(report_pdfs), len(text_files))

    # 3. Build text block (budget-controlled)
    text_block = _build_text_block(
        file_list, extract_dir, docx_model_path,
        text_files, other_docx, form_data
    )

    # 4. Select PDFs to send (report PDFs are most important; id PDFs secondary)
    #    Priority: report PDFs first, then id PDFs, capped at MAX_PDFS_PER_CALL
    selected_pdfs = (report_pdfs + id_pdfs)[:MAX_PDFS_PER_CALL]
    log.info("Sending %d PDFs to Claude", len(selected_pdfs))

    # 5. Build Claude message content
    documents_content = [{"type": "text", "text": text_block}]
    for fpath in selected_pdfs:
        rname = str(fpath.relative_to(extract_dir))
        b64   = _load_pdf_b64(fpath)
        documents_content.append({"type": "document", "data": b64, "name": rname})
        documents_content.append({"type": "text",
                                   "text": f"[Arquivo PDF acima: {rname}]"})

    # 3. Call Claude
    log.info("Calling Claude API...")
    data = call_claude(api_key, "Processar caso e retornar JSON:", documents_content)
    log.info("Claude response received")

    # 4. Unpack modelo docx
    unpacked_dir = session_dir / "unpacked"
    if not unpack_docx(docx_model_path, unpacked_dir):
        return {"error": "Falha ao desempacotar o modelo DOCX."}

    # 5. Apply substitutions
    changes = apply_substitutions(unpacked_dir, data)

    # 6. Repack docx
    cliente_nome = data.get("cliente", {}).get("nome_completo", "Cliente")
    safe_nome    = re.sub(r"[^\w\s]", "", cliente_nome).replace(" ", "_")
    tipo_acao    = re.sub(r"[^\w\s]", "", data.get("processo", {}).get("tipo_acao", "Peticao")).replace(" ", "_")
    data_hoje    = datetime.now().strftime("%Y-%m-%d")
    out_docx_name = f"Peticao_{tipo_acao}_{safe_nome}_{data_hoje}.docx"
    out_docx_path = session_dir / out_docx_name

    if not pack_docx(unpacked_dir, out_docx_path, docx_model_path):
        return {"error": "Falha ao reempacotar o DOCX final."}

    log.info("DOCX generated: %s", out_docx_name)

    # 7. Copy supporting files to delivery dir
    delivery_dir = session_dir / "entrega"
    delivery_dir.mkdir()
    shutil.copy(out_docx_path, delivery_dir / out_docx_name)

    # Copy other files for renaming
    for fpath in all_files:
        if fpath.is_file() and fpath != docx_model_path:
            dest = delivery_dir / fpath.name
            if not dest.exists():
                shutil.copy(fpath, dest)

    # 8. Rename by Rol
    rol = data.get("rol_documentos", [])
    rename_map = {}
    if rol:
        rename_map = rename_files_by_rol(delivery_dir, rol)

    # 9. Pack final ZIP
    zip_name = f"Entrega_{safe_nome}_{data_hoje}.zip"
    zip_out   = OUTPUT_DIR / zip_name
    with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as zf:
        for fpath in delivery_dir.iterdir():
            if fpath.is_file():
                zf.write(fpath, fpath.name)

    log.info("Output ZIP: %s", zip_name)

    return {
        "success":        True,
        "zip_filename":   zip_name,
        "docx_filename":  out_docx_name,
        "cliente":        data.get("cliente", {}),
        "processo":       data.get("processo", {}),
        "questoes":       [{"numero": q.get("numero"), "vicio": q.get("vicio")} for q in data.get("questoes", [])],
        "changes":        changes,
        "rename_map":     rename_map,
        "dados_ausentes": data.get("dados_ausentes", []),
        "relatorio":      data.get("relatorio_alteracoes", []),
    }


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/gerar", methods=["POST"])
def gerar():
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
    session_dir.mkdir()

    zip_path = session_dir / "input.zip"
    zip_file.save(str(zip_path))

    form_data = {
        "tipo_acao": request.form.get("tipo_acao", ""),
        "comarca":   request.form.get("comarca", ""),
        "fatos":     request.form.get("fatos", ""),
        "pedidos":   request.form.get("pedidos", ""),
        "obs":       request.form.get("obs", ""),
    }

    try:
        result = process_zip(zip_path, session_dir, form_data, api_key)
    except anthropic.AuthenticationError:
        return jsonify({"error": "Chave de API inválida. Verifique suas credenciais Anthropic."}), 401
    except anthropic.RateLimitError:
        return jsonify({"error": "Limite de uso da API atingido. Aguarde alguns instantes."}), 429
    except json.JSONDecodeError as e:
        return jsonify({"error": f"Resposta inesperada do Claude (não é JSON válido): {e}"}), 500
    except Exception as e:
        log.exception("Pipeline error")
        return jsonify({"error": str(e)}), 500

    if "error" in result:
        return jsonify(result), 500

    # Store session_id in result for download
    result["session_id"] = session_id
    return jsonify(result)


@app.route("/download/<session_id>/<filename>")
def download(session_id, filename):
    # Security: only alphanumeric/dash for session_id
    if not re.match(r"^[\w\-]+$", session_id):
        return "ID inválido", 400
    path = OUTPUT_DIR / filename
    if not path.exists():
        # Fallback: check session upload dir
        path = UPLOAD_DIR / session_id / filename
    if not path.exists():
        return "Arquivo não encontrado", 404
    return send_file(str(path), as_attachment=True, download_name=filename)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
