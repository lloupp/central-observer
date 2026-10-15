"""Integração do gerador existente com a Central. Dados e arquivos só no computador local."""
from __future__ import annotations
import datetime as dt
import importlib.util
import io
import os
import re
import tempfile
import uuid
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

BASE = Path(__file__).resolve().parent
MODEL_PATH = BASE / "data" / "modelo_certificado.docx"
OUTPUT_DIR = BASE / "data" / "certificados"
REQUIRED_TAGS = {"NOME", "TEXTO", "ESP", "DE", "ATE", "CH", "DIA_EMISSAO", "MES_EMISSAO", "ANO_EMISSAO"}
MODALIDADES = (
    "Observership", "Observership (Experiência Multiprofissional)",
    "Residente", "Doutorando(a)", "Mini Fellowship",
)


def check_template(raw: bytes) -> set[str]:
    if len(raw) > 5_000_000:
        raise ValueError("Modelo Word acima do limite de 5 MB.")
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            if "word/document.xml" not in z.namelist():
                raise ValueError("Arquivo não é um modelo Word .docx válido.")
            info = z.getinfo("word/document.xml")
            if info.file_size > 3_000_000:
                raise ValueError("Modelo Word contém conteúdo excessivo.")
            xml = z.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError, RuntimeError, ET.ParseError) as e:
        raise ValueError("Arquivo inválido: selecione um modelo .docx.") from e
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as e:
        raise ValueError("Estrutura do modelo Word inválida.") from e
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    paragraphs = ["".join(t.text or "" for t in p.iter(ns + "t")) for p in root.iter(ns + "p")]
    text = "\n".join(paragraphs)
    tags = set(re.findall(r"\{\{\s*([A-Za-z_]+)\s*\}\}", text))
    missing = REQUIRED_TAGS - tags
    if missing:
        raise ValueError("O modelo não contém os campos obrigatórios: " + ", ".join(sorted(missing)))
    return tags


def save_template(raw: bytes) -> None:
    check_template(raw)
    MODEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    # Troca atômica: nunca apagar o modelo atual antes de validar o novo.
    temp = MODEL_PATH.with_name(".modelo_" + uuid.uuid4().hex + ".docx")
    try:
        temp.write_bytes(raw)
        os.replace(temp, MODEL_PATH)
    finally:
        temp.unlink(missing_ok=True)


def template_status() -> dict:
    exists = MODEL_PATH.is_file()
    deps = importlib.util.find_spec("docxtpl") is not None and importlib.util.find_spec("openpyxl") is not None
    return {"modelo_configurado": exists, "dependencias_instaladas": deps, "modalidades": list(MODALIDADES)}


def validate_record(record: dict) -> dict:
    if not MODEL_PATH.is_file():
        raise ValueError("Primeiro carregue o modelo original .docx na seção Certificados.")
    if not template_status()["dependencias_instaladas"]:
        raise ValueError("Bibliotecas de certificado ausentes. Execute INSTALAR_CERTIFICADOS.bat.")
    if record.get("status") == "cancelado" or record.get("situacao_aprovacao") != "aprovado":
        raise ValueError("O estágio precisa estar ativo e aprovado.")
    if record.get("status_certificado") == "enviado":
        raise ValueError("Certificado já marcado como enviado. Verifique o histórico antes de gerar outro.")
    if record.get("status_formulario") != "recebido" or record.get("status_cracha") != "devolvido":
        raise ValueError("Confirme recebimento do formulário e devolução do crachá antes da emissão.")
    modalidade = record.get("modalidade") or "Observership"
    if modalidade not in MODALIDADES:
        raise ValueError("Selecione uma modalidade de certificado válida.")
    try:
        inicio = dt.date.fromisoformat(record["inicio"])
        fim = dt.date.fromisoformat(record["fim"])
    except (ValueError, TypeError, KeyError) as e:
        raise ValueError("Informe datas válidas de início e término.") from e
    if fim < inicio:
        raise ValueError("Data de término anterior ao início do estágio.")
    carga = str(record.get("carga_horaria") or "").strip()
    if not re.fullmatch(r"\d{1,4}(?:[.,]\d{1,2})?", carga):
        raise ValueError("Informe a carga horária comprovada (ex.: 40 ou 37,5).")
    ch = float(carga.replace(",", "."))
    if not 0 < ch <= 2000:
        raise ValueError("Carga horária inválida.")
    if modalidade.startswith("Observership") and ch > 40:
        raise ValueError("Observership não pode ultrapassar 40 horas no período; confira o formulário.")
    if not record.get("nome") or not record.get("especialidade") or not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", record.get("email", "")):
        raise ValueError("Informe nome, especialidade e e-mail válido antes da emissão.")
    return {"nome": record["nome"], "email": record["email"], "modalidade": modalidade,
            "especialidade": record["especialidade"], "de": inicio, "ate": fim, "ch": ch}


def generate_record(record: dict) -> tuple[str, bytes]:
    person = validate_record(record)
    # Reutiliza o motor recebido: mesmas modalidades, correção de acentos,
    # data de emissão e texto original; não altera planilhas externas.
    from certificados.gerar_certificados_ensino import gerar_certificado, validar_pessoa
    problems = validar_pessoa(person)
    if problems:
        raise ValueError("; ".join(problems))
    with tempfile.TemporaryDirectory(prefix="observership_cert_") as temporary:
        filename = gerar_certificado(person, MODEL_PATH, Path(temporary))
        result = (Path(temporary) / filename).read_bytes()
    if not zipfile.is_zipfile(io.BytesIO(result)):
        raise ValueError("O certificado resultante não é um .docx válido.")
    return filename, result


def store_certificate(record_id: int, raw: bytes) -> str:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    filename = f"certificado_{record_id}_{uuid.uuid4().hex}.docx"
    path = OUTPUT_DIR / filename
    with path.open("xb") as f:
        f.write(raw)
    return filename


def stored_certificate(filename: str) -> bytes:
    if not re.fullmatch(r"certificado_\d+_[a-f0-9]{32}\.docx", filename or ""):
        raise ValueError("Nome de certificado inválido.")
    path = OUTPUT_DIR / filename
    if not path.is_file():
        raise ValueError("Arquivo local não encontrado; gere novamente a partir do cadastro.")
    return path.read_bytes()
