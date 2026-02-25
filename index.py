import base64
import binascii
import io
import json
import os
import re
import tempfile
import zipfile
from typing import Any
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from mailmerge import MailMerge

app = FastAPI(title="DOCX MailMerge API")
TMP_DIR = tempfile.gettempdir()
PLACEHOLDER_REGEX = re.compile(r"\{\{\s*#([A-Za-z0-9_]+)\s*\}\}")


class GenerateRequest(BaseModel):
    file_base64: str = Field(..., description="Arquivo DOCX em base64")
    merge_data: dict[str, Any] = Field(default_factory=dict, description="Campos para merge")
    uuid_field: str = Field(default="uuid", description="Nome do campo que recebera o UUID")
    return_file_base64: bool = Field(default=False, description="Retorna o arquivo final em base64")


class TemplateFieldsRequest(BaseModel):
    file_base64: str = Field(..., description="Arquivo DOCX em base64")


def _convert_placeholders_to_fields(xml_text: str) -> tuple[str, list[str]]:
    fields: set[str] = set()

    def replace_block(match: re.Match[str]) -> str:
        raw_block = match.group(0)
        text_only = re.sub(r"<[^>]+>", "", raw_block)
        token_match = PLACEHOLDER_REGEX.fullmatch(text_only)
        if not token_match:
            return raw_block

        field_name = token_match.group(1)
        fields.add(field_name)
        return f"<w:fldSimple w:instr=\" MERGEFIELD  {field_name}  \\\\* MERGEFORMAT \"/>"

    updated_xml = re.sub(r"\{\{(?:[^{}]|<[^>]+>)*\}\}", replace_block, xml_text)
    return updated_xml, sorted(fields)


@app.get("/")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/generate")
def generate_docx(payload: GenerateRequest) -> dict[str, Any]:
    request_uuid = str(uuid4())
    input_path = os.path.join(TMP_DIR, f"{request_uuid}_template.docx")
    output_path = os.path.join(TMP_DIR, f"{request_uuid}_output.docx")

    try:
        raw_file = base64.b64decode(payload.file_base64)
    except (binascii.Error, ValueError) as exc:
        raise HTTPException(status_code=400, detail="file_base64 invalido") from exc

    try:
        with open(input_path, "wb") as f:
            f.write(raw_file)

        merge_data = {k: "" if v is None else str(v) for k, v in payload.merge_data.items()}
        merge_data[payload.uuid_field] = request_uuid

        with MailMerge(input_path) as doc:
            doc.merge(**merge_data)
            doc.write(output_path)

        response: dict[str, Any] = {
            "uuid": request_uuid,
            "output_path": output_path,
        }

        if payload.return_file_base64:
            with open(output_path, "rb") as f:
                response["file_base64"] = base64.b64encode(f.read()).decode("utf-8")

        return response
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"erro ao gerar docx: {exc}") from exc


@app.post("/generate-file")
def generate_docx_file(
    file: UploadFile = File(...),
    merge_data_json: str = Form(..., description="JSON com os campos para merge"),
) -> FileResponse:
    request_uuid = str(uuid4())
    input_path = os.path.join(TMP_DIR, f"{request_uuid}_template.docx")
    output_path = os.path.join(TMP_DIR, f"{request_uuid}_output.docx")

    try:
        merge_data_raw = json.loads(merge_data_json)
        if not isinstance(merge_data_raw, dict):
            raise ValueError("merge_data_json deve ser um objeto JSON")
    except (json.JSONDecodeError, ValueError) as exc:
        raise HTTPException(status_code=400, detail="merge_data_json invalido") from exc

    try:
        with open(input_path, "wb") as f:
            f.write(file.file.read())

        merge_data = {k: "" if v is None else str(v) for k, v in merge_data_raw.items()}
        merge_data["uuid"] = request_uuid

        with MailMerge(input_path) as doc:
            doc.merge(**merge_data)
            doc.write(output_path)

        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=f"{request_uuid}.docx",
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"erro ao gerar docx: {exc}") from exc


@app.post("/fields")
def get_template_fields(payload: TemplateFieldsRequest) -> dict[str, Any]:
    request_uuid = str(uuid4())
    input_path = os.path.join(TMP_DIR, f"{request_uuid}_template.docx")

    try:
        raw_file = base64.b64decode(payload.file_base64)
    except (binascii.Error, ValueError) as exc:
        raise HTTPException(status_code=400, detail="file_base64 invalido") from exc

    try:
        with open(input_path, "wb") as f:
            f.write(raw_file)

        with MailMerge(input_path) as doc:
            fields = sorted(doc.get_merge_fields())

        return {
            "fields": fields,
            "count": len(fields),
        }
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"erro ao ler campos: {exc}") from exc


@app.post("/normalize-fields")
def normalize_fields(
    file: UploadFile = File(...),
    merge_data_json: str = Form(..., description="JSON com os campos para merge"),
    uuid_field: str = Form("uuid", description="Nome do campo que recebera o UUID"),
) -> FileResponse:
    request_uuid = str(uuid4())
    normalized_path = os.path.join(TMP_DIR, f"{request_uuid}_normalized.docx")
    output_path = os.path.join(TMP_DIR, f"{request_uuid}_output.docx")

    try:
        merge_data_raw = json.loads(merge_data_json)
        if not isinstance(merge_data_raw, dict):
            raise ValueError("merge_data_json deve ser um objeto JSON")
    except (json.JSONDecodeError, ValueError) as exc:
        raise HTTPException(status_code=400, detail="merge_data_json invalido") from exc

    try:
        raw_docx = file.file.read()
        with zipfile.ZipFile(io.BytesIO(raw_docx), "r") as zip_in:
            file_map = {name: zip_in.read(name) for name in zip_in.namelist()}

        if "word/document.xml" not in file_map:
            raise HTTPException(status_code=400, detail="docx invalido")

        document_xml = file_map["word/document.xml"].decode("utf-8")
        updated_xml, _ = _convert_placeholders_to_fields(document_xml)
        file_map["word/document.xml"] = updated_xml.encode("utf-8")

        with zipfile.ZipFile(normalized_path, "w", compression=zipfile.ZIP_DEFLATED) as zip_out:
            for name, content in file_map.items():
                zip_out.writestr(name, content)

        merge_data = {k: "" if v is None else str(v) for k, v in merge_data_raw.items()}
        merge_data[uuid_field] = request_uuid

        with MailMerge(normalized_path) as doc:
            doc.merge(**merge_data)
            doc.write(output_path)

        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=f"{request_uuid}.docx",
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"erro ao normalizar campos: {exc}") from exc
