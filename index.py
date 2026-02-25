import base64
import binascii
import os
import tempfile
from typing import Any
from uuid import uuid4

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
from mailmerge import MailMerge

app = FastAPI(title="DOCX MailMerge API")
TMP_DIR = tempfile.gettempdir()


class GenerateRequest(BaseModel):
    file_base64: str = Field(..., description="Arquivo DOCX em base64")
    merge_data: dict[str, Any] = Field(default_factory=dict, description="Campos para merge")
    uuid_field: str = Field(default="uuid", description="Nome do campo que recebera o UUID")
    return_file_base64: bool = Field(default=False, description="Retorna o arquivo final em base64")


class TemplateFieldsRequest(BaseModel):
    file_base64: str = Field(..., description="Arquivo DOCX em base64")


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
