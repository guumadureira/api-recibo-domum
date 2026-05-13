import os
import uuid
import subprocess
from typing import Optional

from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, ConfigDict
from docxtpl import DocxTemplate


app = FastAPI(
    title="API Geradora de Recibos Domum Engenharia",
    version="1.4.0"
)

API_SECRET_TOKEN = os.getenv("API_SECRET_TOKEN", "")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "modelo_recibo_domum_automacao.docx")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

os.makedirs(OUTPUT_DIR, exist_ok=True)

app.mount("/outputs", StaticFiles(directory=OUTPUT_DIR), name="outputs")


class ReciboRequest(BaseModel):
    model_config = ConfigDict(extra="ignore")

    numero_recibo: str
    ano_recibo: Optional[str] = None

    nome_cliente: str
    cpf_cnpj_cliente: str
    endereco_cliente: Optional[str] = "não informado"

    valor: str
    valor_extenso: str

    descricao_pagamento: str
    vinculo_documento: Optional[str] = "não informado"
    descricao_obra_servico: str
    endereco_obra: Optional[str] = "não informado"

    forma_pagamento: str
    data_pagamento: str

    cidade_uf: Optional[str] = "Maringá-PR"
    dia: Optional[str] = ""
    mes_extenso: Optional[str] = ""
    ano: Optional[str] = ""

    responsavel: str


@app.get("/")
def health_check():
    return {
        "status": "online",
        "service": "API Geradora de Recibos Domum Engenharia",
        "output": "pdf_link",
        "version": "1.4.0"
    }


@app.post("/gerar-recibo")
def gerar_recibo(
    request: Request,
    dados: ReciboRequest,
    authorization: str = Header(default="")
):
    if not API_SECRET_TOKEN:
        raise HTTPException(
            status_code=500,
            detail="API_SECRET_TOKEN não configurado no servidor."
        )

    expected_token = f"Bearer {API_SECRET_TOKEN}"

    if authorization != expected_token:
        raise HTTPException(status_code=401, detail="Token inválido.")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(
            status_code=500,
            detail=f"Modelo DOCX não encontrado no servidor em: {TEMPLATE_PATH}"
        )

    doc = DocxTemplate(TEMPLATE_PATH)

    contexto = {
        "NUMERO_RECIBO": dados.numero_recibo,
        "ANO_RECIBO": dados.ano_recibo or dados.ano or "",

        "CLIENTE_NOME": dados.nome_cliente,
        "CLIENTE_CPF_CNPJ": dados.cpf_cnpj_cliente,
        "CLIENTE_ENDERECO": dados.endereco_cliente or "não informado",

        "VALOR_NUMERICO": dados.valor,
        "VALOR_EXTENSO": dados.valor_extenso,

        "DESCRICAO_PAGAMENTO": dados.descricao_pagamento,
        "REFERENCIA_PROPOSTA_CONTRATO": dados.vinculo_documento or "não informado",
        "DESCRICAO_OBRA_SERVICO": dados.descricao_obra_servico,
        "ENDERECO_OBRA": dados.endereco_obra or "não informado",

        "FORMA_PAGAMENTO": dados.forma_pagamento,
        "DATA_PAGAMENTO": dados.data_pagamento,

        "CIDADE_UF": dados.cidade_uf or "Maringá-PR",
        "DIA": dados.dia or "",
        "MES_EXTENSO": dados.mes_extenso or "",
        "ANO": dados.ano or "",

        "RESPONSAVEL": dados.responsavel
    }

    docx_filename = f"recibo_domum_{uuid.uuid4().hex}.docx"
    docx_path = os.path.join(OUTPUT_DIR, docx_filename)

    try:
        doc.render(contexto)
        doc.save(docx_path)
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Erro ao preencher o modelo DOCX: {str(e)}"
        )

    try:
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                OUTPUT_DIR,
                docx_path
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
    except subprocess.CalledProcessError as e:
        raise HTTPException(
            status_code=500,
            detail=f"Erro ao converter DOCX para PDF: {e.stderr}"
        )

    pdf_filename = docx_filename.replace(".docx", ".pdf")
    pdf_path = os.path.join(OUTPUT_DIR, pdf_filename)

    if not os.path.exists(pdf_path):
        raise HTTPException(
            status_code=500,
            detail=f"PDF não foi gerado corretamente. Caminho esperado: {pdf_path}"
        )

    base_url = str(request.base_url).rstrip("/")
    download_url = f"{base_url}/outputs/{pdf_filename}"

    return JSONResponse(
        content={
            "status": "success",
            "message": "Recibo gerado com sucesso em PDF.",
            "arquivo": pdf_filename,
            "download_url": download_url
        }
    )