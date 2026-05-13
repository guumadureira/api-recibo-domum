import os
import uuid
import subprocess
from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from docxtpl import DocxTemplate

app = FastAPI(
    title="API Geradora de Recibos Domum Engenharia",
    version="1.2.0"
)

API_SECRET_TOKEN = os.getenv("API_SECRET_TOKEN", "")
TEMPLATE_PATH = "templates/modelo_recibo_domum_automacao.docx"
OUTPUT_DIR = "outputs"

os.makedirs(OUTPUT_DIR, exist_ok=True)

app.mount("/outputs", StaticFiles(directory=OUTPUT_DIR), name="outputs")


class ReciboRequest(BaseModel):
    numero_recibo: str
    nome_cliente: str
    cpf_cnpj_cliente: str
    endereco_cliente: str
    valor: str
    valor_extenso: str
    descricao_pagamento: str
    vinculo_documento: str
    descricao_obra_servico: str
    endereco_obra: str
    forma_pagamento: str
    data_pagamento: str
    cidade_data: str
    responsavel: str


@app.get("/")
def health_check():
    return {
        "status": "online",
        "service": "API Geradora de Recibos Domum Engenharia",
        "output": "pdf_link"
    }


@app.post("/gerar-recibo")
def gerar_recibo(
    request: Request,
    dados: ReciboRequest,
    authorization: str = Header(default="")
):
    if not API_SECRET_TOKEN:
        raise HTTPException(status_code=500, detail="API_SECRET_TOKEN não configurado no servidor.")

    expected_token = f"Bearer {API_SECRET_TOKEN}"

    if authorization != expected_token:
        raise HTTPException(status_code=401, detail="Token inválido.")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail="Modelo DOCX não encontrado no servidor.")

    doc = DocxTemplate(TEMPLATE_PATH)

    contexto = {
        "NUMERO_RECIBO": dados.numero_recibo,
        "NOME_CLIENTE": dados.nome_cliente,
        "CPF_CNPJ_CLIENTE": dados.cpf_cnpj_cliente,
        "ENDERECO_CLIENTE": dados.endereco_cliente,
        "VALOR": dados.valor,
        "VALOR_EXTENSO": dados.valor_extenso,
        "DESCRICAO_PAGAMENTO": dados.descricao_pagamento,
        "VINCULO_DOCUMENTO": dados.vinculo_documento,
        "DESCRICAO_OBRA_SERVICO": dados.descricao_obra_servico,
        "ENDERECO_OBRA": dados.endereco_obra,
        "FORMA_PAGAMENTO": dados.forma_pagamento,
        "DATA_PAGAMENTO": dados.data_pagamento,
        "CIDADE_DATA": dados.cidade_data,
        "RESPONSAVEL": dados.responsavel
    }

    docx_filename = f"recibo_domum_{uuid.uuid4().hex}.docx"
    docx_path = os.path.join(OUTPUT_DIR, docx_filename)

    doc.render(contexto)
    doc.save(docx_path)

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
        raise HTTPException(status_code=500, detail="PDF não foi gerado corretamente.")

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