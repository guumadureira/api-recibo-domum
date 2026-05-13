import os
import uuid
from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docxtpl import DocxTemplate

app = FastAPI(
    title="API Geradora de Recibos Domum Engenharia",
    version="1.0.0"
)

API_SECRET_TOKEN = os.getenv("API_SECRET_TOKEN", "")
TEMPLATE_PATH = "templates/modelo_recibo_domum_automacao.docx"
OUTPUT_DIR = "outputs"

os.makedirs(OUTPUT_DIR, exist_ok=True)


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
        "service": "API Geradora de Recibos Domum Engenharia"
    }


@app.post("/gerar-recibo")
def gerar_recibo(
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

    doc.render(contexto)

    filename = f"recibo_domum_{uuid.uuid4().hex}.docx"
    output_path = os.path.join(OUTPUT_DIR, filename)

    doc.save(output_path)

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename
    )