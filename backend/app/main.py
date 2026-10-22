# app/main.py
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from app.reglas import Sistema, ErrorValidacion, validar

app = FastAPI(title="Assessment Ley 21.719")


class SistemaEntrada(BaseModel):
    nombre: str
    categoria: str
    rol: str
    tipo_usuario: str
    datos_clientes: bool
    datos_sensibles: bool
    ubicacion: str
    contrato_encargo: str
    mfa: bool
    cifrado: bool
    cuenta_individual: bool
    datos_reales_pruebas: bool
    responsable: str = ""
    cuenta: str = ""


@app.get("/salud")
def salud():
    return {"estado": "ok"}


@app.post("/sistemas", status_code=201)
def crear_sistema(entrada: SistemaEntrada):
    try:
        sistema = Sistema(**entrada.model_dump())
        validar(sistema)
    except ErrorValidacion as e:
        raise HTTPException(status_code=422, detail=str(e))
    # El guardado en Supabase se agrega cuando esté conectado
    return {"estado": "validado"}
