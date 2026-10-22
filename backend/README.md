# Backend Ley 21.719 (Fase 1)

Estado: esquema de base de datos, reglas de validación y hallazgos del inventario de sistemas y proveedores, con pruebas de referencia. Aún no hay API ni conexión a Supabase.

## Estructura
- `db/migrations/0001_base_sistemas.sql`: tablas, restricciones y políticas de Row Level Security.
- `app/reglas.py`: validación y hallazgos (funciones puras, fuente oficial de las reglas).
- `tests/test_reglas.py`: casos de referencia escritos a mano.

## Ejecutar las pruebas
```
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements-dev.txt
python -m pytest -q tests
```

## Reglas del proyecto
- Ningún dato real ni secreto en el repositorio.
- Las reglas se prueban con casos escritos a mano, nunca con resultados del propio sistema.
- Las pruebas de seguridad por fila se harán contra una base de datos de pruebas en Supabase, con usuarios de prueba.
