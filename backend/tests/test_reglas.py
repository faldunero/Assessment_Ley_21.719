"""Casos de referencia escritos a mano. Cada resultado esperado se deduce
de la regla, no se obtiene ejecutando el sistema."""
import pytest
from app.reglas import Sistema, ErrorValidacion, validar, hallazgos

def base(**kw):
    d = dict(nombre='CRM', categoria='SaaS (nube multiempresa)', rol='Responsable',
             tipo_usuario='Nominativo', datos_clientes=False, datos_sensibles=False,
             ubicacion='Chile', contrato_encargo='No', mfa=True, cifrado=True,
             cuenta_individual=True, datos_reales_pruebas=False,
             responsable='Gerente TI', cuenta='individual')
    d.update(kw)
    return Sistema(**d)

def codigos(lista):
    return {h['codigo'] for h in hallazgos(lista)}

def test_sin_registros():
    assert codigos([]) == {'SIN_REGISTROS'}

def test_sistema_limpio_no_genera_hallazgos():
    assert codigos([base()]) == {'SIN_HALLAZGOS'}

def test_encargado_fuera_de_chile_sin_contrato_es_alta():
    s = base(rol='Encargado', ubicacion='Fuera de Chile', contrato_encargo='No')
    assert 'ENCARGADO_FUERA_SIN_CONTRATO' in codigos([s])
    nivel = [h['nivel'] for h in hallazgos([s]) if h['codigo'] == 'ENCARGADO_FUERA_SIN_CONTRATO'][0]
    assert nivel == 'Alta'

def test_encargado_con_contrato_no_genera_alta():
    s = base(rol='Encargado', ubicacion='Chile', contrato_encargo='Sí')
    assert not ({'ENCARGADO_SIN_CONTRATO', 'ENCARGADO_FUERA_SIN_CONTRATO'} & codigos([s]))

def test_sensibles_con_cuenta_generica_es_alta_y_va_primero():
    s = base(datos_sensibles=True, tipo_usuario='Genérico')
    r = hallazgos([s])
    assert r[0]['codigo'] == 'SENSIBLES_CUENTA_GENERICA' and r[0]['nivel'] == 'Alta'

def test_clientes_sin_mfa_es_media():
    assert 'CLIENTES_SIN_MFA' in codigos([base(datos_clientes=True, mfa=False)])

def test_productividad_con_clientes_sin_cuenta_individual():
    s = base(categoria='Herramienta de productividad', datos_clientes=True, cuenta_individual=False)
    assert 'PRODUCTIVIDAD_SIN_CONTROL' in codigos([s])

def test_pruebas_con_datos_reales():
    assert 'PRUEBAS_CON_DATOS_REALES' in codigos([base(datos_reales_pruebas=True)])

def test_sin_responsable_es_baja():
    s = base(responsable='')
    r = [h for h in hallazgos([s]) if h['codigo'] == 'SIN_RESPONSABLE']
    assert r and r[0]['nivel'] == 'Baja'

def test_conteo_por_hallazgo():
    lista = [base(datos_clientes=True, mfa=False), base(datos_clientes=True, mfa=False), base()]
    h = [x for x in hallazgos(lista) if x['codigo'] == 'CLIENTES_SIN_MFA'][0]
    assert h['cantidad'] == 2

def test_validar_rechaza_categoria_fuera_de_lista():
    with pytest.raises(ErrorValidacion):
        validar(base(categoria='Otra cosa'))

def test_validar_rechaza_sin_identificar_sin_responsable_ni_cuenta():
    with pytest.raises(ErrorValidacion):
        validar(base(tipo_usuario='Sin identificar', responsable='', cuenta=''))

def test_validar_acepta_registro_correcto():
    validar(base())
