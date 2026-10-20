#!/usr/bin/env python3
"""
Genera un respaldo JSON con datos ficticios de la empresa TREZIVA (retail y ecommerce, Chile)
para cargarlo con la opción 'Importar datos' de la aplicación.

Uso:
  python3 simular_treziva.py
Requiere: openpyxl  (pip3 install openpyxl)
Debe estar en la carpeta simulacion/ del repositorio (junto a catalogo/).
Genera: treziva_datos.json (junto a este script)
"""
import json, os, random, datetime
import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(AQUI, '..'))  # carpeta raíz del repositorio ley21719-assessment
CATALOGO = os.path.join(REPO, 'catalogo', 'checklist_ley21719_integrado.xlsx')
ACTIVIDADES = os.path.join(REPO, 'catalogo', 'fuentes', 'actividades_rat.py')
SALIDA = os.path.join(AQUI, 'treziva_datos.json')

random.seed(20261020)  # resultados reproducibles

# ---------- catálogo de actividades (mismo origen que la aplicación) ----------
ns = {}
exec(open(ACTIVIDADES, encoding='utf-8').read(), ns)
CATRAT = ns['catrat']
CAT = {x['actividad']: x for x in CATRAT}

# ---------- controles del catálogo ----------
wb = openpyxl.load_workbook(CATALOGO, read_only=True)
ws = wb['Controles marcos']
CONTROLES = []
for row in ws.iter_rows(min_row=2, values_only=True):
    marco, cid = row[0], row[1]
    if cid:
        CONTROLES.append((marco, str(cid)))
POR_DIM = {}
ws2 = wb['Requisitos Ley']
for row in ws2.iter_rows(min_row=2, values_only=True):
    if row[0]:
        POR_DIM.setdefault(row[0], 0)
        POR_DIM[row[0]] += 1

HOY = datetime.date(2026, 10, 20)
def iso(d): return d.strftime('%Y-%m-%d')

# ---------- RAT: 14 actividades completas (25 campos) ----------
ACTIVIDADES_TREZIVA = [
    ("Gestión de clientes y contratos", "Administración de clientes", "Área comercial"),
    ("Venta en ecommerce y cuenta de cliente", "Ecommerce", "Gerencia Ecommerce"),
    ("Venta en tienda física", "Tiendas", "Gerencia de Tiendas"),
    ("Programa de fidelización", "Fidelización", "Gerencia Comercial"),
    ("Marketing directo y boletines", "Marketing", "Gerencia de Marketing"),
    ("Cookies y analítica web", "Marketing", "Gerencia de Marketing"),
    ("Medios de pago con tarjeta", "Finanzas", "Gerencia de Finanzas"),
    ("Facturación electrónica y documentos tributarios", "Finanzas", "Gerencia de Finanzas"),
    ("Atención al cliente y call center", "Servicio al cliente", "Gerencia de Servicio"),
    ("Logística y despacho a domicilio", "Operaciones", "Gerencia de Operaciones"),
    ("Videovigilancia en tiendas y bodegas", "Seguridad", "Jefatura de Seguridad"),
    ("Selección y contratación de personal", "RRHH", "Gerencia de RRHH"),
    ("Administración de remuneraciones y previsión", "RRHH", "Gerencia de RRHH"),
    ("Gestión de proveedores y contratos", "Compras", "Gerencia de Compras"),
]
CAMPOS_RAT = ['ID', 'Actividad', 'Finalidad', 'Base legal', 'Categorías de datos', 'Sensibles',
  'Titulares', 'Destinatarios', 'Transferencia', 'País', 'Plazo', 'Base legal conservación',
  'Medidas', 'Sistema', 'Responsable', 'Última revisión', 'Fuente', 'Cómo se recolectan',
  'Ubicación', 'Forma de eliminación', 'Encargados', 'Decisiones automatizadas', 'Riesgo',
  'EIPD', 'Derechos ARCOP', 'Estado']
rat = []
for i, (act, sistema, resp) in enumerate(ACTIVIDADES_TREZIVA, start=1):
    it = CAT[act]
    fuente = 'Directamente del titular' if it['tema'] in ('Personas (RRHH)', 'Clientes y ventas', 'Marketing') else 'Sistema interno'
    riesgo = 'Alto' if it['sensibles'] == 'Sí' else ('Medio' if it['transferencia'] == 'Sí' else 'Bajo')
    eipd = 'Sí' if riesgo == 'Alto' else ('Por evaluar' if it['transferencia'] == 'Sí' else 'No')
    ubic = 'Chile y fuera de Chile' if it['transferencia'] == 'Sí' else 'Chile'
    vals = [f'RAT-{i:03d}', it['actividad'], it['finalidad'], it['base'], it['categorias'], it['sensibles'],
            it['titulares'], it['destinatarios'], it['transferencia'], it['pais'] or 'No aplica',
            it['plazo'], it['base'], it['medidas'], sistema, resp, iso(HOY - datetime.timedelta(days=random.randint(5, 60))),
            fuente, 'Por el canal habitual del proceso', ubic, 'Borrado seguro y registro de la eliminación',
            'Proveedores definidos en el contrato', 'No', riesgo, eipd,
            'Acceso, rectificación, supresión, oposición y portabilidad', 'Vigente']
    rat.append({f'f{k}': v for k, v in enumerate(vals)})

# ---------- ARCOP ----------
ARCOP = [
    ('ARC-001', 'C.R.', 'Acceso', HOY - datetime.timedelta(days=40), 'Respondido', 'Entregada copia de los datos'),
    ('ARC-002', 'M.L.', 'Rectificación', HOY - datetime.timedelta(days=28), 'Respondido', 'Dirección corregida'),
    ('ARC-003', 'J.P.', 'Cancelación / Supresión', HOY - datetime.timedelta(days=12), 'En proceso', 'Verificando retención legal'),
    ('ARC-004', 'A.S.', 'Oposición', HOY - datetime.timedelta(days=3), 'Pendiente', 'Recibida por correo'),
]
arcop = []
for cid, ini, tipo, recep, est, obs in ARCOP:
    lim = recep + datetime.timedelta(days=30)
    resp = iso(recep + datetime.timedelta(days=random.randint(5, 20))) if est == 'Respondido' else ''
    arcop.append({'f0': cid, 'f1': ini, 'f2': tipo, 'f3': iso(recep), 'f4': iso(lim), 'f5': resp, 'f6': est, 'f7': obs})

# ---------- Transferencias ----------
TRANSF = [
    ('RAT-006', 'Proveedor de correo masivo', 'EE.UU.', 'Datos de contacto y preferencias', 'Envío de boletines', 'Cláusulas contractuales tipo', 'Vigente', 'Revisión anual pendiente'),
    ('RAT-005', 'Proveedor de analítica web', 'EE.UU.', 'Navegación y dirección IP', 'Medición del sitio', 'Cláusulas contractuales tipo', 'Vigente', ''),
    ('RAT-002', 'Plataforma de pagos', 'Brasil', 'Datos de transacción', 'Procesamiento de pagos', 'Contrato de encargo con cláusulas de privacidad', 'Vigente', ''),
]
transf = [{'f0': a, 'f1': b, 'f2': c, 'f3': d, 'f4': e, 'f5': f, 'f6': g, 'f7': h, 'f8': i} for a, b, c, d, e, f, g, h, i in
          [(t[0], t[1], t[2], t[3], t[4], t[5], t[5], t[6], t[7]) for t in TRANSF]]
for t in transf:  # 'Garantías' en f6 (la aplicación cuenta garantía definida desde f6)
    t['f6'] = t['f5']

# ---------- Entrevistas y registros de controles ----------
COLECCION_CONTROLES = [c for c in CONTROLES]
ENTREVISTAS = [
    ('2026-09-08', ['D1', 'D8'], 'Gerencia General', 'Gerente General', 'Gerencia (rol)'),
    ('2026-09-15', ['D2', 'D5'], 'Gerencia de Tecnología', 'Gerente de TI', 'Jefatura TI (rol)'),
    ('2026-09-22', ['D3', 'D18'], 'Gerencia de Marketing', 'Gerente de Marketing', 'Marketing (rol)'),
    ('2026-10-01', ['D4', 'D17'], 'Servicio al cliente y RRHH', 'Jefatura de Servicio', 'Servicio (rol)'),
    ('2026-10-08', ['D9', 'D10'], 'Operaciones y Compras', 'Gerente de Operaciones', 'Operaciones (rol)'),
    ('2026-10-15', ['D14', 'D15'], 'Seguridad y Tiendas', 'Jefe de Seguridad', 'Seguridad (rol)'),
]
RESP_BASE = [
    'Existe un procedimiento escrito, con responsables asignados.',
    'Se aplica en la operación diaria, con registros parciales.',
    'La práctica varía entre tiendas y el ecommerce.',
    'Hay evidencia de la última ejecución, pero sin revisión periódica.',
    'Se revisa cada trimestre y se registran las excepciones.',
]
EVID = ['Sí', 'Parcial', 'No']
PUNT = ['2', '3', '3', '4', '4', '4', '5', 'N/A']

entrevistas, registros = [], []
for n, (fecha, dims, area, cargo, entrevistado) in enumerate(ENTREVISTAS, start=1):
    sid = f'ENT-TRZ{n:02d}'
    respuestas = {}
    for dim in dims:
        for q in range(4):
            respuestas[f'{dim}|{q}'] = {'resp': random.choice(RESP_BASE), 'evid': random.choice(EVID), 'obs': ''}
    controles = {}
    cand = [c for m, c in CONTROLES]
    random.shuffle(cand)
    for cid in cand[:7]:
        p = random.choice(PUNT)
        ver = 'Sí' if p != '2' and random.random() > 0.35 else 'No'
        controles[cid] = {'afirma': 'Afirmación registrada en la entrevista.', 'verificada': ver, 'puntaje': p, 'obs': ''}
        registros.append({'id': f'REG-{sid}-{cid}', 'sesion': sid, 'fecha': fecha, 'dimension': ', '.join(dims),
                          'area': area, 'cargo': cargo, 'control': cid, 'afirma': controles[cid]['afirma'],
                          'evidencia': '', 'verificada': ver, 'puntaje': p, 'obs': ''})
    entrevistas.append({'id': sid, 'fecha': fecha, 'dims': dims, 'dimension': ', '.join(dims), 'area': area,
                        'cargo': cargo, 'entrevistado': entrevistado, 'respuestas': respuestas,
                        'controles': controles, 'ctrlFoco': ''})

# ---------- Historial de evaluaciones (registros de meses anteriores) ----------
historial = [
    {'fecha': '2026-07-15', 'total': 2.31, 'meta': 4.5, 'nivel': 'Inicial', 'cobertura': 0.31, 'incluidos': 198, 'evaluados': 61, 'dims': {}},
    {'fecha': '2026-08-20', 'total': 2.84, 'meta': 4.5, 'nivel': 'Inicial', 'cobertura': 0.44, 'incluidos': 198, 'evaluados': 87, 'dims': {}},
    {'fecha': '2026-09-22', 'total': 3.12, 'meta': 4.5, 'nivel': 'En desarrollo', 'cobertura': 0.58, 'incluidos': 198, 'evaluados': 115, 'dims': {}},
]

# ---------- Metadatos de la organización ----------
meta = {
    'org': 'Treziva SpA (empresa ficticia)',
    'responsable': 'Gerencia General',
    'fecha': '2026-10-20',
    'evaluador': 'Equipo de cumplimiento',
    'acepta': True,
    'aceptante': 'Gerente de Cumplimiento',
    'fechaAcepta': '2026-10-01',
    'criterios': {'sensibles': True, 'automatizadas': False, 'terceros': True, 'vigilancia': True},
    'sugerida': 4.5,
    'meta': 4.5,
    'ajustada': False,
    'justificacion': '',
}

estado = {
    'registros': registros,
    'entrevistas': entrevistas,
    'historial': historial,
    'meta': meta,
    'tables': {'rat': rat, 'arcop': arcop, 'transf': transf},
}

with open(SALIDA, 'w', encoding='utf-8') as f:
    json.dump(estado, f, ensure_ascii=False, indent=2)

print('Generado:', SALIDA)
print(f'  Actividades RAT: {len(rat)} | ARCOP: {len(arcop)} | Transferencias: {len(transf)}')
print(f'  Entrevistas: {len(entrevistas)} | Registros de controles: {len(registros)} | Historial: {len(historial)}')
print()
print('Pasos en la aplicación:')
print('  1. Solicitar acceso con los datos de Treziva (correo ficticio, RUT ficticio).')
print('  2. En Administración, con el PIN, autorizar la empresa Treziva.')
print('  3. Ingresar a Treziva.')
print('  4. Ir a Más > Importar datos (JSON) y seleccionar treziva_datos.json.')
print('  Importante: la importación reemplaza los datos de la empresa activa.')
