"""Reglas de validación y hallazgos del inventario de sistemas y proveedores.

Funciones puras: no acceden a base de datos. El servidor las usa tanto al
guardar como al calcular hallazgos. Son la fuente oficial; el front solo
muestra lo que el servidor entrega.
"""
from dataclasses import dataclass

CATEGORIAS = ('SaaS (nube multiempresa)', 'Licencia instalada', 'Desarrollo a medida',
              'Infraestructura en nube (IaaS o PaaS)', 'Herramienta de productividad',
              'Pagos', 'Videovigilancia o biometría', 'Otro')
ROLES = ('Responsable', 'Encargado', 'Responsable y encargado')
TIPOS_USUARIO = ('Nominativo', 'Genérico', 'Mixto', 'Sin identificar')
UBICACIONES = ('Chile', 'Fuera de Chile', 'Chile y fuera de Chile', 'Por definir')
CONTRATOS = ('Sí', 'No', 'En revisión')


class ErrorValidacion(ValueError):
    pass


@dataclass(frozen=True)
class Sistema:
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
    responsable: str = ''
    cuenta: str = ''


def validar(s: Sistema) -> None:
    """Lanza ErrorValidacion si el registro no cumple las reglas."""
    if not (1 <= len(s.nombre.strip()) <= 150):
        raise ErrorValidacion('El nombre debe tener entre 1 y 150 caracteres.')
    if s.categoria not in CATEGORIAS:
        raise ErrorValidacion(f'Categoría no permitida: {s.categoria}')
    if s.rol not in ROLES:
        raise ErrorValidacion(f'Rol no permitido: {s.rol}')
    if s.tipo_usuario not in TIPOS_USUARIO:
        raise ErrorValidacion(f'Tipo de usuario no permitido: {s.tipo_usuario}')
    if s.ubicacion not in UBICACIONES:
        raise ErrorValidacion(f'Ubicación no permitida: {s.ubicacion}')
    if s.contrato_encargo not in CONTRATOS:
        raise ErrorValidacion(f'Contrato no permitido: {s.contrato_encargo}')
    if s.tipo_usuario == 'Sin identificar' and not (s.responsable.strip() or s.cuenta.strip()):
        raise ErrorValidacion('Un sistema sin usuario identificado requiere responsable o cuenta.')


def es_encargado(s: Sistema) -> bool:
    return s.rol in ('Encargado', 'Responsable y encargado')


def hallazgos(sistemas: list[Sistema]) -> list[dict]:
    """Devuelve los hallazgos del inventario, con su nivel. Ordenados por nivel."""
    if not sistemas:
        return [{'nivel': 'Ok', 'codigo': 'SIN_REGISTROS',
                 'texto': 'Aún no hay sistemas registrados.'}]

    def generico(s):
        return s.tipo_usuario in ('Genérico', 'Mixto')

    reglas = [
        ('Alta', 'ENCARGADO_FUERA_SIN_CONTRATO',
         lambda s: es_encargado(s) and s.ubicacion != 'Chile' and s.contrato_encargo != 'Sí',
         'Encargado que trata datos fuera de Chile sin contrato de encargo.'),
        ('Alta', 'ENCARGADO_SIN_CONTRATO',
         lambda s: es_encargado(s) and s.contrato_encargo != 'Sí' and s.ubicacion == 'Chile',
         'Encargado que trata datos sin contrato de encargo.'),
        ('Alta', 'SENSIBLES_CUENTA_GENERICA',
         lambda s: s.datos_sensibles and generico(s),
         'Datos sensibles con cuenta genérica o mixta.'),
        ('Media', 'CLIENTES_SIN_MFA',
         lambda s: s.datos_clientes and not s.mfa,
         'Datos de clientes sin autenticación multifactor.'),
        ('Media', 'CLIENTES_CUENTA_GENERICA',
         lambda s: s.datos_clientes and generico(s),
         'Datos de clientes con cuenta genérica o mixta.'),
        ('Media', 'PRODUCTIVIDAD_SIN_CONTROL',
         lambda s: s.categoria == 'Herramienta de productividad' and s.datos_clientes and not s.cuenta_individual,
         'Herramienta de productividad con datos de clientes sin cuenta individual.'),
        ('Media', 'PRUEBAS_CON_DATOS_REALES',
         lambda s: s.datos_reales_pruebas,
         'Datos reales en ambientes de prueba.'),
        ('Media', 'SIN_IDENTIFICAR',
         lambda s: s.tipo_usuario == 'Sin identificar',
         'Sistema sin usuario identificado.'),
        ('Baja', 'SIN_RESPONSABLE',
         lambda s: not s.responsable.strip(),
         'Sistema sin responsable asignado.'),
    ]
    orden = {'Alta': 0, 'Media': 1, 'Baja': 2}
    salida = []
    for nivel, codigo, condicion, texto in reglas:
        n = sum(1 for s in sistemas if condicion(s))
        if n:
            salida.append({'nivel': nivel, 'codigo': codigo, 'cantidad': n,
                           'texto': f'{n} sistema(s): {texto}'})
    salida.sort(key=lambda h: orden[h['nivel']])
    if not salida:
        salida.append({'nivel': 'Ok', 'codigo': 'SIN_HALLAZGOS',
                       'texto': 'Sin hallazgos críticos en el inventario de sistemas.'})
    return salida
