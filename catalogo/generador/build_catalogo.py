import re, datetime, json, shutil, subprocess, os
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

ns={}; exec(open('/tmp/build.py',encoding='utf-8').read().split('\nwb=Workbook()')[0], ns)
DIMS, ITEMS = ns['DIMS'], ns['ITEMS']

NAVY="1F3864"; LIGHT="D9E2F3"; INPUT="FFF2CC"
thin=Side(style="thin",color="BFBFBF"); B=Border(left=thin,right=thin,top=thin,bottom=thin)
HF=PatternFill("solid",fgColor=NAVY); H=Font(bold=True,color="FFFFFF")
WR=Alignment(wrap_text=True,vertical="top"); CT=Alignment(horizontal="center",vertical="center",wrap_text=True)
INF=PatternFill("solid",fgColor=INPUT)

def header(ws,row,vals,widths=None):
    for i,v in enumerate(vals,1):
        c=ws.cell(row,i,v); c.font=H; c.fill=HF; c.alignment=CT; c.border=B
    if widths:
        for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w

# ---------- catálogo de controles de marcos ----------
iso=json.load(open('/tmp/iso.json'))
ctrls=[("ISO 27001:2022",cid,name,tema) for cid,name,tema,tipo in iso]
nist={"GOVERN":["GV.OC-01","GV.OC-02","GV.OC-03","GV.OC-04","GV.OC-05","GV.RR-01","GV.RR-02","GV.RR-03","GV.RR-04","GV.PO-01","GV.PO-02"],
"IDENTIFY":["ID.AM-01","ID.AM-02","ID.AM-03","ID.AM-05","ID.AM-08","ID.RA-01","ID.RA-08"],
"PROTECT":["PR.AA-01","PR.AA-02","PR.AA-03","PR.AA-04","PR.AA-05","PR.DS-11","PR.PS-01","PR.PS-02","PR.PS-03","PR.PS-05"],
"RESPOND":["RS.MA-01","RS.MA-02","RS.MA-03","RS.MA-04","RS.MA-05","RS.AN-03","RS.AN-06","RS.AN-07","RS.AN-08","RS.CO-02","RS.CO-03","RS.MI-01","RS.MI-02"],
"RECOVER":["RC.RP-03"]}
for f,ids in nist.items():
    for i in ids: ctrls.append(("NIST CSF 2.0","NIST "+i,"Subcategoría NIST CSF 2.0",f))
for k in ["CIS 1","CIS 2","CIS 4","CIS 5","CIS 6","CIS 7","CIS 9","CIS 10","CIS 11","CIS 17"]:
    ctrls.append(("CIS Controls v8.1",k,"Control CIS v8.1","Controles CIS"))
for k in ["BM-OS","BM-CLOUD","BM-DB","BM-NET","BM-CONT","BM-APP","BM-M365"]:
    ctrls.append(("CIS Controls v8.1",k,"Benchmark CIS v8.1","Benchmarks CIS"))
for k in ["BAI06","BAI07","BAI09","BAI10"]:
    ctrls.append(("COBIT 5",k,"Proceso COBIT 5","BAI"))
for k in ["PCI Req 2","PCI Req 5","PCI Req 7","PCI Req 8"]:
    ctrls.append(("PCI DSS v4.0.1",k,"Requisito PCI DSS v4.0.1","Requisitos PCI"))
for k in ["L21.719 Art.14 sexies","L21.719 Art.49","L21.719 Art.50","L21.719 Programa adecuación",
          "L21.459 Art.14","L21.459 Art.18 N1","L20.393 Art.4","L20.393 Art.4 N3","L20.393 Art.4 N2","PNCS"]:
    ctrls.append(("Normativa chilena",k,"Norma chilena (txt de alcance)","Normativa"))
ctrls.append(("Normativa chilena","L21.663 Ley de Ciberseguridad","Norma chilena (ciberseguridad)","Normativa"))
assert len(ctrls)==171, len(ctrls)

# --- v7: descripciones en español para la hoja Controles marcos ---
_DESC={"NIST GV.OC-01": "Verificar que la misión y los objetivos del negocio se conocen y se usan para decidir el nivel de riesgo de privacidad y seguridad.", "NIST GV.OC-02": "Verificar que están identificados clientes, empleados, reguladores y proveedores, y qué esperan de la organización en privacidad y seguridad.", "NIST GV.OC-03": "Verificar que existe una lista de leyes y contratos aplicables (incluida la Ley 21.719) y que alguien es responsable de mantenerla.", "NIST GV.OC-04": "Verificar que los compromisos con clientes y terceros (por ejemplo, de seguridad de pagos) están documentados y comunicados.", "NIST GV.OC-05": "Verificar que se conocen los servicios de los que depende la organización (nube, pagos, logística) y su criticidad.", "NIST GV.RR-01": "Verificar que la dirección asume la responsabilidad del riesgo de seguridad y privacidad (actas, reportes a directorio).", "NIST GV.RR-02": "Verificar que los roles de seguridad y privacidad están definidos, comunicados y que sus responsables saben qué deben hacer.", "NIST GV.RR-03": "Verificar que el presupuesto y el personal para seguridad y privacidad son suficientes para el alcance real.", "NIST GV.RR-04": "Verificar que la contratación, la evaluación y la salida de personal incluyen requisitos de seguridad y privacidad.", "NIST GV.PO-01": "Verificar que existe una política de gestión de riesgo de ciberseguridad aprobada, difundida y que se aplica.", "NIST GV.PO-02": "Verificar la fecha de la última revisión de la política y que considera cambios legales y tecnológicos recientes.", "NIST ID.AM-01": "Verificar que existe un listado de equipos (servidores, PC, POS, impresoras) con responsable y ubicación.", "NIST ID.AM-02": "Verificar que existe un listado de software y servicios en uso, incluidos los de nube y los no autorizados.", "NIST ID.AM-03": "Verificar que hay un diagrama de flujos de datos personales dentro de la organización y hacia terceros.", "NIST ID.AM-05": "Verificar que los activos se priorizan (crítico, alto, medio, bajo) y que la prioridad guía los controles.", "NIST ID.AM-08": "Verificar que los sistemas se gestionan desde su compra hasta su baja, con responsables en cada etapa.", "NIST ID.RA-01": "Verificar que se realizan escaneos de vulnerabilidades periódicos y que los hallazgos se registran y validan.", "NIST ID.RA-08": "Verificar que existe un canal público o interno para que terceros reporten vulnerabilidades y un proceso para responderlas.", "NIST PR.AA-01": "Verificar que el alta, cambio y baja de usuarios, servicios y equipos se hacen con procedimiento y registro.", "NIST PR.AA-02": "Verificar que la identidad de cada usuario se confirma antes de entregar acceso (por ejemplo, con cédula o contrato).", "NIST PR.AA-03": "Verificar que todos los accesos exigen autenticación y que las cuentas de servicio también tienen control.", "NIST PR.AA-04": "Verificar que las sesiones y tokens de autenticación se transmiten y almacenan de forma protegida (cifrados, con expiración).", "NIST PR.AA-05": "Verificar que los permisos se asignan por rol, se revisan periódicamente y no se acumulan sin control.", "NIST PR.DS-11": "Verificar que existen respaldos periódicos, cifrados, almacenados fuera del sistema principal y con prueba de restauración.", "NIST PR.PS-01": "Verificar que existen configuraciones base (hardening) para servidores, PC y bases de datos, y que se aplican.", "NIST PR.PS-02": "Verificar que el software tiene soporte del fabricante y que el software sin soporte se retira o aísla.", "NIST PR.PS-03": "Verificar que los equipos sin soporte o con falla se reemplazan según plan y que el retiro es seguro.", "NIST PR.PS-05": "Verificar que los usuarios no pueden instalar software sin autorización y que existe un listado de software permitido.", "NIST RS.MA-01": "Verificar que el plan de respuesta incluye a proveedores clave (hosting, pagos, SOC) y sus contactos.", "NIST RS.MA-02": "Verificar que cada reporte de incidente se clasifica y se confirma antes de actuar.", "NIST RS.MA-03": "Verificar que los incidentes tienen categoría y prioridad según criterios escritos.", "NIST RS.MA-04": "Verificar que existen reglas claras de cuándo escalar un incidente a gerencia, legal o al responsable de privacidad.", "NIST RS.MA-05": "Verificar que hay criterios escritos para decidir cuándo se da por terminada la contención y se inicia la recuperación.", "NIST RS.AN-03": "Verificar que los incidentes importantes tienen un informe de causa raíz con fecha y responsable.", "NIST RS.AN-06": "Verificar que las acciones durante un incidente quedan registradas (quién, qué, cuándo) sin posibilidad de alteración.", "NIST RS.AN-07": "Verificar que se conserva la información del incidente (logs, capturas, correos) con cadena de custodia.", "NIST RS.AN-08": "Verificar que se estima cuántos datos y personas afectó un incidente y que la estimación se valida.", "NIST RS.CO-02": "Verificar que están definidos quiénes deben ser notificados de un incidente (autoridad, titulares, clientes) y cómo.", "NIST RS.CO-03": "Verificar que la información compartida durante un incidente se limita a quienes la necesitan.", "NIST RS.MI-01": "Verificar que existen pasos para aislar sistemas o cuentas comprometidas en un incidente y que se han probado.", "NIST RS.MI-02": "Verificar que, tras un incidente, se elimina la causa (malware, cuenta, vulnerabilidad) antes de restablecer el servicio.", "NIST RC.RP-03": "Verificar que los respaldos se validan antes de restaurar datos, para no restaurar archivos dañados o infectados.", "CIS 1": "Verificar que existe un inventario de todos los equipos conectados a la red, actualizado y con responsable.", "CIS 2": "Verificar que existe un inventario de software instalado y que se controla lo que no está autorizado.", "CIS 4": "Verificar que los equipos y software tienen configuración segura documentada y revisada.", "CIS 5": "Verificar que las cuentas de usuario, administrador y servicio se inventarian, se revisan y se desactivan cuando no se usan.", "CIS 6": "Verificar que el acceso a sistemas y datos se otorga según el rol y se revisa periódicamente.", "CIS 7": "Verificar que se escanea la organización periódicamente y que los hallazgos críticos se corrigen en plazo definido.", "CIS 9": "Verificar que el correo y la navegación tienen filtros contra phishing, malware y sitios maliciosos.", "CIS 10": "Verificar que todos los equipos tienen antivirus o EDR activo, actualizado y monitoreado.", "CIS 11": "Verificar que existen respaldos periódicos de datos críticos y que se prueba la recuperación.", "CIS 17": "Verificar que existe un plan de respuesta a incidentes, roles asignados y ejercicios realizados.", "BM-OS": "Verificar que los servidores y equipos usan una configuración base de un estándar reconocido (por ejemplo, CIS Benchmark) y que se revisa.", "BM-CLOUD": "Verificar que las cuentas y servicios en la nube (AWS, Azure, Google u otros) están configurados según un estándar y revisados.", "BM-DB": "Verificar que las bases de datos tienen configuración segura: usuarios, cifrado, auditoría y puertos restringidos.", "BM-NET": "Verificar que firewalls, switches y routers tienen contraseñas, accesos y servicios configurados según un estándar.", "BM-CONT": "Verificar que los contenedores y la orquestación (Docker, Kubernetes) usan imágenes y configuraciones seguras. Aplica solo si hay contenedores.", "BM-APP": "Verificar que las aplicaciones y navegadores corporativos tienen configuración segura y actualizada.", "BM-M365": "Verificar que las suites de oficina y correo (Microsoft 365, Google Workspace) tienen MFA, reglas de compartición y auditoría activas.", "BAI06": "Verificar que los cambios a sistemas se solicitan, evalúan (incluido el impacto en datos personales), aprueban y registran.", "BAI07": "Verificar que un cambio se prueba y se aprueba formalmente antes de pasar a producción, con evidencia.", "BAI09": "Verificar que los activos de TI tienen registro, responsable y se dan de baja de forma controlada.", "BAI10": "Verificar que existe un registro de la configuración de sistemas y que se mantiene actualizado.", "PCI Req 2": "Verificar que los POS, servidores y equipos del entorno de pagos usan configuraciones seguras y sin valores por defecto.", "PCI Req 5": "Verificar que los equipos del entorno de pagos tienen protección contra malware actualizada y monitoreada.", "PCI Req 7": "Verificar que el acceso a datos de pago se limita a quien lo necesita para su trabajo, con revisión periódica.", "PCI Req 8": "Verificar que cada usuario tiene identificación única y que el acceso al entorno de pago exige autenticación fuerte.", "L21.719 Art.14 sexies": "Verificar que existe un procedimiento para avisar a la autoridad de vulneraciones que afecten datos personales, con plazo definido según el texto oficial.", "L21.719 Art.49": "Verificar que existe un modelo documentado para prevenir infracciones (políticas, controles, capacitación y monitoreo).", "L21.719 Art.50": "Verificar que el delegado de protección tiene funciones, autoridad y acceso a la dirección definidos por escrito.", "L21.719 Programa adecuación": "Verificar que existe un plan de adecuación con responsables, fechas y avance a la fecha.", "L21.459 Art.14": "Verificar que la evidencia electrónica de incidentes se conserva íntegra y con custodia.", "L21.459 Art.18 N1": "Verificar que se puede preservar de forma provisoria datos cuando lo requiera una autoridad, sin borrarlos.", "L20.393 Art.4": "Verificar que existe un modelo de prevención de delitos (incluidos delitos informáticos y de datos) con políticas y controles.", "L20.393 Art.4 N3": "Verificar que existe una persona designada como encargado de prevención, con autoridad y reporte a la dirección.", "L20.393 Art.4 N2": "Verificar que existen protocolos y procedimientos que controlan las actividades de riesgo (por ejemplo, pagos y compras).", "PNCS": "Verificar que la organización conoce la Política Nacional de Ciberseguridad y la considera en su programa.", "5.1": "Existe una política de seguridad aprobada por la dirección, con fecha de revisión y difundida al personal.", "5.2": "Están definidos y documentados quién hace qué en seguridad (matriz de responsabilidades).", "5.3": "Las funciones incompatibles (por ejemplo, crear y aprobar pagos) están separadas o tienen control compensatorio.", "5.4": "La dirección revisa y respalda la seguridad: actas, compromisos y recursos asignados.", "5.5": "Existen contactos definidos con autoridades (Agencia de Protección de Datos, CPLT, ANCI, policía) y canales conocidos.", "5.6": "La organización participa en foros o CERT que entregan información de seguridad.", "5.7": "Se recibe y analiza información sobre amenazas que pueden afectar a la organización.", "5.8": "Los proyectos incluyen revisión de seguridad y privacidad desde su inicio.", "5.9": "Existe un inventario de activos de información con dueño, clasificación y ubicación.", "5.10": "Existe una política de uso aceptable de la información, firmada por los usuarios.", "5.11": "Al salir o cambiar de cargo, el personal devuelve equipos e información, con registro.", "5.12": "La información se clasifica por nivel de sensibilidad y la clasificación se aplica en la práctica.", "5.13": "La clasificación se marca en documentos y sistemas según la política.", "5.14": "Los intercambios con terceros usan canales seguros y acuerdos de confidencialidad.", "5.15": "Existe una política de control de acceso basada en necesidad, aplicada en los sistemas.", "5.16": "El alta, cambio y baja de usuarios están controlados y con evidencia.", "5.17": "Las contraseñas y credenciales se gestionan con reglas formales (longitud, renovación, almacenamiento).", "5.18": "Los permisos se revisan periódicamente y se ajustan al cargo.", "5.19": "Existe un proceso para evaluar la seguridad de los proveedores antes y durante la relación.", "5.20": "Los contratos con proveedores incluyen requisitos de seguridad, confidencialidad, auditoría y devolución de datos.", "5.21": "Se identifican y evalúan los proveedores TIC críticos y sus subcontratistas.", "5.22": "Se revisa el desempeño y el cumplimiento de seguridad de los proveedores.", "5.23": "Los servicios en nube tienen responsabilidades definidas y su configuración se revisa.", "5.24": "Existe un plan de respuesta a incidentes con roles, contactos y canales de reporte.", "5.25": "Los eventos de seguridad se evalúan para decidir si son incidentes.", "5.26": "Los incidentes se atienden según procedimiento, con registro y notificaciones.", "5.27": "Después de cada incidente se documentan las causas y las mejoras.", "5.28": "Se conserva evidencia de los incidentes de forma íntegra para investigaciones.", "5.29": "Existen medidas de seguridad activas durante una interrupción operativa.", "5.30": "Los sistemas críticos tienen plan de recuperación probado, con RTO y RPO definidos.", "5.31": "Se identifican y actualizan las obligaciones legales (incluida la Ley 21.719) y quién es responsable.", "5.32": "Existe control de licencias de software y de propiedad intelectual.", "5.33": "Los registros (contratos, logs, expedientes) están protegidos y tienen plazo de retención.", "5.34": "Existe un programa de privacidad: bases legales, avisos, derechos de los titulares y registro de tratamientos.", "5.35": "La seguridad es revisada por auditoría interna o externa.", "5.36": "Se verifica periódicamente que las políticas se cumplen.", "5.37": "Los procedimientos operativos están documentados y accesibles para quien los necesita.", "6.1": "Se verifican antecedentes del personal antes de contratar, respetando la ley.", "6.2": "Los contratos de trabajo incluyen responsabilidades de seguridad y privacidad.", "6.3": "El personal recibe capacitación en seguridad y privacidad, con registro de asistencia.", "6.4": "Existe un proceso formal de sanciones por incumplimiento de políticas.", "6.5": "Al término o cambio de cargo se revocan accesos y se recuperan activos.", "6.6": "Existen acuerdos de confidencialidad firmados por el personal y terceros.", "6.7": "El trabajo remoto tiene reglas sobre equipos, redes y lugares de trabajo.", "6.8": "El personal sabe cómo y dónde reportar un evento sospechoso.", "7.1": "Las áreas sensibles están delimitadas con barreras físicas.", "7.2": "El ingreso a áreas seguras está controlado y registrado.", "7.3": "Oficinas, bodegas y recintos tienen protección física acorde a su sensibilidad.", "7.4": "Existe vigilancia (CCTV, alarmas) con plazo de retención definido.", "7.5": "Existe protección contra incendio, inundación, temperatura y otras amenazas ambientales.", "7.6": "Existen reglas para trabajar en áreas seguras.", "7.7": "Documentos y pantallas quedan protegidos cuando el personal se ausenta.", "7.8": "Los equipos están ubicados para evitar acceso o daño no autorizado.", "7.9": "Los equipos y datos fuera de la oficina tienen protección (cifrado, control de salida).", "7.10": "Los medios removibles (USB, discos) están controlados y cifrados.", "7.11": "Energía, agua y comunicaciones están protegidas (UPS, respaldos).", "7.12": "El cableado está protegido contra daño o interceptación.", "7.13": "Los mantenimientos están programados y registrados.", "7.14": "Los equipos y medios se borran o destruyen de forma segura y con certificado.", "8.1": "Los equipos de usuario tienen configuración, cifrado y protección mínima definida.", "8.2": "Las cuentas administrativas están controladas, con registro y revisión.", "8.3": "El acceso a la información se limita según el rol.", "8.4": "El código fuente tiene acceso restringido y control de versiones.", "8.5": "Los sistemas con datos personales exigen autenticación robusta (MFA y contraseñas seguras).", "8.6": "La capacidad de servidores y servicios se monitorea.", "8.7": "Existe antivirus o EDR activo y actualizado en todos los equipos.", "8.8": "Existen escaneos de vulnerabilidades y parches con plazos definidos.", "8.9": "Existen configuraciones base (hardening) definidas y controladas.", "8.10": "Los datos se eliminan al cumplir su plazo, con evidencia de la eliminación.", "8.11": "Los datos personales se enmascaran en pruebas, reportes o entornos no productivos.", "8.12": "Existen controles para evitar envíos o copias no autorizadas de datos.", "8.13": "Existen respaldos periódicos, cifrados y con prueba de restauración.", "8.14": "Los sistemas críticos tienen redundancia de infraestructura.", "8.15": "Se registran accesos y acciones sobre datos, con logs protegidos y retención definida.", "8.16": "Existen alertas y revisión periódica de logs.", "8.17": "Los relojes de los sistemas están sincronizados.", "8.18": "Las herramientas de administración con privilegios están restringidas.", "8.19": "Solo se instala software autorizado.", "8.20": "Las redes están protegidas con firewalls y reglas definidas.", "8.21": "Los servicios de red (VPN, DNS, otros) están asegurados y acordados.", "8.22": "Las redes están separadas según sensibilidad de la información.", "8.23": "La navegación está filtrada contra sitios riesgosos.", "8.24": "El cifrado usa estándares definidos, en tránsito y en reposo.", "8.25": "El ciclo de desarrollo integra seguridad desde el diseño.", "8.26": "Los requisitos de seguridad se definen antes de desarrollar o comprar una aplicación.", "8.27": "Se aplican principios de arquitectura segura.", "8.28": "El código se escribe y revisa con prácticas de codificación segura.", "8.29": "Existen pruebas de seguridad antes de pasar a producción.", "8.30": "Los proveedores de desarrollo cumplen requisitos de seguridad.", "8.31": "Los entornos de desarrollo, prueba y producción están separados.", "8.32": "Los cambios se aprueban y registran antes de implementarse.", "8.33": "Los datos de prueba no contienen datos personales reales sin protección.", "8.34": "Las auditorías se planifican para no afectar la operación."}; _NAME={"NIST GV.OC-01": "Misión y contexto de la organización", "NIST GV.OC-02": "Partes interesadas y sus expectativas", "NIST GV.OC-03": "Obligaciones legales, regulatorias y contractuales", "NIST GV.OC-04": "Expectativas de partes externas sobre la organización", "NIST GV.OC-05": "Dependencias y servicios críticos de la organización", "NIST GV.RR-01": "Liderazgo responsable del riesgo de ciberseguridad", "NIST GV.RR-02": "Roles, responsabilidades y autoridades definidos", "NIST GV.RR-03": "Recursos asignados a la ciberseguridad", "NIST GV.RR-04": "Ciberseguridad incluida en RRHH", "NIST GV.PO-01": "Política de gestión del riesgo establecida", "NIST GV.PO-02": "Política revisada y actualizada", "NIST ID.AM-01": "Inventario de hardware", "NIST ID.AM-02": "Inventario de software y servicios", "NIST ID.AM-03": "Mapa de flujos de datos y comunicaciones", "NIST ID.AM-05": "Priorización de activos por criticidad", "NIST ID.AM-08": "Gestión de activos a lo largo de su ciclo de vida", "NIST ID.RA-01": "Identificación de vulnerabilidades", "NIST ID.RA-08": "Recepción de divulgaciones de vulnerabilidades", "NIST PR.AA-01": "Gestión de identidades y credenciales", "NIST PR.AA-02": "Verificación de identidades", "NIST PR.AA-03": "Autenticación de usuarios, servicios y equipos", "NIST PR.AA-04": "Protección de aserciones de identidad", "NIST PR.AA-05": "Permisos con mínimo privilegio y separación de funciones", "NIST PR.DS-11": "Respaldos de datos", "NIST PR.PS-01": "Gestión de configuración segura", "NIST PR.PS-02": "Software mantenido o retirado según riesgo", "NIST PR.PS-03": "Hardware mantenido o retirado según riesgo", "NIST PR.PS-05": "Prevención de software no autorizado", "NIST RS.MA-01": "Plan de respuesta ejecutado con terceros", "NIST RS.MA-02": "Reportes de incidentes clasificados y validados", "NIST RS.MA-03": "Incidentes categorizados y priorizados", "NIST RS.MA-04": "Escalamiento de incidentes", "NIST RS.MA-05": "Criterios para iniciar la recuperación", "NIST RS.AN-03": "Análisis de causa raíz", "NIST RS.AN-06": "Registro de acciones durante la investigación", "NIST RS.AN-07": "Recolección de datos del incidente", "NIST RS.AN-08": "Estimación de la magnitud del incidente", "NIST RS.CO-02": "Notificación a partes interesadas", "NIST RS.CO-03": "Información compartida con partes designadas", "NIST RS.MI-01": "Contención de incidentes", "NIST RS.MI-02": "Erradicación de incidentes", "NIST RC.RP-03": "Verificación de integridad de respaldos antes de restaurar", "CIS 1": "Inventario de activos empresariales", "CIS 2": "Inventario de software", "CIS 4": "Configuración segura de activos y software", "CIS 5": "Gestión de cuentas", "CIS 6": "Gestión del control de acceso", "CIS 7": "Gestión continua de vulnerabilidades", "CIS 9": "Protección de correo y navegador", "CIS 10": "Defensas contra malware", "CIS 11": "Recuperación de datos", "CIS 17": "Gestión de respuesta a incidentes", "BM-OS": "Benchmark de sistemas operativos", "BM-CLOUD": "Benchmark de proveedores de nube", "BM-DB": "Benchmark de bases de datos", "BM-NET": "Benchmark de dispositivos de red", "BM-CONT": "Benchmark de contenedores y orquestación", "BM-APP": "Benchmark de aplicaciones y navegadores", "BM-M365": "Benchmark de suites de productividad", "BAI06": "Gestión de cambios", "BAI07": "Aceptación del cambio y de la transición", "BAI09": "Gestión de activos", "BAI10": "Gestión de la configuración", "PCI Req 2": "Configuraciones seguras en todos los componentes", "PCI Req 5": "Protección de sistemas contra software malicioso", "PCI Req 7": "Acceso restringido según necesidad de saber", "PCI Req 8": "Identificación de usuarios y autenticación del acceso", "L21.719 Art.14 sexies": "Deber de reportar vulneraciones de seguridad", "L21.719 Art.49": "Modelo de prevención de infracciones", "L21.719 Art.50": "Atribuciones del delegado de protección de datos", "L21.719 Programa adecuación": "Programa de adecuación de la organización", "L21.459 Art.14": "Preservación y custodia de evidencia electrónica", "L21.459 Art.18 N1": "Preservación provisoria de datos", "L20.393 Art.4": "Modelo de prevención de delitos vigente", "L20.393 Art.4 N3": "Encargado de prevención de delitos", "L20.393 Art.4 N2": "Protocolos y procedimientos de control", "PNCS": "Política Nacional de Ciberseguridad", "5.1": "Políticas de seguridad de la información", "5.2": "Roles y responsabilidades de SI", "5.3": "Segregación de funciones", "5.4": "Responsabilidades de la dirección", "5.5": "Contacto con autoridades", "5.6": "Contacto con grupos de interés especial", "5.7": "Inteligencia de amenazas", "5.8": "SI en gestión de proyectos", "5.9": "Inventario de información y activos", "5.10": "Uso aceptable de información", "5.11": "Devolución de activos", "5.12": "Clasificación de la información", "5.13": "Etiquetado de la información", "5.14": "Transferencia de información", "5.15": "Control de acceso", "5.16": "Gestión de identidades", "5.17": "Información de autenticación", "5.18": "Derechos de acceso", "5.19": "Seguridad en relaciones con proveedores", "5.20": "Seguridad en contratos con proveedores", "5.21": "Gestión de seguridad en cadena TIC", "5.22": "Monitoreo de proveedores", "5.23": "SI en servicios en la nube", "5.24": "Planificación y preparación de gestión de incidentes", "5.25": "Evaluación y decisión sobre eventos de SI", "5.26": "Respuesta a incidentes de SI", "5.27": "Aprendizaje de los incidentes", "5.28": "Recolección de evidencia", "5.29": "Seguridad de la información durante disrupción", "5.30": "Preparación TIC para continuidad", "5.31": "Requisitos legales, estatutarios y contractuales", "5.32": "Derechos de propiedad intelectual", "5.33": "Protección de registros", "5.34": "Privacidad y protección de datos personales", "5.35": "Revisión independiente de SI", "5.36": "Cumplimiento con políticas y normas de SI", "5.37": "Procedimientos de operación documentados", "6.1": "Selección de personal", "6.2": "Términos y condiciones de empleo", "6.3": "Concientización, educación y capacitación en SI", "6.4": "Proceso disciplinario", "6.5": "Responsabilidades ante terminación o cambio de empleo", "6.6": "Acuerdos de confidencialidad", "6.7": "Trabajo remoto", "6.8": "Reporte de eventos de SI", "7.1": "Perímetros de seguridad física", "7.2": "Controles de entrada física", "7.3": "Seguridad de oficinas, recintos e instalaciones", "7.4": "Monitoreo de seguridad física", "7.5": "Protección contra amenazas físicas y ambientales", "7.6": "Trabajo en áreas seguras", "7.7": "Escritorio limpio y pantalla limpia", "7.8": "Ubicación y protección de equipos", "7.9": "Seguridad de activos fuera de las instalaciones", "7.10": "Medios de almacenamiento", "7.11": "Servicios de suministro", "7.12": "Seguridad del cableado", "7.13": "Mantenimiento de equipos", "7.14": "Eliminación o reutilización segura de equipos", "8.1": "Dispositivos de usuario final", "8.2": "Derechos de acceso privilegiado", "8.3": "Restricción de acceso a la información", "8.4": "Acceso al código fuente", "8.5": "Autenticación segura", "8.6": "Gestión de capacidad", "8.7": "Protección contra malware", "8.8": "Gestión de vulnerabilidades técnicas", "8.9": "Gestión de la configuración", "8.10": "Eliminación de información", "8.11": "Enmascaramiento de datos", "8.12": "Prevención de fuga de datos", "8.13": "Respaldo de información", "8.14": "Redundancia de instalaciones de procesamiento", "8.15": "Registro de actividades (logging)", "8.16": "Actividades de monitoreo", "8.17": "Sincronización de relojes", "8.18": "Uso de programas de utilidad con privilegios", "8.19": "Instalación de software en sistemas operativos", "8.20": "Seguridad en redes", "8.21": "Seguridad de los servicios de red", "8.22": "Segregación de redes", "8.23": "Filtrado web", "8.24": "Uso de criptografía", "8.25": "Ciclo de vida de desarrollo seguro", "8.26": "Requisitos de seguridad en aplicaciones", "8.27": "Principios de arquitectura segura", "8.28": "Codificación segura", "8.29": "Pruebas de seguridad en desarrollo y aceptación", "8.30": "Desarrollo externalizado", "8.31": "Separación de entornos", "8.32": "Gestión de cambios", "8.33": "Información de prueba", "8.34": "Protección de sistemas de información en auditoría"}; _TEMA={"GOVERN": "Gobernar (GOVERN)", "IDENTIFY": "Identificar (IDENTIFY)", "PROTECT": "Proteger (PROTECT)", "RESPOND": "Responder (RESPOND)", "RECOVER": "Recuperar (RECOVER)", "Controles CIS": "Controles CIS v8.1", "Benchmarks": "Benchmarks CIS v8.1", "BAI": "Construir, adquirir e implementar (BAI)", "Requisitos PCI": "Requisitos PCI DSS v4.0.1", "Normativa": "Normativa chilena"}
_DESC["L21.663 Ley de Ciberseguridad"]="Verificar si la organización está sujeta a la Ley 21.663 (servicios esenciales u operadores de importancia vital) y qué obligaciones de ciberseguridad y de reporte de incidentes le aplican."
_NAME["L21.663 Ley de Ciberseguridad"]="Ley 21.663 Marco de ciberseguridad"
_new=[]
for _m,_k,_d,_t in ctrls:
    _n=_NAME.get(_k,_d)
    _tt=_TEMA.get(_t,_t)
    _new.append((_m,_k,f"{_n}. Qué debe levantar: {_DESC[_k]}",_tt))
ctrls=_new

keys=[c[1] for c in ctrls]; assert len(set(keys))==len(keys)
ROW={k:i+2 for i,k in enumerate(keys)}
CMN="'Controles marcos'"

# ---------- mapeo requisito -> controles ----------
MAP={
"1.1":["5.1","NIST GV.PO-01"],"1.2":["5.2","5.3","NIST GV.RR-02","L21.719 Art.50"],
"1.3":["5.4","NIST GV.RR-03"],"1.4":["5.31","5.36","NIST GV.OC-03","L21.719 Programa adecuación","L20.393 Art.4"],
"1.5":["5.36","5.35","NIST GV.PO-02"],"1.6":["5.31","NIST GV.OC-02"],
"2.1":["RAT","5.9"],"2.2":["5.9","NIST ID.AM-02"],"2.3":["NIST ID.AM-03"],"2.4":["NIST ID.AM-01","5.9"],"2.5":["NIST ID.AM-05"],
"3.1":["5.34"],"3.2":["5.34"],"3.3":["5.34","5.33"],"3.4":["5.34"],"3.5":["5.34"],"3.6":["5.34"],
"4.1":["ARCOP","5.34"],"4.2":["ARCOP"],"4.3":["5.34","5.18"],"4.4":["5.33","8.10"],"4.5":["5.33"],"4.6":["5.34"],
"5.1":["5.34"],"5.2":["5.34","8.11"],"5.3":["5.34"],"5.4":["5.33"],"5.5":["8.10"],"5.6":["5.19","5.14"],
"6.1":["5.12","5.9"],"6.2":["5.34"],"6.3":["8.11","8.24"],"6.4":["5.34"],"6.5":["5.34"],
"7.1":["5.34"],"7.2":["5.34","5.8"],"7.3":["5.8","8.25","8.27"],"7.4":["8.9"],"7.5":["5.34"],
"8.1":["5.1"],"8.2":["8.2","5.15"],"8.3":["5.17","8.5"],"8.4":["8.24"],"8.5":["8.8"],"8.6":["8.15"],
"8.7":["8.13"],"8.8":["8.8","5.35"],"8.9":["8.11"],"8.10":["8.2","8.16"],
"9.1":["5.24","NIST RS.MA-01"],"9.2":["6.8","5.25","NIST RS.AN-07"],"9.3":["5.26","L21.719 Art.14 sexies"],
"9.4":["5.26","NIST RS.CO-02"],"9.5":["5.26","NIST RS.CO-03"],"9.6":["5.24"],"9.7":["5.27","NIST RS.AN-03"],
"9.8":["5.29","5.30","8.14"],
"10.1":["5.19"],"10.2":["5.20"],"10.3":["5.21","5.22"],"10.4":["5.20"],"10.5":["TRANSF","5.14"],"10.6":["6.6"],
"11.1":["6.3"],"11.2":["6.3"],"11.3":["6.3","6.8"],"11.4":["6.6"],
"12.1":["5.35"],"12.2":["5.4"],"12.3":["5.35","5.36"],"12.4":["5.1"],"12.5":["5.31","5.7"],
"13.1":["NIST ID.RA-01"],"13.2":["NIST ID.RA-08"],"13.3":["NIST PR.PS-01","CIS 4"],"13.4":["NIST PR.PS-02","NIST PR.PS-03"],
"13.5":["NIST PR.PS-05","CIS 2"],"13.6":["CIS 9"],"13.7":["CIS 10"],"13.8":["NIST RS.AN-06","NIST RS.AN-07"],
"13.9":["L21.459 Art.14","L21.459 Art.18 N1"],"13.10":["NIST RS.MI-01","NIST RS.MI-02"],"13.11":["NIST RC.RP-03","CIS 11"],
"13.12":["COBIT BAI06","COBIT BAI07"],"13.13":["COBIT BAI09","COBIT BAI10"],
"13.14":["BM-OS","BM-CLOUD","BM-DB","BM-NET","BM-APP","BM-M365"],
"13.15":["PCI Req 2","PCI Req 5"],"13.16":["PCI Req 7","PCI Req 8"],"13.17":["L20.393 Art.4"],"13.18":["PNCS"],
}

# ---------- v3: extensión del mapeo, D14 y pesos ----------
DIMS=[
("D1","Gobierno y responsabilidad proactiva",8),
("D2","Inventario y Registro de Actividades de Tratamiento (RAT)",6),
("D3","Licitud del tratamiento y consentimiento",6),
("D4","Derechos de los titulares (ARCOP)",6),
("D5","Principios y ciclo de vida del dato",5),
("D6","Datos sensibles, biométricos y de menores",5),
("D7","Evaluación de impacto y privacidad desde el diseño",4),
("D8","Seguridad de la información (RBAC, malware, MFA)",11),
("D9","Gestión de incidentes, continuidad y DRP",7),
("D10","Encargados, proveedores y transferencias internacionales",5),
("D11","Capacitación, cultura y phishing",4),
("D12","Auditoría, monitoreo y mejora continua",3),
("D13","Marcos de referencia complementarios (NIST CSF 2.0, CIS, PCI DSS, normativa)",6),
("D14","Seguridad física y del entorno",3),
("D15","Infraestructura, comunicaciones y bases de datos",6),
("D16","Gobierno de TI y procesos (COBIT 5)",4),
("D17","Recursos humanos: candidatos, empleados y expediente",6),
("D18","Clientes, leads, prospectos y marketing",5),
]
ITEMS["D14"]=[]
ITEMS["D15"]=[]
ITEMS["D16"]=[]
ITEMS["D17"]=[]
ITEMS["D18"]=[]
EXTRA={
"D1":[("Contacto formal con autoridades y reguladores (Agencia de Protección de Datos, CPLT, ANCI) mantenido y actualizado.",["5.5","5.6"]),
("Cultura de cumplimiento y gestión del riesgo de privacidad liderada por la dirección.",["NIST GV.RR-01","NIST GV.OC-01"]),
("Marco de cumplimiento comunicado a partes interesadas externas.",["NIST GV.OC-04","NIST GV.OC-05"]),
("Procedimientos operativos documentados para el tratamiento de datos personales.",["5.37"]),
("Cumplimiento de licencias y propiedad intelectual en sistemas que tratan datos.",["5.32"])],
"D2":[("Clasificación y etiquetado de datos personales en repositorios y documentos.",["5.12","5.13"]),
("Gestión de sistemas y datos a lo largo de su ciclo de vida.",["NIST ID.AM-08"])],
"D8":[("Protección de dispositivos de usuario final (MDM, cifrado, EDR).",["8.1"]),
("Protección contra malware en endpoints y servidores.",["8.7","CIS 10"]),
("Restricción de acceso a información y código fuente.",["8.3","8.4"]),
("Prevención de fuga de datos.",["8.12"]),
("Gestión de capacidad y sincronización de sistemas.",["8.6","8.17"]),
("Seguridad de redes, servicios de red y segregación.",["8.20","8.21","8.22","8.23"]),
("Control de utilidades privilegiadas y de software instalado.",["8.18","8.19","CIS 2"]),
("Gestión de identidades y credenciales.",["5.16","NIST PR.AA-01","NIST PR.AA-02","CIS 5"]),
("Autenticación e integridad de aserciones de identidad.",["NIST PR.AA-03","NIST PR.AA-04","5.17"]),
("Permisos y revisión de accesos según mínimo privilegio.",["5.18","NIST PR.AA-05","CIS 6"]),
("Respaldos protegidos y resguardados.",["NIST PR.DS-11"]),
("Desarrollo seguro: requisitos, codificación y pruebas.",["8.26","8.28","8.29"]),
("Externalización de desarrollo y separación de entornos.",["8.30","8.31","8.33"]),
("Gestión de cambios en sistemas en producción.",["8.32","COBIT BAI06"])],
"D9":[("Clasificación, priorización y escalamiento de incidentes.",["NIST RS.MA-02","NIST RS.MA-03","NIST RS.MA-04","NIST RS.MA-05"]),
("Estimación de magnitud y recolección de evidencia de incidentes.",["NIST RS.AN-08","5.28"])],
"D10":[("Servicios en la nube evaluados y gestionados por seguridad y privacidad.",["5.23"])],
"D11":[("Selección y términos contractuales del personal con cláusulas de privacidad.",["6.1","6.2"]),
("Proceso disciplinario por incumplimiento de políticas de datos.",["6.4"]),
("Terminación o cambio de empleo con revocación de accesos y devolución de activos.",["6.5","5.11"]),
("Trabajo remoto y uso aceptable de la información.",["6.7","5.10"]),
("Ciberseguridad y privacidad integradas en prácticas de RRHH.",["NIST GV.RR-04"])],
"D12":[("Protección de sistemas de información durante auditorías.",["8.34"])],
"D13":[("Benchmarks de contenedores y orquestación aplicados (si aplica).",["BM-CONT"])],
"D14":[("Perímetros de seguridad física y control de entrada.",["7.1","7.2"]),
("Protección de oficinas, recintos y áreas seguras.",["7.3","7.6"]),
("Monitoreo físico y protección frente a amenazas ambientales.",["7.4","7.5"]),
("Escritorio limpio, pantalla limpia y ubicación de equipos.",["7.7","7.8"]),
("Activos fuera de instalaciones, medios de almacenamiento y eliminación segura de equipos.",["7.9","7.10","7.14"]),
("Servicios de suministro, cableado y mantenimiento de equipos.",["7.11","7.12","7.13"])],
}

EXTRA.setdefault("D1",[]).append(("Políticas técnicas aprobadas y vigentes: control de acceso por roles, respaldo y recuperación, malware, phishing y uso aceptable.",["5.10","5.1","NIST GV.PO-01"]))
EXTRA.setdefault("D8",[]).extend([
("Modelo RBAC documentado: roles, perfiles y matriz de accesos revisada periódicamente.",["5.15","5.18","NIST PR.AA-05","PCI Req 7"]),
("Autenticación multifactor y control de credenciales según PCI DSS.",["8.5","PCI Req 8"]),
("Protección contra malware alineada con PCI DSS: antivirus/EDR actualizado y monitoreado.",["8.7","PCI Req 5"])])
EXTRA.setdefault("D11",[]).append(("Protección contra phishing: filtros de correo, MFA y simulacros con métricas.",["CIS 9","6.3","6.8"]))
EXTRA.setdefault("D13",[]).append(("Alcance del entorno de pagos (tiendas, bodegas y ecommerce) identificado y segmentado.",["PCI Req 2","PCI Req 7"]))
EXTRA["D15"]=[
("Inventario y clasificación de servidores, equipos de red y servicios en la nube.",["CIS 1","NIST ID.AM-01"]),
("Hardening de sistemas operativos según benchmark aplicable.",["BM-OS","CIS 4"]),
("Segmentación de red por nivel de sensibilidad (VLAN, zonas y reglas de firewall).",["8.22","8.20","NIST ID.AM-03"]),
("Hardening de bases de datos y control de acceso privilegiado a ellas.",["BM-DB","8.2"]),
("Cifrado de bases de datos y de datos personales en reposo.",["8.24","8.11"]),
("Registro de consultas y cambios sobre datos personales en bases de datos.",["8.15","8.16"]),
("Enmascaramiento o anonimización de datos personales en entornos de prueba.",["8.33","8.11"]),
("Comunicaciones externas cifradas (TLS 1.2 o superior) y acceso remoto por VPN.",["8.24","6.7","8.20"]),
("Protección del correo y la navegación: filtrado, autenticación de dominio y bloqueo de sitios riesgosos.",["CIS 9","8.23"]),
("Infraestructura en nube y contenedores con configuración segura.",["BM-CLOUD","BM-CONT","5.23"]),
("Plataformas de productividad (M365) configuradas según benchmark.",["BM-M365","CIS 4"]),
("Respaldo y restauración de bases de datos con pruebas documentadas.",["8.13","CIS 11","NIST RC.RP-03"]),
("Entornos de pago (ecommerce, POS y bodegas) con configuración segura.",["PCI Req 2","PCI Req 5"]),
("Recuperación ante desastres (DRP) con RTO y RPO probados para bases de datos y sistemas críticos.",["5.30","8.14"]),
]
EXTRA["D16"]=[
("Gestión de cambios con evaluación de impacto en privacidad (BAI06).",["COBIT BAI06","8.32"]),
("Aceptación formal de cambios y transición a producción (BAI07).",["COBIT BAI07","8.29"]),
("Gestión de activos de TI durante su ciclo de vida (BAI09).",["COBIT BAI09","5.9"]),
("Gestión de configuración y registro de elementos de configuración (BAI10).",["COBIT BAI10","8.9"]),
("Riesgo de TI integrado con la gestión de riesgos del SGSI y reportado a la dirección.",["NIST GV.RR-03","NIST GV.OC-02"]),
]


EXTRA["D10"]=EXTRA.get("D10",[])+[
("Sub-encargados identificados y autorizados contractualmente.",["5.20","5.21"]),
("Devolución o eliminación de datos al término del contrato, con certificado.",["5.20","8.10"])]
EXTRA["D17"]=[
("Aviso de privacidad y finalidad para candidatos; CV y entrevistas tratados solo para selección.",["5.34","6.1"]),
("Plazo de retención de candidatos no contratados y eliminación segura al cumplirse.",["8.10","5.33"]),
("Evaluaciones psicológicas, tests y perfilamiento de candidatos con base legal y revisión humana de decisiones.",["5.34","6.1"]),
("Contratos de trabajo, anexos y modificaciones con cláusulas de privacidad y confidencialidad.",["6.2","6.6"]),
("Onboarding con entrega de políticas y accesos según rol aprobado.",["6.3","5.18"]),
("Offboarding con revocación de accesos, devolución de activos y cierre de expediente.",["6.5","5.11","5.18"]),
("Expediente de empleado con acceso restringido, retención definida y eliminación al vencer el plazo.",["5.33","8.10"]),
("Datos de salud (licencias médicas, certificados) separados del expediente general, con acceso restringido.",["5.34","8.11"]),
("Evaluaciones de desempeño y de otra índole con finalidad declarada y acceso del titular a su resultado.",["5.34","5.37"]),
("Registros de capacitación y cursos con retención y acceso definidos.",["6.3"]),
("Control de asistencia y acceso: si usa huella o reconocimiento facial, tratamiento de dato biométrico con consentimiento explícito y medidas reforzadas.",["5.34","7.2","8.15"]),
("Videovigilancia en tiendas y bodegas con aviso, finalidad y retención definida.",["7.4","5.34"]),
("Monitoreo de correo y equipos de trabajo informado al personal y proporcional.",["5.34","8.15"]),
("Transferencia de datos de nómina a previsión y salud con base legal documentada.",["5.19","5.14"])]
EXTRA["D18"]=[
("Origen lícito de bases de datos de clientes, leads y prospectos (propias o compradas) documentado.",["5.34","5.14"]),
("Consentimiento para marketing y prospección a personas naturales; prospección B2B con base legal documentada.",["5.34"]),
("Cookies y seguimiento web con consentimiento previo y sin terceros no informados.",["5.34","8.23"]),
("Datos de ecommerce (compras y pagos) con control de acceso y cifrado.",["8.24","PCI Req 7"]),
("Perfilamiento y segmentación de clientes informados y, si aplica, con evaluación de impacto.",["5.34","5.8"]),
("Programas de fidelización y puntos con finalidad declarada y consentimiento.",["5.34","5.9"]),
("Opt-out de marketing atendido de inmediato y registrado.",["5.34","8.10"]),
("Retención de leads y clientes inactivos con plazo definido y eliminación.",["8.10","5.33"]),
("Grabaciones de call center con aviso, finalidad y retención definida.",["5.34","8.15"]),
("Control de acceso al CRM y restricción de exportaciones masivas.",["8.3","8.12"])]


GP_ITEMS=[
("D1","Comité de privacidad formal con actas que revisa tratamientos de alto riesgo y decisiones de cumplimiento.",["5.4","NIST GV.RR-01"]),
("D1","Evaluación de privacidad y seguridad antes de contratar proveedores tecnológicos nuevos (due diligence).",["5.19","5.21"]),
("D3","Gestión de cookies con plataforma de consentimiento, con registro de versiones y posibilidad de revocar.",["5.34","8.23"]),
("D4","Procedimiento para solicitudes ARCOP que llegan por canales informales (WhatsApp, redes sociales, call center), con registro y plazo.",["ARCOP","5.34"]),
("D5","Minimización en formularios web y POS: solo campos obligatorios justificados por la finalidad.",["5.34","8.11"]),
("D5","Eliminación automática de datos según la tabla de retención.",["5.33","8.10"]),
("D6","Verificación de edad en ecommerce cuando se venden productos restringidos.",["5.34"]),
("D8","Prevención de fuga de datos (DLP) en correo y nube.",["8.12"]),
("D8","Política de dispositivos personales (BYOD) y de trabajo remoto con cifrado y borrado remoto.",["8.1","6.7"]),
("D9","Cadena de custodia documentada para la evidencia de incidentes.",["5.28"]),
("D9","Seguro de ciberseguridad revisado en relación con las obligaciones de notificación.",["5.31"]),
("D10","Cláusulas tipo de transferencia internacional revisadas al menos una vez al año para proveedores de marketing y nube.",["5.14","5.20"]),
("D13","Ley 21.663 de Ciberseguridad identificada y cruzada con las obligaciones de la Ley 21.719.",["L21.663 Ley de Ciberseguridad","5.31"]),
]
GP_TEXTS=set(t for _,t,_ in GP_ITEMS)
for _code,_t,_refs in GP_ITEMS:
    EXTRA.setdefault(_code,[]).append((_t,_refs))


# --- requisitos incorporados: decisiones automatizadas, derechos ARCOP y retención verificable ---
EXTRA.setdefault("D7",[]).extend([
 ("Las decisiones que afectan a personas y se toman solo por sistemas automáticos están identificadas. Verificar: inventario de decisiones automatizadas (por ejemplo, scoring, aprobaciones o precios personalizados).",["5.34"]),
 ("Ante una decisión automatizada, el titular puede pedir información sobre su lógica y solicitar revisión por una persona. Verificar: procedimiento escrito y un caso de prueba documentado.",["5.34","5.8"])])
EXTRA.setdefault("D4",[]).extend([
 ("La organización puede localizar los datos de un titular en todos sus sistemas y proveedores dentro de un plazo razonable. Verificar: prueba con un caso real o de prueba que muestre los sistemas consultados y el tiempo empleado.",["5.33","8.15"]),
 ("Los datos entregados por portabilidad se exportan en un formato estructurado y de uso común. Verificar: archivo de ejemplo exportado y formato utilizado.",["5.34"]),
 ("Una supresión llega también a los respaldos y a los proveedores, con un procedimiento para hacerlo en plazos definidos. Verificar: procedimiento escrito y último caso ejecutado.",["8.13","5.20"])])
EXTRA.setdefault("D5",[]).extend([
 ("Existe una tabla de retención por categoría de dato y las eliminaciones se ejecutan y quedan registradas. Verificar: último registro de eliminación, con fecha y responsable.",["5.33","8.10"]),
 ("Cada eliminación deja evidencia: qué se eliminó, cuándo, por quién y cómo se verificó. Verificar: muestra de tres registros de eliminación.",["8.10"])])
for code,ext in EXTRA.items():
    start=len(ITEMS[code])
    for n,(txt,refs) in enumerate(ext):
        iid=f"{code[1:]}.{start+n+1}"
        ITEMS[code].append((iid,txt)); MAP[iid]=refs
for rid,refs in {"1.2":["L20.393 Art.4 N3"],"1.4":["L21.719 Art.49","L20.393 Art.4 N2"],"2.2":["CIS 1"],"8.5":["CIS 7"],"9.1":["CIS 17"]}.items():
    MAP[rid]=MAP[rid]+refs

# normaliza prefijos usados en el mapeo al id de la hoja
def key(x):
    if x in ("RAT","ARCOP","TRANSF"): return x
    if x.startswith("NIST ") or x.startswith("CIS ") or x.startswith("BM-") or x.startswith("PCI ") or x.startswith("L2") or x=="PNCS" or x in ROW: return x
    if x.startswith("COBIT "): return x[6:]
    return x
for rid,lst in MAP.items():
    for k in lst:
        assert key(k) in ROW or key(k) in ("RAT","ARCOP","TRANSF"), (rid,k)
req_ids=[iid for code,_,_ in DIMS for iid,_ in ITEMS[code]]
assert set(req_ids)==set(MAP.keys()), set(req_ids)^set(MAP.keys())
used=set()
for lst in MAP.values():
    for k in lst:
        if key(k) in ROW: used.add(key(k))

REQ_TXT={
"1.1":"La organización tiene una política de protección de datos personales aprobada por la dirección y comunicada a todo el personal. Verificar: documento firmado, fecha de aprobación y constancia de difusión.",
"1.2":"Existe una persona designada formalmente como responsable de privacidad (DPO o equivalente), con funciones y autoridad escritas. Verificar: nombramiento, descripción de funciones y a quién reporta.",
"1.3":"La dirección asigna personas, herramientas y presupuesto para cumplir la ley. Verificar: presupuesto anual de privacidad y dedicación del responsable.",
"1.4":"Existe un programa de adecuación a la ley con modelo de prevención de infracciones documentado. Verificar: plan con responsables, fechas y avance.",
"1.5":"Las evidencias de cumplimiento están ordenadas y se pueden presentar a la autoridad en cualquier momento. Verificar: carpeta de evidencias actualizada y con fechas.",
"1.6":"Para cada tratamiento está definido si la organización actúa como responsable (decide para qué se usan los datos) o encargado (los trata por cuenta de otro). Verificar: matriz de roles por tratamiento.",
"1.7":"Existen contactos formales con la Agencia de Protección de Datos, el CPLT y la ANCI, y se mantienen actualizados. Verificar: directorio de contactos y fecha de la última actualización.",
"1.8":"La dirección lidera la cultura de cumplimiento y el riesgo de privacidad forma parte de sus decisiones. Verificar: actas o reportes donde la dirección revisa privacidad.",
"1.9":"Clientes, proveedores y otras partes externas conocen el compromiso de la organización con la privacidad y la seguridad. Verificar: políticas publicadas y declaraciones en contratos o sitio web.",
"1.10":"Los procedimientos para tratar datos personales están escritos y disponibles para quien los usa. Verificar: lista de procedimientos con versión y fecha.",
"1.11":"Los sistemas que tratan datos cuentan con licencias vigentes y se controla el uso de software con propiedad intelectual. Verificar: inventario de licencias.",
"1.12":"Existen políticas técnicas escritas y aprobadas sobre control de acceso por roles, respaldos, malware, phishing y uso aceptable. Verificar: cada política, su fecha y quién la aprueba.",
"2.1":"Existe un Registro de Actividades de Tratamiento (RAT) que describe, para cada actividad, la finalidad, la base legal, los tipos de datos, los titulares, los destinatarios, los plazos y las medidas de seguridad. Verificar: hoja RAT completa (indicador automático).",
"2.2":"Existe un listado de los sistemas, aplicaciones y bases de datos que tratan datos personales. Verificar: inventario con nombre del sistema, dueño y tipo de datos.",
"2.3":"Hay un diagrama de los flujos de datos personales dentro de la organización y hacia terceros (incluida la nube). Verificar: diagrama actualizado y su fecha.",
"2.4":"El RAT se revisa al menos una vez al año y cada vez que cambia un proceso o sistema. Verificar: fecha de última revisión en la hoja RAT y registro de cambios.",
"2.5":"Los tratamientos de alto riesgo (por ejemplo, datos sensibles a gran escala, biometría o videovigilancia) están identificados en el RAT. Verificar: marca de riesgo en cada actividad.",
"2.6":"Los datos personales están clasificados y etiquetados según su sensibilidad en repositorios y documentos. Verificar: esquema de clasificación y ejemplos etiquetados.",
"2.7":"Los sistemas y datos se gestionan desde su adquisición hasta su baja, con responsable en cada etapa. Verificar: procedimiento de ciclo de vida.",
"3.1":"Cada tratamiento de datos personales tiene una base legal (consentimiento, contrato, obligación legal, interés legítimo u otra) definida antes de iniciarse. Verificar: base legal registrada en el RAT.",
"3.2":"El consentimiento es libre, informado, específico e inequívoco, y no hay casillas premarcadas. Verificar: formularios reales con su captura de pantalla.",
"3.3":"Cada consentimiento queda registrado con fecha, canal, versión del aviso y la finalidad aceptada. Verificar: muestra de registros de consentimiento.",
"3.4":"El titular puede revocar su consentimiento con la misma facilidad con que lo otorgó. Verificar: prueba del proceso de revocación (cuántos pasos y en qué canal).",
"3.5":"Cuando la base legal es interés legítimo, existe una evaluación escrita que pondera ese interés frente a los derechos del titular. Verificar: documento de ponderación por tratamiento.",
"3.6":"Al recolectar datos se entrega al titular el aviso de privacidad completo (quién trata, para qué, base legal, derechos y plazos). Verificar: aviso publicado y se muestra en cada punto de captura.",
"4.1":"Existe un procedimiento formal para atender solicitudes de derechos (acceso, rectificación, supresión, oposición y portabilidad) y un canal publicado para presentarlas. Verificar: procedimiento y canal visible para el titular.",
"4.2":"Los plazos de respuesta a las solicitudes están definidos y se monitorean. Verificar: plazo aplicado (según el texto oficial) y reporte de solicitudes fuera de plazo.",
"4.3":"Antes de entregar o modificar datos, se verifica la identidad del solicitante. Verificar: método de verificación usado.",
"4.4":"La organización puede ubicar, rectificar, suprimir, exportar y restringir los datos de una persona en todos sus sistemas. Verificar: prueba práctica con un caso real o de prueba.",
"4.5":"Cada solicitud, su respuesta y el motivo de cualquier rechazo quedan registrados. Verificar: hoja ARCOP completa.",
"4.6":"Si hay decisiones automatizadas que afectan al titular, se le informa y puede pedir revisión humana. Verificar: si existen decisiones automatizadas (por ejemplo, scoring de crédito o aprobación de clientes) y cómo se informan.",
"5.1":"Los datos se usan solo para las finalidades declaradas al titular, sin usos incompatibles posteriores. Verificar: finalidades del aviso frente a usos reales en los sistemas.",
"5.2":"Solo se recolectan y mantienen los datos estrictamente necesarios para la finalidad. Verificar: formularios con campos justificados y revisión de campos innecesarios.",
"5.3":"Existen mecanismos para mantener los datos exactos y actualizados. Verificar: cómo se corrigen datos y con qué frecuencia se revisan.",
"5.4":"Existe una política de retención con plazos concretos para cada categoría de dato. Verificar: tabla de retención con plazos y base legal.",
"5.5":"Al cumplirse el plazo, los datos se eliminan o anonimizan de forma segura, y queda evidencia. Verificar: registro de una eliminación reciente.",
"5.6":"Los datos obtenidos de terceros o de fuentes públicas se validan en cuanto a su origen y licitud antes de usarlos. Verificar: origen documentado de cada base comprada o recibida.",
"6.1":"Los datos sensibles (salud, biométricos, origen, creencias, antecedentes penales) y los de menores están identificados en el inventario. Verificar: lista de datos sensibles por sistema.",
"6.2":"Para tratar datos sensibles se obtiene consentimiento explícito del titular. Verificar: formulario de consentimiento específico para datos sensibles.",
"6.3":"Los datos sensibles tienen controles reforzados: acceso restringido, cifrado y registro de accesos. Verificar: matriz de acceso y logs de esos datos.",
"6.4":"Si se tratan datos de niños, niñas o adolescentes, se aplica el interés superior del menor y se obtiene el consentimiento de su representante cuando corresponde. Verificar: si hay tratamientos con menores y cómo se controlan.",
"6.5":"El tratamiento biométrico y la vigilancia sistemática tienen restricciones y justificación escrita. Verificar: si existe biometría o vigilancia y su justificación.",
"7.1":"Están definidos criterios de alto riesgo y se comunican a las áreas que proponen nuevos tratamientos. Verificar: criterios escritos y su difusión.",
"7.2":"Se realiza una evaluación de impacto en protección de datos (EIPD/DPIA) antes de iniciar cada tratamiento de alto riesgo. Verificar: EIPD existentes y su fecha.",
"7.3":"La privacidad se revisa como parte de la gestión de proyectos y de los cambios a sistemas. Verificar: checklist de privacidad en la metodología de proyectos.",
"7.4":"Por defecto, los sistemas y productos tratan el mínimo de datos y dan el mínimo de accesos. Verificar: configuración por defecto en formularios y sistemas.",
"7.5":"Si el riesgo residual de un tratamiento es alto, existe un procedimiento para consultar a la autoridad antes de iniciarlo. Verificar: procedimiento escrito.",
"8.1":"Existe un sistema de gestión de seguridad de la información (ISO/IEC 27001) implementado o en implementación. Verificar: alcance del SGSI y su avance.",
"8.2":"El acceso a los sistemas sigue el principio de mínimo privilegio, con separación de funciones y revisión periódica. Verificar: matriz de accesos y última revisión.",
"8.3":"Todos los sistemas que contienen datos personales exigen autenticación multifactor (MFA). Verificar: lista de sistemas con MFA activo y los que no lo tienen.",
"8.4":"Los datos se cifran en tránsito y en reposo, incluidos respaldos y equipos portátiles. Verificar: configuración de cifrado por sistema y en equipos.",
"8.5":"Las vulnerabilidades se corrigen con plazos según su criticidad. Verificar: último escaneo, hallazgos abiertos y plazos de parcheo.",
"8.6":"Los accesos a datos personales quedan registrados, los registros están protegidos y tienen plazo de retención. Verificar: logs de accesos y su retención.",
"8.7":"Los respaldos están cifrados y se prueban restaurándolos. Verificar: última prueba de restauración documentada.",
"8.8":"Se realizan pruebas de intrusión o evaluaciones técnicas al menos una vez al año, y los hallazgos críticos se cierran. Verificar: informe de pentest del último año.",
"8.9":"Los datos se seudonimizan o anonimizan cuando la finalidad lo permite. Verificar: casos donde se aplica y donde podría aplicarse.",
"8.10":"Las cuentas privilegiadas (administradores) tienen control reforzado y sus sesiones quedan registradas. Verificar: lista de cuentas privilegiadas y herramienta de control.",
"8.11":"Los equipos de usuario tienen MDM, cifrado y EDR. Verificar: cobertura de cada herramienta sobre el parque de equipos.",
"8.12":"Servidores y equipos tienen protección contra malware actualizada y monitoreada. Verificar: porcentaje de equipos con EDR y alertas recientes.",
"8.13":"El acceso al código fuente y a la información está restringido por rol. Verificar: permisos del repositorio y de carpetas críticas.",
"8.14":"Existen controles para evitar la fuga de datos (envíos masivos, copias a dispositivos externos). Verificar: reglas de DLP o equivalentes y su alerta.",
"8.15":"Se controla la capacidad de los servicios y la sincronización horaria de los sistemas. Verificar: monitoreo de capacidad y configuración NTP.",
"8.16":"Las redes están segmentadas y protegidas con firewalls y servicios de red seguros. Verificar: diagrama de red y reglas de firewall.",
"8.17":"Las herramientas de administración con privilegios están restringidas y solo se instala software autorizado. Verificar: lista de herramientas privilegiadas y de software permitido.",
"8.18":"Las identidades y credenciales de usuarios, servicios y equipos se gestionan con procedimiento y registro. Verificar: proceso de alta y baja con tickets.",
"8.19":"Los accesos verifican la autenticación de cada usuario y protegen las aserciones de identidad. Verificar: configuración de autenticación y tokens.",
"8.20":"Los permisos se asignan por mínimo privilegio y se revisan periódicamente. Verificar: última revisión de accesos (recertificación).",
"8.21":"Los respaldos están protegidos, separados del sistema principal y resguardados. Verificar: ubicación de los respaldos y control de acceso.",
"8.22":"Las aplicaciones se desarrollan con requisitos de seguridad, codificación segura y pruebas antes de salir a producción. Verificar: procedimiento de desarrollo y evidencia de pruebas.",
"8.23":"Si se externaliza desarrollo, el proveedor cumple requisitos de seguridad, y los entornos de desarrollo, prueba y producción están separados. Verificar: contratos y diagrama de entornos.",
"8.24":"Los cambios a sistemas en producción se aprueban, prueban y registran antes de implementarse. Verificar: registro de cambios de los últimos meses.",
"8.25":"Existe un modelo RBAC documentado (roles, perfiles y matriz de accesos) revisado periódicamente. Verificar: matriz de roles y fecha de revisión.",
"8.26":"El acceso al entorno de pagos exige autenticación fuerte según PCI DSS (requisitos 8). Verificar: MFA en acceso a entornos de pago.",
"8.27":"Los equipos del entorno de pagos tienen antivirus o EDR actualizado y monitoreado (PCI DSS requisito 5). Verificar: cobertura y última alerta.",
"9.1":"Existe un procedimiento de respuesta a incidentes con roles y responsabilidades asignados. Verificar: procedimiento y nombres de responsables.",
"9.2":"Existe un mecanismo para detectar y registrar vulneraciones que afecten datos personales. Verificar: fuentes de detección (alertas, reportes del personal) y registro.",
"9.3":"Las vulneraciones que afectan datos personales se notifican a la autoridad dentro del plazo legal. Verificar: plazo exacto según el texto oficial y el procedimiento que lo cumple.",
"9.4":"Cuando hay riesgo elevado para los titulares, se les notifica sin demora. Verificar: criterio de riesgo y canal de notificación.",
"9.5":"Existen plantillas listas de notificación a la autoridad y a los titulares. Verificar: plantillas vigentes.",
"9.6":"Se realiza al menos una vez al año un simulacro de brecha de datos personales. Verificar: fecha y resultados del último simulacro.",
"9.7":"Cada vulneración queda registrada con causa raíz, impacto y acciones correctivas. Verificar: bitácora de incidentes.",
"9.8":"Existe un plan de continuidad y recuperación ante desastres con RTO y RPO definidos y probados. Verificar: plan, RTO y RPO por sistema y resultado de la última prueba.",
"9.9":"Los incidentes se clasifican, priorizan y escalan con criterios escritos. Verificar: matriz de severidad y reglas de escalamiento.",
"9.10":"Se estima la magnitud de cada incidente (datos y personas afectadas) y se recolecta la evidencia. Verificar: plantilla de evaluación y de evidencia.",
"10.1":"Existe un inventario de los proveedores que tratan datos personales por cuenta de la organización. Verificar: listado con tipo de datos tratados.",
"10.2":"Los contratos con proveedores que tratan datos incluyen las cláusulas exigidas: objeto, duración, obligaciones, subencargo y eliminación al término. Verificar: muestra de contratos.",
"10.3":"Los proveedores se evalúan en seguridad y privacidad antes de contratar y periódicamente. Verificar: cuestionario de evaluación y fecha.",
"10.4":"Los contratos permiten auditar al proveedor. Verificar: cláusula de auditoría en contratos.",
"10.5":"Las transferencias de datos fuera de Chile solo se hacen a países con nivel adecuado o con garantías contractuales. Verificar: hoja Transferencias completa (indicador automático).",
"10.6":"Los terceros que acceden a información firman acuerdos de confidencialidad. Verificar: acuerdos firmados.",
"10.7":"Los servicios en la nube se evalúan y se gestionan con criterios de seguridad y privacidad. Verificar: inventario de servicios cloud y su evaluación.",
"10.8":"Los sub-encargados (proveedores de tus proveedores) están identificados y autorizados por contrato. Verificar: lista de sub-encargados por proveedor.",
"10.9":"Al terminar un contrato, el proveedor devuelve o elimina los datos y entrega certificado. Verificar: cláusula y ejemplo de cierre de contrato.",
"11.1":"Todo el personal recibe capacitación en protección de datos al ingresar y al menos una vez al año. Verificar: registro de asistencia y porcentaje de cumplimiento.",
"11.2":"La capacitación tiene contenidos según el rol (TI, ventas, RRHH, atención al cliente, dirección). Verificar: programa por rol.",
"11.3":"Se realizan campañas de concienciación y simulacros de phishing con métricas. Verificar: resultados del último simulacro.",
"11.4":"Todo el personal y los contratistas firman compromisos de confidencialidad. Verificar: firmas en expedientes o contratos.",
"11.5":"El proceso de selección y los contratos de trabajo incluyen cláusulas de privacidad. Verificar: contrato tipo con la cláusula.",
"11.6":"Existe un proceso disciplinario para quienes incumplen las políticas de datos. Verificar: reglamento interno y casos aplicados.",
"11.7":"Al terminar o cambiar de cargo, se revocan accesos y se devuelven activos con registro. Verificar: checklist de salida de los últimos casos.",
"11.8":"El trabajo remoto y el uso aceptable de la información tienen reglas escritas. Verificar: política de trabajo remoto.",
"11.9":"La contratación, evaluación y salida de personal incluyen requisitos de ciberseguridad y privacidad. Verificar: procesos de RRHH con esos requisitos.",
"11.10":"Existen filtros de correo, MFA y simulacros de phishing con métricas. Verificar: herramientas usadas y resultados.",
"12.1":"Existe un plan anual de auditoría interna de cumplimiento aprobado. Verificar: plan aprobado y fecha.",
"12.2":"Se miden y reportan a la dirección indicadores de privacidad: solicitudes ARCOP, incidentes, capacitación y EIPD. Verificar: último reporte a la dirección.",
"12.3":"Los hallazgos de auditoría tienen plan de acción, responsable y plazo. Verificar: seguimiento de hallazgos abiertos.",
"12.4":"La política y el programa de cumplimiento se revisan al menos una vez al año. Verificar: fecha de última revisión.",
"12.5":"Se monitorean cambios en leyes, reglamentos y criterios de la autoridad de control. Verificar: quién monitorea y cómo se comunica.",
"12.6":"Las auditorías se realizan sin afectar la operación ni los datos de producción. Verificar: procedimiento de auditoría.",
"13.1":"Las vulnerabilidades de los sistemas que tratan datos personales se identifican, validan y registran. Verificar: último escaneo y registro de hallazgos.",
"13.2":"Existe un canal y un proceso para recibir y responder reportes de vulnerabilidades de terceros. Verificar: canal publicado y responsable.",
"13.3":"Los sistemas con datos personales tienen configuración segura documentada. Verificar: configuraciones base y su revisión.",
"13.4":"El software y hardware sin soporte del fabricante se retiran o se aíslan si contienen datos personales. Verificar: inventario de equipos y software sin soporte.",
"13.5":"El software no autorizado se bloquea y se registra. Verificar: lista de software permitido y alertas de instalación.",
"13.6":"El correo y la navegación están protegidos contra phishing y sitios maliciosos. Verificar: filtros activos.",
"13.7":"Todos los endpoints y servidores tienen defensas contra malware. Verificar: cobertura de antivirus o EDR.",
"13.8":"Las acciones y datos de los incidentes quedan registrados con integridad y origen. Verificar: registro de un incidente reciente.",
"13.9":"La evidencia electrónica se preserva según la Ley 21.459 (arts. 14 y 18 N°1) dentro del procedimiento de incidentes. Verificar: procedimiento de preservación.",
"13.10":"Los incidentes con datos personales se contienen y se erradican con pasos documentados. Verificar: procedimiento y último caso.",
"13.11":"Los respaldos se verifican antes de restaurar datos personales. Verificar: procedimiento de verificación y última restauración.",
"13.12":"Los cambios tienen evaluación de impacto en privacidad y aceptación formal (COBIT BAI06 y BAI07). Verificar: formulario de cambio con esos campos.",
"13.13":"Los activos y las configuraciones se gestionan durante todo su ciclo de vida (COBIT BAI09 y BAI10). Verificar: inventario y registro de configuración.",
"13.14":"Los sistemas con datos personales siguen los benchmarks de configuración aplicables (sistemas operativos, bases de datos, redes, aplicaciones, nube y suite de oficina). Verificar: benchmark usado y cumplimiento.",
"13.15":"El entorno de pagos cumple PCI DSS requisitos 2 y 5 (configuración segura y malware). Verificar: aplica solo si la organización procesa tarjetas; indicar N/A si no.",
"13.16":"El entorno de pagos cumple PCI DSS requisitos 7 y 8 (acceso y autenticación). Verificar: aplica solo si procesa tarjetas; indicar N/A si no.",
"13.17":"El modelo de prevención de delitos de la Ley 20.393 (art. 4) está integrado con el programa de datos personales. Verificar: modelo vigente y su vínculo con privacidad.",
"13.18":"El programa de cumplimiento está alineado con la Política Nacional de Ciberseguridad. Verificar: mapeo de la política con el programa.",
"13.19":"Los contenedores y su orquestación siguen benchmarks de configuración. Verificar: aplica solo si hay contenedores; indicar N/A si no.",
"13.20":"El alcance del entorno de pagos (tiendas, bodegas y ecommerce) está identificado y segmentado. Verificar: diagrama de alcance.",
"14.1":"Las áreas sensibles tienen perímetro físico y control de entrada. Verificar: tarjetas, biométrico o libro de ingreso.",
"14.2":"Oficinas, recintos y áreas seguras están protegidos. Verificar: políticas de acceso y revisión de instalaciones.",
"14.3":"Existe monitoreo físico (CCTV, alarmas) y protección frente a incendio, inundación u otras amenazas. Verificar: sistemas instalados y su mantenimiento.",
"14.4":"Hay reglas de escritorio limpio y pantalla bloqueada, y los equipos están ubicados para evitar accesos indebidos. Verificar: revisión de puestos de trabajo.",
"14.5":"Los equipos y datos fuera de la oficina están protegidos, los medios removibles están controlados y la eliminación de equipos es segura. Verificar: procedimiento y certificados de destrucción.",
"14.6":"Energía, cableado y mantenimiento de equipos están protegidos y registrados. Verificar: UPS, planos de cableado y registro de mantenimientos.",
"15.1":"Existe inventario y clasificación de servidores, equipos de red y servicios en la nube. Verificar: inventario con criticidad.",
"15.2":"Los sistemas operativos tienen configuración segura según un benchmark reconocido. Verificar: benchmark usado y resultado.",
"15.3":"La red está segmentada por nivel de sensibilidad (VLAN, zonas y reglas de firewall). Verificar: diagrama y reglas.",
"15.4":"Las bases de datos tienen configuración segura y el acceso privilegiado está controlado. Verificar: inventario de bases y cuentas administradoras.",
"15.5":"Las bases de datos y los datos personales en reposo están cifrados. Verificar: configuración de cifrado por base.",
"15.6":"Las consultas y cambios sobre datos personales en bases de datos quedan registradas. Verificar: auditoría activa y retención.",
"15.7":"Los datos personales se enmascaran o anonimizan en entornos de prueba. Verificar: entornos de prueba y datos usados.",
"15.8":"Las comunicaciones externas usan TLS 1.2 o superior y el acceso remoto usa VPN. Verificar: configuración de los servicios expuestos y de la VPN.",
"15.9":"El correo y la navegación están filtrados, con autenticación de dominio (SPF, DKIM, DMARC) y bloqueo de sitios riesgosos. Verificar: configuración del dominio y del filtro.",
"15.10":"Los servicios en nube y los contenedores tienen configuración segura. Verificar: revisión de configuración en la nube.",
"15.11":"Las plataformas de productividad (por ejemplo, Microsoft 365) están configuradas según benchmark. Verificar: configuración de seguridad de la suite.",
"15.12":"Las bases de datos se respaldan y se prueba su restauración. Verificar: última prueba documentada.",
"15.13":"Los entornos de pago (ecommerce, POS y bodegas) tienen configuración segura (PCI DSS 2 y 5). Verificar: aplica solo si procesa tarjetas.",
"15.14":"La recuperación ante desastres tiene RTO y RPO probados para bases de datos y sistemas críticos. Verificar: valores definidos y última prueba.",
"16.1":"Los cambios se evalúan por impacto en privacidad antes de aprobarse (COBIT BAI06). Verificar: formulario de cambio.",
"16.2":"Los cambios se aceptan formalmente antes de pasar a producción (COBIT BAI07). Verificar: aprobación registrada.",
"16.3":"Los activos de TI se gestionan durante todo su ciclo de vida (COBIT BAI09). Verificar: inventario con estado y responsable.",
"16.4":"La configuración de los elementos de TI se registra y se mantiene actualizada (COBIT BAI10). Verificar: registro de configuración (CMDB u otro).",
"16.5":"El riesgo de TI se integra con la gestión de riesgos del SGSI y se reporta a la dirección. Verificar: último reporte de riesgo.",
"17.1":"Los candidatos reciben un aviso de privacidad y su CV y entrevistas se usan solo para la selección. Verificar: aviso en el proceso de postulación.",
"17.2":"Los CV de candidatos no contratados se eliminan al cumplirse un plazo definido. Verificar: plazo y evidencia de eliminación.",
"17.3":"Las evaluaciones psicológicas, tests y perfilamiento de candidatos tienen base legal y decisiones revisadas por una persona. Verificar: qué tests se usan y quién decide.",
"17.4":"Los contratos de trabajo y sus anexos incluyen cláusulas de privacidad y confidencialidad. Verificar: contrato tipo y anexos.",
"17.5":"En el onboarding se entregan las políticas y se asignan accesos según el rol aprobado. Verificar: checklist de ingreso.",
"17.6":"En el offboarding se revocan accesos, se devuelven activos y se cierra el expediente. Verificar: checklist de salida con fechas.",
"17.7":"El expediente de cada empleado tiene acceso restringido, plazo de retención y eliminación al vencer. Verificar: quién accede al expediente y plazo aplicado.",
"17.8":"Los datos de salud (licencias médicas, certificados) están separados del expediente general y con acceso restringido. Verificar: dónde se guardan y quién los ve.",
"17.9":"Las evaluaciones de desempeño y de otro tipo tienen finalidad declarada, y el titular puede acceder a su resultado. Verificar: política de evaluaciones.",
"17.10":"Los registros de capacitación y cursos tienen retención y acceso definidos. Verificar: dónde se guardan y por cuánto tiempo.",
"17.11":"El control de asistencia y acceso con huella o reconocimiento facial trata un dato biométrico: exige consentimiento explícito y medidas reforzadas. Verificar: si el control usa biometría y, de ser así, el consentimiento obtenido.",
"17.12":"La videovigilancia en tiendas y bodegas tiene aviso visible, finalidad definida y plazo de retención. Verificar: avisos instalados y plazo de grabación.",
"17.13":"El monitoreo del correo y los equipos de trabajo está informado al personal y es proporcional. Verificar: política de monitoreo firmada.",
"17.14":"La transferencia de datos de nómina a previsión y salud tiene base legal documentada. Verificar: convenio o base legal de cada transferencia.",
"18.1":"El origen de las bases de datos de clientes, leads y prospectos (propias o compradas) está documentado y es lícito. Verificar: origen y contrato de cada base comprada.",
"18.2":"El marketing a personas naturales requiere consentimiento; la prospección B2B tiene base legal documentada. Verificar: formularios y base legal del envío comercial.",
"18.3":"Las cookies y el seguimiento web requieren consentimiento previo, y los terceros que rastrean están informados. Verificar: banner de cookies y lista de terceros.",
"18.4":"Los datos de compras y pagos en ecommerce tienen control de acceso y cifrado. Verificar: configuración de la plataforma de venta.",
"18.5":"El perfilamiento y la segmentación de clientes se informan al titular y, si aplica, tienen evaluación de impacto. Verificar: si hay perfilamiento y cómo se informa.",
"18.6":"Los programas de fidelización y puntos tienen finalidad declarada y consentimiento. Verificar: términos del programa.",
"18.7":"El opt-out de marketing se atiende de inmediato y queda registrado. Verificar: prueba de desuscripción y registro.",
"18.8":"Los leads y clientes inactivos se eliminan o anonimizan al cumplir un plazo definido. Verificar: plazo y última depuración.",
"18.9":"Las grabaciones de call center tienen aviso, finalidad y plazo de retención. Verificar: aviso al inicio de la llamada y plazo configurado.",
"18.10":"El acceso al CRM está restringido por rol y se controlan las exportaciones masivas. Verificar: permisos del CRM y registro de exportaciones.",
}
READABLE_HELP=True

# Descripciones en español: qué se debe levantar para cada control (redacción propia, orientada a entrevista/revisión)
NAME_ES={"L21.663 Ley de Ciberseguridad":"Ley 21.663 Marco de ciberseguridad",
"NIST GV.OC-01":"Misión y contexto de la organización","NIST GV.OC-02":"Partes interesadas y sus expectativas","NIST GV.OC-03":"Obligaciones legales, regulatorias y contractuales","NIST GV.OC-04":"Expectativas de partes externas sobre la organización","NIST GV.OC-05":"Dependencias y servicios críticos de la organización",
"NIST GV.RR-01":"Liderazgo responsable del riesgo de ciberseguridad","NIST GV.RR-02":"Roles, responsabilidades y autoridades definidos","NIST GV.RR-03":"Recursos asignados a la ciberseguridad","NIST GV.RR-04":"Ciberseguridad incluida en RRHH",
"NIST GV.PO-01":"Política de gestión del riesgo establecida","NIST GV.PO-02":"Política revisada y actualizada",
"NIST ID.AM-01":"Inventario de hardware","NIST ID.AM-02":"Inventario de software y servicios","NIST ID.AM-03":"Mapa de flujos de datos y comunicaciones","NIST ID.AM-05":"Priorización de activos por criticidad","NIST ID.AM-08":"Gestión de activos a lo largo de su ciclo de vida",
"NIST ID.RA-01":"Identificación de vulnerabilidades","NIST ID.RA-08":"Recepción de divulgaciones de vulnerabilidades",
"NIST PR.AA-01":"Gestión de identidades y credenciales","NIST PR.AA-02":"Verificación de identidades","NIST PR.AA-03":"Autenticación de usuarios, servicios y equipos","NIST PR.AA-04":"Protección de aserciones de identidad","NIST PR.AA-05":"Permisos con mínimo privilegio y separación de funciones",
"NIST PR.DS-11":"Respaldos de datos","NIST PR.PS-01":"Gestión de configuración segura","NIST PR.PS-02":"Software mantenido o retirado según riesgo","NIST PR.PS-03":"Hardware mantenido o retirado según riesgo","NIST PR.PS-05":"Prevención de software no autorizado",
"NIST RS.MA-01":"Plan de respuesta ejecutado con terceros","NIST RS.MA-02":"Reportes de incidentes clasificados y validados","NIST RS.MA-03":"Incidentes categorizados y priorizados","NIST RS.MA-04":"Escalamiento de incidentes","NIST RS.MA-05":"Criterios para iniciar la recuperación",
"NIST RS.AN-03":"Análisis de causa raíz","NIST RS.AN-06":"Registro de acciones durante la investigación","NIST RS.AN-07":"Recolección de datos del incidente","NIST RS.AN-08":"Estimación de la magnitud del incidente",
"NIST RS.CO-02":"Notificación a partes interesadas","NIST RS.CO-03":"Información compartida con partes designadas","NIST RS.MI-01":"Contención de incidentes","NIST RS.MI-02":"Erradicación de incidentes",
"NIST RC.RP-03":"Verificación de integridad de respaldos antes de restaurar",
"CIS 1":"Inventario de activos empresariales","CIS 2":"Inventario de software","CIS 4":"Configuración segura de activos y software","CIS 5":"Gestión de cuentas","CIS 6":"Gestión del control de acceso","CIS 7":"Gestión continua de vulnerabilidades","CIS 9":"Protección de correo y navegador","CIS 10":"Defensas contra malware","CIS 11":"Recuperación de datos","CIS 17":"Gestión de respuesta a incidentes",
"BM-OS":"Benchmark de sistemas operativos","BM-CLOUD":"Benchmark de proveedores de nube","BM-DB":"Benchmark de bases de datos","BM-NET":"Benchmark de dispositivos de red","BM-CONT":"Benchmark de contenedores y orquestación","BM-APP":"Benchmark de aplicaciones y navegadores","BM-M365":"Benchmark de suites de productividad",
"BAI06":"Gestión de cambios","BAI07":"Aceptación del cambio y de la transición","BAI09":"Gestión de activos","BAI10":"Gestión de la configuración",
"PCI Req 2":"Configuraciones seguras en todos los componentes","PCI Req 5":"Protección de sistemas contra software malicioso","PCI Req 7":"Acceso restringido según necesidad de saber","PCI Req 8":"Identificación de usuarios y autenticación del acceso",
"L21.719 Art.14 sexies":"Deber de reportar vulneraciones de seguridad","L21.719 Art.49":"Modelo de prevención de infracciones","L21.719 Art.50":"Atribuciones del delegado de protección de datos","L21.719 Programa adecuación":"Programa de adecuación de la organización",
"L21.459 Art.14":"Preservación y custodia de evidencia electrónica","L21.459 Art.18 N1":"Preservación provisoria de datos","L20.393 Art.4":"Modelo de prevención de delitos vigente","L20.393 Art.4 N3":"Encargado de prevención de delitos","L20.393 Art.4 N2":"Protocolos y procedimientos de control","PNCS":"Política Nacional de Ciberseguridad",
}
DESC={
"NIST GV.OC-01":"Verificar que la misión y los objetivos del negocio se conocen y se usan para decidir el nivel de riesgo de privacidad y seguridad.",
"NIST GV.OC-02":"Verificar que están identificados clientes, empleados, reguladores y proveedores, y qué esperan de la organización en privacidad y seguridad.",
"NIST GV.OC-03":"Verificar que existe una lista de leyes y contratos aplicables (incluida la Ley 21.719) y que alguien es responsable de mantenerla.",
"NIST GV.OC-04":"Verificar que los compromisos con clientes y terceros (por ejemplo, de seguridad de pagos) están documentados y comunicados.",
"NIST GV.OC-05":"Verificar que se conocen los servicios de los que depende la organización (nube, pagos, logística) y su criticidad.",
"NIST GV.RR-01":"Verificar que la dirección asume la responsabilidad del riesgo de seguridad y privacidad (actas, reportes a directorio).",
"NIST GV.RR-02":"Verificar que los roles de seguridad y privacidad están definidos, comunicados y que sus responsables saben qué deben hacer.",
"NIST GV.RR-03":"Verificar que el presupuesto y el personal para seguridad y privacidad son suficientes para el alcance real.",
"NIST GV.RR-04":"Verificar que la contratación, la evaluación y la salida de personal incluyen requisitos de seguridad y privacidad.",
"NIST GV.PO-01":"Verificar que existe una política de gestión de riesgo de ciberseguridad aprobada, difundida y que se aplica.",
"NIST GV.PO-02":"Verificar la fecha de la última revisión de la política y que considera cambios legales y tecnológicos recientes.",
"NIST ID.AM-01":"Verificar que existe un listado de equipos (servidores, PC, POS, impresoras) con responsable y ubicación.",
"NIST ID.AM-02":"Verificar que existe un listado de software y servicios en uso, incluidos los de nube y los no autorizados.",
"NIST ID.AM-03":"Verificar que hay un diagrama de flujos de datos personales dentro de la organización y hacia terceros.",
"NIST ID.AM-05":"Verificar que los activos se priorizan (crítico, alto, medio, bajo) y que la prioridad guía los controles.",
"NIST ID.AM-08":"Verificar que los sistemas se gestionan desde su compra hasta su baja, con responsables en cada etapa.",
"NIST ID.RA-01":"Verificar que se realizan escaneos de vulnerabilidades periódicos y que los hallazgos se registran y validan.",
"NIST ID.RA-08":"Verificar que existe un canal público o interno para que terceros reporten vulnerabilidades y un proceso para responderlas.",
"NIST PR.AA-01":"Verificar que el alta, cambio y baja de usuarios, servicios y equipos se hacen con procedimiento y registro.",
"NIST PR.AA-02":"Verificar que la identidad de cada usuario se confirma antes de entregar acceso (por ejemplo, con cédula o contrato).",
"NIST PR.AA-03":"Verificar que todos los accesos exigen autenticación y que las cuentas de servicio también tienen control.",
"NIST PR.AA-04":"Verificar que las sesiones y tokens de autenticación se transmiten y almacenan de forma protegida (cifrados, con expiración).",
"NIST PR.AA-05":"Verificar que los permisos se asignan por rol, se revisan periódicamente y no se acumulan sin control.",
"NIST PR.DS-11":"Verificar que existen respaldos periódicos, cifrados, almacenados fuera del sistema principal y con prueba de restauración.",
"NIST PR.PS-01":"Verificar que existen configuraciones base (hardening) para servidores, PC y bases de datos, y que se aplican.",
"NIST PR.PS-02":"Verificar que el software tiene soporte del fabricante y que el software sin soporte se retira o aísla.",
"NIST PR.PS-03":"Verificar que los equipos sin soporte o con falla se reemplazan según plan y que el retiro es seguro.",
"NIST PR.PS-05":"Verificar que los usuarios no pueden instalar software sin autorización y que existe un listado de software permitido.",
"NIST RS.MA-01":"Verificar que el plan de respuesta incluye a proveedores clave (hosting, pagos, SOC) y sus contactos.",
"NIST RS.MA-02":"Verificar que cada reporte de incidente se clasifica y se confirma antes de actuar.",
"NIST RS.MA-03":"Verificar que los incidentes tienen categoría y prioridad según criterios escritos.",
"NIST RS.MA-04":"Verificar que existen reglas claras de cuándo escalar un incidente a gerencia, legal o al responsable de privacidad.",
"NIST RS.MA-05":"Verificar que hay criterios escritos para decidir cuándo se da por terminada la contención y se inicia la recuperación.",
"NIST RS.AN-03":"Verificar que los incidentes importantes tienen un informe de causa raíz con fecha y responsable.",
"NIST RS.AN-06":"Verificar que las acciones durante un incidente quedan registradas (quién, qué, cuándo) sin posibilidad de alteración.",
"NIST RS.AN-07":"Verificar que se conserva la información del incidente (logs, capturas, correos) con cadena de custodia.",
"NIST RS.AN-08":"Verificar que se estima cuántos datos y personas afectó un incidente y que la estimación se valida.",
"NIST RS.CO-02":"Verificar que están definidos quiénes deben ser notificados de un incidente (autoridad, titulares, clientes) y cómo.",
"NIST RS.CO-03":"Verificar que la información compartida durante un incidente se limita a quienes la necesitan.",
"NIST RS.MI-01":"Verificar que existen pasos para aislar sistemas o cuentas comprometidas en un incidente y que se han probado.",
"NIST RS.MI-02":"Verificar que, tras un incidente, se elimina la causa (malware, cuenta, vulnerabilidad) antes de restablecer el servicio.",
"NIST RC.RP-03":"Verificar que los respaldos se validan antes de restaurar datos, para no restaurar archivos dañados o infectados.",
"CIS 1":"Verificar que existe un inventario de todos los equipos conectados a la red, actualizado y con responsable.",
"CIS 2":"Verificar que existe un inventario de software instalado y que se controla lo que no está autorizado.",
"CIS 4":"Verificar que los equipos y software tienen configuración segura documentada y revisada.",
"CIS 5":"Verificar que las cuentas de usuario, administrador y servicio se inventarian, se revisan y se desactivan cuando no se usan.",
"CIS 6":"Verificar que el acceso a sistemas y datos se otorga según el rol y se revisa periódicamente.",
"CIS 7":"Verificar que se escanea la organización periódicamente y que los hallazgos críticos se corrigen en plazo definido.",
"CIS 9":"Verificar que el correo y la navegación tienen filtros contra phishing, malware y sitios maliciosos.",
"CIS 10":"Verificar que todos los equipos tienen antivirus o EDR activo, actualizado y monitoreado.",
"CIS 11":"Verificar que existen respaldos periódicos de datos críticos y que se prueba la recuperación.",
"CIS 17":"Verificar que existe un plan de respuesta a incidentes, roles asignados y ejercicios realizados.",
"BM-OS":"Verificar que los servidores y equipos usan una configuración base de un estándar reconocido (por ejemplo, CIS Benchmark) y que se revisa.",
"BM-CLOUD":"Verificar que las cuentas y servicios en la nube (AWS, Azure, Google u otros) están configurados según un estándar y revisados.",
"BM-DB":"Verificar que las bases de datos tienen configuración segura: usuarios, cifrado, auditoría y puertos restringidos.",
"BM-NET":"Verificar que firewalls, switches y routers tienen contraseñas, accesos y servicios configurados según un estándar.",
"BM-CONT":"Verificar que los contenedores y la orquestación (Docker, Kubernetes) usan imágenes y configuraciones seguras. Aplica solo si hay contenedores.",
"BM-APP":"Verificar que las aplicaciones y navegadores corporativos tienen configuración segura y actualizada.",
"BM-M365":"Verificar que las suites de oficina y correo (Microsoft 365, Google Workspace) tienen MFA, reglas de compartición y auditoría activas.",
"BAI06":"Verificar que los cambios a sistemas se solicitan, evalúan (incluido el impacto en datos personales), aprueban y registran.",
"BAI07":"Verificar que un cambio se prueba y se aprueba formalmente antes de pasar a producción, con evidencia.",
"BAI09":"Verificar que los activos de TI tienen registro, responsable y se dan de baja de forma controlada.",
"BAI10":"Verificar que existe un registro de la configuración de sistemas y que se mantiene actualizado.",
"PCI Req 2":"Verificar que los POS, servidores y equipos del entorno de pagos usan configuraciones seguras y sin valores por defecto.",
"PCI Req 5":"Verificar que los equipos del entorno de pagos tienen protección contra malware actualizada y monitoreada.",
"PCI Req 7":"Verificar que el acceso a datos de pago se limita a quien lo necesita para su trabajo, con revisión periódica.",
"PCI Req 8":"Verificar que cada usuario tiene identificación única y que el acceso al entorno de pago exige autenticación fuerte.",
"L21.719 Art.14 sexies":"Verificar que existe un procedimiento para avisar a la autoridad de vulneraciones que afecten datos personales, con plazo definido según el texto oficial.",
"L21.719 Art.49":"Verificar que existe un modelo documentado para prevenir infracciones (políticas, controles, capacitación y monitoreo).",
"L21.719 Art.50":"Verificar que el delegado de protección tiene funciones, autoridad y acceso a la dirección definidos por escrito.",
"L21.719 Programa adecuación":"Verificar que existe un plan de adecuación con responsables, fechas y avance a la fecha.",
"L21.459 Art.14":"Verificar que la evidencia electrónica de incidentes se conserva íntegra y con custodia.",
"L21.459 Art.18 N1":"Verificar que se puede preservar de forma provisoria datos cuando lo requiera una autoridad, sin borrarlos.",
"L20.393 Art.4":"Verificar que existe un modelo de prevención de delitos (incluidos delitos informáticos y de datos) con políticas y controles.",
"L20.393 Art.4 N3":"Verificar que existe una persona designada como encargado de prevención, con autoridad y reporte a la dirección.",
"L20.393 Art.4 N2":"Verificar que existen protocolos y procedimientos que controlan las actividades de riesgo (por ejemplo, pagos y compras).",
"PNCS":"Verificar que la organización conoce la Política Nacional de Ciberseguridad y la considera en su programa.",
}
TEMA_ES={"GOVERN":"Gobernar (GOVERN)","IDENTIFY":"Identificar (IDENTIFY)","PROTECT":"Proteger (PROTECT)","RESPOND":"Responder (RESPOND)","RECOVER":"Recuperar (RECOVER)",
"Controles CIS":"Controles CIS v8.1","Benchmarks":"Benchmarks CIS v8.1","BAI":"Construir, adquirir e implementar (BAI)","Requisitos PCI":"Requisitos PCI DSS v4.0.1","Normativa":"Normativa chilena"}

ISO_NAME={"5.1": "Políticas de seguridad de la información", "5.2": "Roles y responsabilidades de SI", "5.3": "Segregación de funciones", "5.4": "Responsabilidades de la dirección", "5.5": "Contacto con autoridades", "5.6": "Contacto con grupos de interés especial", "5.7": "Inteligencia de amenazas", "5.8": "SI en gestión de proyectos", "5.9": "Inventario de información y activos", "5.10": "Uso aceptable de información", "5.11": "Devolución de activos", "5.12": "Clasificación de la información", "5.13": "Etiquetado de la información", "5.14": "Transferencia de información", "5.15": "Control de acceso", "5.16": "Gestión de identidades", "5.17": "Información de autenticación", "5.18": "Derechos de acceso", "5.19": "Seguridad en relaciones con proveedores", "5.20": "Seguridad en contratos con proveedores", "5.21": "Gestión de seguridad en cadena TIC", "5.22": "Monitoreo de proveedores", "5.23": "SI en servicios en la nube", "5.24": "Planificación y preparación de gestión de incidentes", "5.25": "Evaluación y decisión sobre eventos de SI", "5.26": "Respuesta a incidentes de SI", "5.27": "Aprendizaje de los incidentes", "5.28": "Recolección de evidencia", "5.29": "Seguridad de la información durante disrupción", "5.30": "Preparación TIC para continuidad", "5.31": "Requisitos legales, estatutarios y contractuales", "5.32": "Derechos de propiedad intelectual", "5.33": "Protección de registros", "5.34": "Privacidad y protección de datos personales", "5.35": "Revisión independiente de SI", "5.36": "Cumplimiento con políticas y normas de SI", "5.37": "Procedimientos de operación documentados", "6.1": "Selección de personal", "6.2": "Términos y condiciones de empleo", "6.3": "Concientización, educación y capacitación en SI", "6.4": "Proceso disciplinario", "6.5": "Responsabilidades ante terminación o cambio de empleo", "6.6": "Acuerdos de confidencialidad", "6.7": "Trabajo remoto", "6.8": "Reporte de eventos de SI", "7.1": "Perímetros de seguridad física", "7.2": "Controles de entrada física", "7.3": "Seguridad de oficinas, recintos e instalaciones", "7.4": "Monitoreo de seguridad física", "7.5": "Protección contra amenazas físicas y ambientales", "7.6": "Trabajo en áreas seguras", "7.7": "Escritorio limpio y pantalla limpia", "7.8": "Ubicación y protección de equipos", "7.9": "Seguridad de activos fuera de las instalaciones", "7.10": "Medios de almacenamiento", "7.11": "Servicios de suministro", "7.12": "Seguridad del cableado", "7.13": "Mantenimiento de equipos", "7.14": "Eliminación o reutilización segura de equipos", "8.1": "Dispositivos de usuario final", "8.2": "Derechos de acceso privilegiado", "8.3": "Restricción de acceso a la información", "8.4": "Acceso al código fuente", "8.5": "Autenticación segura", "8.6": "Gestión de capacidad", "8.7": "Protección contra malware", "8.8": "Gestión de vulnerabilidades técnicas", "8.9": "Gestión de la configuración", "8.10": "Eliminación de información", "8.11": "Enmascaramiento de datos", "8.12": "Prevención de fuga de datos", "8.13": "Respaldo de información", "8.14": "Redundancia de instalaciones de procesamiento", "8.15": "Registro de actividades (logging)", "8.16": "Actividades de monitoreo", "8.17": "Sincronización de relojes", "8.18": "Uso de programas de utilidad con privilegios", "8.19": "Instalación de software en sistemas operativos", "8.20": "Seguridad en redes", "8.21": "Seguridad de los servicios de red", "8.22": "Segregación de redes", "8.23": "Filtrado web", "8.24": "Uso de criptografía", "8.25": "Ciclo de vida de desarrollo seguro", "8.26": "Requisitos de seguridad en aplicaciones", "8.27": "Principios de arquitectura segura", "8.28": "Codificación segura", "8.29": "Pruebas de seguridad en desarrollo y aceptación", "8.30": "Desarrollo externalizado", "8.31": "Separación de entornos", "8.32": "Gestión de cambios", "8.33": "Información de prueba", "8.34": "Protección de sistemas de información en auditoría"}

IND_LABEL={"RAT":"Indicador automático: completitud del RAT (hoja RAT)",
"ARCOP":"Indicador automático: solicitudes ARCOP atendidas en plazo (hoja ARCOP)",
"TRANSF":"Indicador automático: transferencias con garantía (hoja Transferencias)"}
def _label(k):
    if k in IND_LABEL: return IND_LABEL[k]
    if k in ISO_NAME: return f"{k} · {ISO_NAME[k]}"
    if k.startswith("COBIT "): return f"COBIT {k[6:]} · proceso COBIT 5"
    if k in NAME_ES: return f"{k} · {NAME_ES[k]}"
    if k.startswith("NIST "): return f"{k} · subcategoría NIST CSF 2.0"
    return k
def READ(lst): return "\n".join("• "+_label(k) for k in lst)

wb=Workbook()

# ---------- Portada ----------
ws=wb.active; ws.title="Portada"
ws.column_dimensions["A"].width=3; ws.column_dimensions["B"].width=34; ws.column_dimensions["C"].width=90
ws["B2"]="CHECKLIST INTEGRADO DE CUMPLIMIENTO — LEY N° 21.719"; ws["B2"].font=Font(bold=True,size=14,color=NAVY)
ws["B3"]="Integra RAT, ARCOP, transferencias, ISO/IEC 27001:2022, NIST CSF 2.0, CIS, COBIT, PCI DSS y normativa chilena"; ws["B3"].font=Font(italic=True)
r=5
for k in ["Organización","Responsable de cumplimiento","Fecha de evaluación","Próxima revisión","Evaluador"]:
    ws.cell(r,2,k).font=Font(bold=True); c=ws.cell(r,3,"[Completar]"); c.fill=INF; r+=1
r+=1; ws.cell(r,2,"CÓMO FUNCIONA").font=Font(bold=True,color=NAVY); r+=1
steps=[
"1. Controles marcos: asignar a cada control (ISO, NIST, CIS, COBIT, PCI y normativa) un puntaje de 1 a 5, o N/A si no aplica.",
"2. RAT, ARCOP y Transferencias: mantener los registros. Los indicadores de la hoja Resumen se calculan solos desde ahí.",
"3. Requisitos Ley: el puntaje de cada requisito es el promedio de sus controles mapeados (y de sus indicadores, cuando corresponda). No se asigna puntaje manual.",
"4. Resumen: puntaje por dimensión, total ponderado, vista por marco y por función NIST / tema ISO.",
"5. Un requisito sin controles puntuados queda en blanco y no pesa en el total."]
for s in steps:
    ws.cell(r,2,s).alignment=WR; ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=3); ws.row_dimensions[r].height=30; r+=1

# ---------- Resumen (se crea ahora para fijar nombre; se llena después) ----------
wr=wb.create_sheet("Resumen")

# ---------- Requisitos Ley ----------
wq=wb.create_sheet("Requisitos Ley")
header(wq,1,["Dimensión","ID","Requisito: qué debe existir y cómo verificarlo","Controles e indicadores que alimentan el puntaje (el puntaje es su promedio)","Puntaje (1-5)","Evidencia encontrada (escriba aquí)","Brecha / plan de acción (escriba aquí)","Responsable","Fecha límite"],
       [11,8,70,58,12,40,36,16,13])
IND={"RAT":"Resumen!$C$4","ARCOP":"Resumen!$C$5","TRANSF":"Resumen!$C$6"}
text_of={}
for code,_,_ in DIMS:
    for iid,req in ITEMS[code]: text_of[iid]=(code,req)
q=2
for code,_,_ in DIMS:
    for iid,req in ITEMS[code]:
        lst=MAP[iid]
        refs=[IND[k] if k in IND else f"{CMN}!$E${ROW[key(k)]}" for k in lst]
        formula=f'=IFERROR(AVERAGE({",".join(refs)}),"")'
        wq.append([code,iid,REQ_TXT.get(iid,req),READ(lst),formula,None,None,None,None])
        q+=1
qlast=q-1
for rr in wq.iter_rows(min_row=2,max_row=qlast):
    for c in rr: c.alignment=WR; c.border=B
    rr[4].alignment=CT; rr[4].number_format="0.00"

EXTRA_IDS=set()
for _code in ["D13","D14","D16"]:
    for _iid,_ in ITEMS[_code]: EXTRA_IDS.add(_iid)
EXTRA_IDS |= {"8.26","8.27","11.10","15.2","15.10","15.11","15.13"}
for _c in ITEMS:
    for _iid,_t in ITEMS[_c]:
        if _t in GP_TEXTS: EXTRA_IDS.add(_iid)
for _r in range(2,qlast+1):
    _iid=wq.cell(_r,2).value
    _extra=_iid in EXTRA_IDS
    wq.cell(_r,10,"Extra (marco o buena práctica)" if _extra else "Obligatorio Ley 21.719").alignment=WR
    _k=wq.cell(_r,11,"No" if _extra else "Sí"); _k.fill=INF; _k.alignment=CT; from openpyxl.styles import Protection as _P; _k.protection=_P(locked=False)
    wq.cell(_r,12,f'=IF(AND(K{_r}="Sí",ISNUMBER(E{_r})),E{_r},"")').number_format="0.00"
    for _j in (10,11,12): wq.cell(_r,_j).border=B
for _j,(_h,_w) in enumerate([("Tipo de requisito",26),("Incluir en puntaje (Sí/No)",16),("Puntaje considerado (automático)",18)],10):
    _c=wq.cell(1,_j,_h); _c.font=H; _c.fill=HF; _c.alignment=CT; _c.border=B
    wq.column_dimensions[get_column_letter(_j)].width=_w

for _r in range(2,qlast+1): wq.row_dimensions[_r].height=75
wq.freeze_panes="D2"; wq.auto_filter.ref=f"A1:I{qlast}"

# ---------- Controles marcos ----------
wcm=wb.create_sheet("Controles marcos")
header(wcm,1,["Marco","Control","Nombre del control y qué debe levantar","Tema / función","Puntaje (1-5 / N/A)","Evidencia encontrada (escriba aquí)","Usado en requisitos"],
       [18,22,70,30,16,40,14])
for marco,k,desc,tema in ctrls:
    rr=ROW[k]
    wcm.cell(rr,1,marco); wcm.cell(rr,2,k); wcm.cell(rr,3,desc); wcm.cell(rr,4,tema)
    wcm.cell(rr,5).fill=INF; wcm.cell(rr,5).alignment=CT
    wcm.cell(rr,7,"Sí" if k in used else "No")
    for i in range(1,8): wcm.cell(rr,i).border=B; 
    for i in range(1,8): wcm.cell(rr,i).alignment=WR
clast=len(ctrls)+1
for _r in range(2,clast+1): wcm.row_dimensions[_r].height=62
dv=DataValidation(type="list",formula1='"1,2,3,4,5,N/A"',allow_blank=True); wcm.add_data_validation(dv); dv.add(f"E2:E{clast}")
wcm.conditional_formatting.add(f"E2:E{clast}",CellIsRule(operator="between",formula=["1","2"],fill=PatternFill("solid",fgColor="F8CBAD")))
wcm.conditional_formatting.add(f"E2:E{clast}",CellIsRule(operator="equal",formula=["3"],fill=PatternFill("solid",fgColor="FFE699")))
wcm.conditional_formatting.add(f"E2:E{clast}",CellIsRule(operator="between",formula=["4","5"],fill=PatternFill("solid",fgColor="C6E0B4")))
wcm.freeze_panes="C2"; wcm.auto_filter.ref=f"A1:G{clast}"

# ---------- RAT ----------
src=openpyxl.load_workbook('/mnt/project/registro_rat.xlsx')
def copy_sheet(ws_src, ws_dst):
    rows=[]
    for row in ws_src.iter_rows(values_only=True):
        if any(c is not None for c in row): rows.append(row)
    for row in rows: ws_dst.append(list(row))
    return rows
wrat=wb.create_sheet("RAT")
copy_sheet(src['RAT — Actividades'], wrat)
hdr_rat=[c.value for c in wrat[1]]
for c in wrat[1]: c.font=H; c.fill=HF; c.alignment=CT; c.border=B
CAMPOS_RAT_NUEVOS=['Fuente de los datos','Cómo se recolectan','Ubicación del almacenamiento','Forma de eliminación','Encargados (proveedores que tratan datos)','Decisiones automatizadas','Nivel de riesgo','Requiere evaluación de impacto (EIPD)','Derechos ARCOP que aplican','Estado de la actividad']
for _j,_h in enumerate(CAMPOS_RAT_NUEVOS, start=17):
    _c=wrat.cell(1,_j,_h); _c.font=H; _c.fill=HF; _c.alignment=CT; _c.border=B
wrat.cell(1,27,"Completitud (1 = todos los campos)").font=H
wrat.cell(1,27).fill=HF; wrat.cell(1,27).alignment=CT
for i,w in enumerate([10,28,34,24,30,14,28,30,20,22,28,26,34,22,20,14,16],1): wrat.column_dimensions[get_column_letter(i)].width=w
ratr=[]
for r in range(2,wrat.max_row+1):
    if str(wrat.cell(r,1).value or "").startswith("RAT-"):
        wrat.cell(r,27,f'=COUNTA(B{r}:Z{r})/25').number_format="0%"
        ratr.append(r)
rat_rng=f"RAT!$AA${min(ratr)}:$AA${max(ratr)}" if ratr else None
for row in wrat.iter_rows(min_row=2):
    for c in row: c.alignment=WR

# ---------- ARCOP ----------
warc=wb.create_sheet("ARCOP")
arows=copy_sheet(src['Derechos ARCOP'], warc)
dre=re.compile(r"^\d{2}/\d{2}/\d{4}$")
for row in warc.iter_rows(min_row=1):
    for c in row:
        if isinstance(c.value,str) and dre.match(c.value.strip()):
            c.value=datetime.datetime.strptime(c.value.strip(),"%d/%m/%Y"); c.number_format="dd/mm/yyyy"
header(warc,1,[c.value for c in warc[1]]+["Estado de plazo"],[10,12,22,14,16,14,16,50,18])
arc_rows=[]
for r in range(1,warc.max_row+1):
    if str(warc.cell(r,1).value or "").startswith("ARC-"):
        warc.cell(r,9,f'=IF(ISNUMBER(F{r}),IF(F{r}<=E{r},"A tiempo","Fuera de plazo"),IF(TODAY()>E{r},"Vencida","En plazo"))')
        arc_rows.append(r)
for row in warc.iter_rows(min_row=2):
    for c in row: c.alignment=WR

# ---------- Transferencias ----------
wtr=wb.create_sheet("Transferencias")
copy_sheet(src['Transferencias Internacionales'], wtr)
header(wtr,1,[c.value for c in wtr[1]]+["Garantía definida"],[12,26,12,28,26,30,36,30,30,18])
for r in range(2,wtr.max_row+1):
    if str(wtr.cell(r,1).value or "").startswith("RAT-"):
        wtr.cell(r,10,f'=IF(LEN(G{r})>0,"Sí","No")')
for row in wtr.iter_rows(min_row=2):
    for c in row: c.alignment=WR

# ---------- Resumen ----------
wr.column_dimensions["A"].width=12; wr.column_dimensions["B"].width=58
for col,w in zip("CDEFGH",[14,12,14,14,16,16]): wr.column_dimensions[col].width=w
wr["A1"]="RESUMEN INTEGRADO — LEY 21.719"; wr["A1"].font=Font(bold=True,size=13,color=NAVY)
wr["A3"]="INDICADORES DERIVADOS DEL LEVANTAMIENTO (escala 1–5)"; wr["A3"].font=Font(bold=True,color=NAVY)
rat_f=f'=IFERROR(1+4*AVERAGE({rat_rng}),"")' if rat_rng else '=""'
ind=[("Completitud del RAT (campos obligatorios por actividad)",rat_f),
("Derechos ARCOP atendidos en plazo",'=IFERROR(1+4*(COUNTIF(ARCOP!I:I,"A tiempo")+COUNTIF(ARCOP!I:I,"En plazo"))/COUNTIF(ARCOP!A:A,"ARC-*"),"")'),
("Transferencias con garantía definida",'=IFERROR(1+4*COUNTIF(Transferencias!J:J,"Sí")/COUNTIF(Transferencias!A:A,"RAT-*"),"")'),
("Solicitudes ARCOP vencidas (informativo, no suma al puntaje)",'=COUNTIF(ARCOP!I:I,"Vencida")')]
for i,(lab,f) in enumerate(ind):
    rr=4+i; wr.cell(rr,1,lab); wr.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=2)
    c=wr.cell(rr,3,f); c.number_format="0.00"; c.fill=PatternFill("solid",fgColor=LIGHT); c.border=B
wr["C7"].number_format="0"
hdr=["Código","Dimensión","Peso","Requisitos","Evaluados","Puntaje (1-5)","% cumplimiento","Nivel"]
for i,h in enumerate(hdr,1):
    c=wr.cell(9,i,h); c.font=H; c.fill=HF; c.alignment=CT; c.border=B
RQ="'Requisitos Ley'"
def nivel(cell): return f'=IF({cell}=0,"Sin evaluar",IF({cell}<2,"Crítico",IF({cell}<3,"Inicial",IF({cell}<4,"En desarrollo",IF({cell}<4.6,"Avanzado","Optimizado")))))'
first=10
for i,(code,name,w) in enumerate(DIMS):
    r=first+i
    wr.cell(r,1,code); wr.cell(r,2,name)
    c=wr.cell(r,3,w/100); c.number_format="0%"; c.fill=INF
    wr.cell(r,4,f'=COUNTIFS({RQ}!$A$2:$A${qlast},A{r},{RQ}!$K$2:$K${qlast},"Sí")')
    wr.cell(r,5,f'=COUNTIFS({RQ}!$A$2:$A${qlast},A{r},{RQ}!$L$2:$L${qlast},">0")')
    wr.cell(r,6,f'=IFERROR(AVERAGEIFS({RQ}!$L$2:$L${qlast},{RQ}!$A$2:$A${qlast},A{r}),0)').number_format="0.00"
    wr.cell(r,7,f'=IF(F{r}>0,F{r}/5,"")').number_format="0%"
    wr.cell(r,8,nivel(f"F{r}"))
    for j in range(1,9): wr.cell(r,j).border=B
lr=first+len(DIMS)-1; tr=lr+1
wr.cell(tr,1,"TOTAL").font=Font(bold=True); wr.cell(tr,2,"Puntaje total ponderado (solo dimensiones evaluadas)").font=Font(bold=True)
wr.cell(tr,3,f"=SUM(C{first}:C{lr})").number_format="0%"
wr.cell(tr,4,f"=SUM(D{first}:D{lr})"); wr.cell(tr,5,f"=SUM(E{first}:E{lr})")
wr.cell(tr,6,f'=IFERROR(SUMPRODUCT((F{first}:F{lr}>0)*F{first}:F{lr}*C{first}:C{lr})/SUMPRODUCT((F{first}:F{lr}>0)*C{first}:C{lr}),0)').number_format="0.00"
wr.cell(tr,7,f'=IF(F{tr}>0,F{tr}/5,"")').number_format="0%"
wr.cell(tr,8,nivel(f"F{tr}")).font=Font(bold=True)
for j in range(1,9): wr.cell(tr,j).border=B; wr.cell(tr,j).fill=PatternFill("solid",fgColor=LIGHT)
wr.cell(tr+1,1,f'=IF(ROUND(C{tr},4)=1,"Pesos correctos (100%)","ATENCIÓN: los pesos no suman 100%")').font=Font(italic=True,color="595959")

fb=tr+3
wr.cell(fb,1,"VISTA POR MARCO DE REFERENCIA (promedio de controles puntuados)").font=Font(bold=True,color=NAVY)
for i,h in enumerate(["Marco","","Controles","Evaluados","Puntaje (1-5)"],1):
    c=wr.cell(fb+1,i,h); c.font=H; c.fill=HF; c.alignment=CT
marcos=["ISO 27001:2022","NIST CSF 2.0","CIS Controls v8.1","COBIT 5","PCI DSS v4.0.1","Normativa chilena"]
CM=f"{CMN}"
rr=fb+2
for m in marcos:
    wr.cell(rr,1,m); wr.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=2)
    wr.cell(rr,3,f'=COUNTIF({CM}!$A$2:$A${clast},A{rr})')
    wr.cell(rr,4,f'=COUNTIFS({CM}!$A$2:$A${clast},A{rr},{CM}!$E$2:$E${clast},">0")')
    wr.cell(rr,5,f'=IFERROR(AVERAGEIFS({CM}!$E$2:$E${clast},{CM}!$A$2:$A${clast},A{rr}),"")').number_format="0.00"
    rr+=1
rr+=1
wr.cell(rr,1,"ISO 27001 por tema").font=Font(bold=True,color=NAVY); rr+=1
temas=[]
for _,_,_,t in ctrls:
    if _=="ISO 27001:2022" and t not in temas: temas.append(t)
for t in temas:
    wr.cell(rr,1,t); wr.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=2)
    wr.cell(rr,3,f'=COUNTIFS({CM}!$D$2:$D${clast},A{rr},{CM}!$A$2:$A${clast},"ISO 27001:2022")')
    wr.cell(rr,4,f'=COUNTIFS({CM}!$D$2:$D${clast},A{rr},{CM}!$A$2:$A${clast},"ISO 27001:2022",{CM}!$E$2:$E${clast},">0")')
    wr.cell(rr,5,f'=IFERROR(AVERAGEIFS({CM}!$E$2:$E${clast},{CM}!$D$2:$D${clast},A{rr},{CM}!$A$2:$A${clast},"ISO 27001:2022"),"")').number_format="0.00"
    rr+=1
rr+=1
wr.cell(rr,1,"NIST CSF 2.0 por función").font=Font(bold=True,color=NAVY); rr+=1
for f in nist:
    wr.cell(rr,1,f); wr.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=2)
    wr.cell(rr,3,f'=COUNTIFS({CM}!$D$2:$D${clast},A{rr},{CM}!$A$2:$A${clast},"NIST CSF 2.0")')
    wr.cell(rr,4,f'=COUNTIFS({CM}!$D$2:$D${clast},A{rr},{CM}!$A$2:$A${clast},"NIST CSF 2.0",{CM}!$E$2:$E${clast},">0")')
    wr.cell(rr,5,f'=IFERROR(AVERAGEIFS({CM}!$E$2:$E${clast},{CM}!$D$2:$D${clast},A{rr},{CM}!$A$2:$A${clast},"NIST CSF 2.0"),"")').number_format="0.00"
    rr+=1

# ---------- Escala ----------
we=wb.create_sheet("Escala")
we.column_dimensions["A"].width=14; we.column_dimensions["B"].width=24; we.column_dimensions["C"].width=84
header(we,1,["Puntaje","Nivel","Criterio"])
for row in [(1,"No existe","No hay control ni práctica documentada."),
(2,"Inicial","Práctica informal o parcial, sin documentación ni responsable definido."),
(3,"Definido","Control documentado con responsable, aplicado de forma inconsistente o sin evidencia completa."),
(4,"Implementado","Control aplicado en la operación, con evidencia disponible."),
(5,"Optimizado","Control implementado, evidenciado, medido y revisado periódicamente.")]:
    we.append(list(row))
we.append([]); we.append(["Madurez","Rango","Interpretación"])
for b in [("Crítico","1.0 – 1.9","Exposición alta a sanciones. Plan inmediato."),("Inicial","2.0 – 2.9","Brechas significativas. Priorizar D1, D2, D3 y D9."),
("En desarrollo","3.0 – 3.9","Base instalada con brechas de evidencia y consistencia."),("Avanzado","4.0 – 4.5","Cumplimiento sólido. Foco en medición y auditoría."),
("Optimizado","4.6 – 5.0","Madurez con mejora continua demostrable.")]:
    we.append(list(b))
for rr_ in we.iter_rows(min_row=2):
    for c in rr_: c.alignment=WR; c.border=B
for c in we[7]: c.font=H; c.fill=HF

# ---------- Notas ----------
wn=wb.create_sheet("Notas")
wn.column_dimensions["A"].width=120
unused=[k for k in keys if k not in used]
notes=["NOTAS DE USO Y VERIFICACIÓN",
"1. Regla de puntaje: cada requisito legal toma el promedio de sus controles mapeados (y de sus indicadores RAT, ARCOP o transferencias cuando corresponden). No hay puntaje manual.",
f"2. Controles sin requisito asociado: {len(unused)} controles del catálogo no están mapeados a ningún requisito y no afectan el total. Están marcados 'No' en la columna 'Usado en requisitos' de la hoja Controles marcos. Revisar si corresponde mapearlos.",
"3. ARCOP: el estado de plazo se calcula con la fecha de hoy. Hoy hay solicitudes vencidas (ARC-003 y ARC-004, con plazos de julio de 2026). Ver el indicador de la hoja Resumen.",
"4. Plazo de notificación de vulneraciones: las fuentes consultadas no coinciden (algunas indican 72 horas y otras, notificar sin dilación indebida). Confirmar en la BCN antes de fijar plazos internos.",
"5. Vigencia: la ley entra en vigencia el 1 de diciembre de 2026. Hay un proyecto para postergarla a diciembre de 2027; verificar su estado.",
"6. Las referencias a artículos y los controles de los marcos vienen de los documentos del proyecto y no se verificaron una por una contra el texto oficial.",
"7. PCI DSS y COBIT aplican solo si la organización procesa tarjetas o gestiona esos procesos. Usar N/A con justificación cuando no correspondan.",
"8. Benchmark BM-CONT (contenedores) no está mapeado a ningún requisito; se agregó solo al catálogo.",
"9. PCI DSS: el alcance del txt incluye solo los requisitos 2, 5, 7 y 8. Los requisitos 3, 4, 6, 10, 11 y 12 no están en el catálogo. Agregarlos cambia el alcance cotizado y debe confirmarse antes.",
"10. Esta herramienta es de autoevaluación y no constituye asesoría legal."]
for i,n in enumerate(notes,1):
    c=wn.cell(i,1,n); c.alignment=WR
    if i==1: c.font=Font(bold=True,color=NAVY,size=12)
    else: wn.row_dimensions[i].height=42


# ---------- control de cobertura ----------
assert not [k for k in keys if k not in used], "hay controles sin mapear"

# ---------- Dashboard ----------
from openpyxl.chart import BarChart, RadarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
marc_first=fb+2
wd=wb.create_sheet("Dashboard")
wd.sheet_view.showGridLines=False
for col,w in zip("ABCDEFGHIJKLMNOP",[2,12,62,14,16,2,2,2,2,2,2,2,2,2,2,2]): wd.column_dimensions[col].width=w
wd["B1"]="DASHBOARD — CUMPLIMIENTO LEY 21.719"; wd["B1"].font=Font(bold=True,size=16,color=NAVY)
wd["B2"]='="Actualizado: "&TEXT(TODAY(),"dd/mm/yyyy")&"  ·  Escala 1 (crítico) a 5 (optimizado)"'
wd["B2"].font=Font(italic=True,color="595959")
kpis=[("Puntaje total ponderado (1-5)",f"=Resumen!F{tr}","0.00"),
("Nivel de madurez",f'=IF(Resumen!E{tr}/Resumen!D{tr}<0.5,"Cobertura insuficiente (<50%)",Resumen!H{tr})',None),
("Requisitos evaluados",f'=Resumen!E{tr}&" de "&Resumen!D{tr}',None),
("Controles puntuados",f'=COUNTIF({CMN}!E2:E{clast},">0")&" de {len(ctrls)}"',None),
("Solicitudes ARCOP vencidas","=Resumen!C7","0"),
("Completitud del RAT (1-5)","=Resumen!C4","0.00")]
for i,(lab,f,fmt) in enumerate(kpis):
    r=4+i
    c=wd.cell(r,3,lab); c.font=Font(bold=True,color="FFFFFF"); c.fill=HF; c.alignment=Alignment(vertical="center")
    v=wd.cell(r,4,f); v.font=Font(bold=True,size=12,color=NAVY); v.fill=PatternFill("solid",fgColor=LIGHT); v.alignment=CT
    if fmt: v.number_format=fmt
    wd.row_dimensions[r].height=22
wd.conditional_formatting.add("D8",CellIsRule(operator="greaterThan",formula=["0"],fill=PatternFill("solid",fgColor="F8CBAD")))

wd.cell(10,3,"Requisitos incluidos en el puntaje").font=Font(bold=True,color="FFFFFF"); wd.cell(10,3).fill=HF
wd.cell(10,4,f'=COUNTIF(\'Requisitos Ley\'!K2:K{qlast},"Sí")&" de "&COUNTA(\'Requisitos Ley\'!B2:B{qlast})').font=Font(bold=True,size=12,color=NAVY)
wd.cell(10,4).fill=PatternFill("solid",fgColor=LIGHT); wd.cell(10,4).alignment=CT
# tabla dimensiones
dh=11
for i,h in enumerate(["Código","Dimensión","Puntaje (1-5)","Nivel"],2):
    c=wd.cell(dh,i,h); c.font=H; c.fill=HF; c.alignment=CT
for i in range(len(DIMS)):
    r=dh+1+i; s=first+i
    wd.cell(r,2,f"=Resumen!A{s}"); wd.cell(r,3,f"=Resumen!B{s}")
    wd.cell(r,4,f"=Resumen!F{s}").number_format="0.00"; wd.cell(r,5,f"=Resumen!H{s}")
    for j in range(2,6): wd.cell(r,j).border=B
dlast=dh+len(DIMS)
wd.conditional_formatting.add(f"D{dh+1}:D{dlast}",ColorScaleRule(start_type="num",start_value=1,start_color="F8696B",mid_type="num",mid_value=3,mid_color="FFEB84",end_type="num",end_value=5,end_color="63BE7B"))

# tabla marcos
mh=dlast+2
for i,h in enumerate(["","Marco de referencia","Evaluados","Puntaje (1-5)"],2):
    c=wd.cell(mh,i,h); c.font=H; c.fill=HF; c.alignment=CT
for i in range(len(marcos)):
    r=mh+1+i; s=marc_first+i
    wd.cell(r,3,f"=Resumen!A{s}"); wd.cell(r,4,f"=Resumen!D{s}"); wd.cell(r,5,f"=Resumen!E{s}").number_format="0.00"
    for j in range(2,6): wd.cell(r,j).border=B
mlast=mh+len(marcos)
wd.conditional_formatting.add(f"E{mh+1}:E{mlast}",ColorScaleRule(start_type="num",start_value=1,start_color="F8696B",mid_type="num",mid_value=3,mid_color="FFEB84",end_type="num",end_value=5,end_color="63BE7B"))

# brechas prioritarias
bh=mlast+2
for i,h in enumerate(["Código","Brecha prioritaria (menor puntaje evaluado)","Puntaje (1-5)"],2):
    c=wd.cell(bh,i,h); c.font=H; c.fill=HF; c.alignment=CT
RQR=f"'Requisitos Ley'!$L$2:$L${qlast}"
for k in range(1,6):
    r=bh+k
    wd.cell(r,4,f'=IFERROR(SMALL({RQR},{k}),"")').number_format="0.00"
    wd.cell(r,2,f'=IF(D{r}="","",INDEX(\'Requisitos Ley\'!$B$2:$B${qlast},MATCH(D{r},{RQR},0)))')
    wd.cell(r,3,f'=IF(D{r}="","",INDEX(\'Requisitos Ley\'!$C$2:$C${qlast},MATCH(D{r},{RQR},0)))').alignment=WR
    for j in range(2,6): wd.cell(r,j).border=B
    wd.row_dimensions[r].height=30
wd.conditional_formatting.add(f"D{bh+1}:D{bh+5}",ColorScaleRule(start_type="num",start_value=1,start_color="F8696B",mid_type="num",mid_value=3,mid_color="FFEB84",end_type="num",end_value=5,end_color="63BE7B"))

# gráficos
bar=BarChart(); bar.type="bar"; bar.style=10; bar.title="Puntaje por dimensión (1-5)"
bar.add_data(Reference(wd,min_col=4,min_row=dh,max_row=dlast),titles_from_data=True)
bar.set_categories(Reference(wd,min_col=2,min_row=dh+1,max_row=dlast))
bar.y_axis.scaling.min=0; bar.y_axis.scaling.max=5; bar.legend=None; bar.height=10; bar.width=18
wd.add_chart(bar,"G4")
rad=RadarChart(); rad.title="Perfil de cumplimiento por dimensión"; rad.style=26
rad.add_data(Reference(wd,min_col=4,min_row=dh,max_row=dlast),titles_from_data=True)
rad.set_categories(Reference(wd,min_col=2,min_row=dh+1,max_row=dlast))
rad.y_axis.scaling.min=0; rad.y_axis.scaling.max=5; rad.height=10; rad.width=18
wd.add_chart(rad,"G26")
mb=BarChart(); mb.type="col"; mb.style=10; mb.title="Puntaje por marco de referencia (1-5)"
mb.add_data(Reference(wd,min_col=5,min_row=mh,max_row=mlast),titles_from_data=True)
mb.set_categories(Reference(wd,min_col=3,min_row=mh+1,max_row=mlast))
mb.y_axis.scaling.min=0; mb.y_axis.scaling.max=5; mb.legend=None; mb.height=9; mb.width=18
wd.add_chart(mb,"G48")


# ================= v5: instrucciones, validaciones y protección =================
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Protection
UNL=Protection(locked=False)
def unlock(ws_,rng):
    for row in ws_[rng]:
        for c in row:
            c.protection=UNL; c.fill=PatternFill("solid",fgColor=INPUT)
def dv_list(ws_,rng,src,title,msg,warn=False,prompt=None):
    dv_=DataValidation(type="list",formula1=src,allow_blank=True,showErrorMessage=True,
        errorStyle="warning" if warn else "stop",errorTitle=title,error=msg,
        showInputMessage=bool(prompt),promptTitle=title if prompt else None,prompt=prompt)
    ws_.add_data_validation(dv_); dv_.add(rng)
def dv_date(ws_,rng,warn=False):
    dv_=DataValidation(type="date",operator="between",formula1="43831",formula2="49674",allow_blank=True,
        showErrorMessage=True,errorStyle="warning" if warn else "stop",errorTitle="Fecha no válida",
        error="Ingrese una fecha válida (dd/mm/aaaa) entre 2020 y 2035.")
    ws_.add_data_validation(dv_); dv_.add(rng)
def dv_decimal(ws_,rng):
    dv_=DataValidation(type="decimal",operator="between",formula1="0",formula2="1",allow_blank=True,
        showErrorMessage=True,errorTitle="Peso no válido",error="Ingrese un porcentaje entre 0% y 100%.")
    ws_.add_data_validation(dv_); dv_.add(rng)

# Listas de valores aceptados
wl=wb.create_sheet("Listas")
LIST={"A":("Base legal (Art. 13)",["Lit. a — Consentimiento expreso","Lit. b — Ejecución de contrato","Lit. c — Obligación legal","Lit. d — Interés vital","Lit. e — Interés legítimo","Lit. f — Interés público o autoridad pública"]),
"B":("Tipo de derecho ARCOP",["Acceso","Rectificación","Cancelación / Supresión","Oposición","Portabilidad","Decisiones automatizadas"]),
"C":("Estado ARCOP",["Pendiente","En proceso","Respondido","Rechazado"]),
"D":("Sí / No",["Sí","No"]),
"E":("Puntaje",["1","2","3","4","5","N/A"])}
for col,(titulo,vals) in LIST.items():
    wl[f"{col}1"]=titulo; wl[f"{col}1"].font=Font(bold=True)
    for i,v in enumerate(vals): wl[f"{col}{i+2}"]=v
    wl.column_dimensions[col].width=44
def lref(col): return f"=Listas!${col}$2:${col}${len(LIST[col][1])+1}"

# Portada: campos editables
unlock(wb["Portada"],"C5:C9")

# Resumen: pesos editables; RAT indicador con rango ampliado
unlock(wr,f"C{first}:C{lr}"); dv_decimal(wr,f"C{first}:C{lr}")
wr["C4"]='=IFERROR(1+4*AVERAGE(RAT!$AA$2:$AA$60),"")'

# Requisitos Ley: solo evidencia, plan, responsable y fecha
unlock(wq,f"F2:I{qlast}")
dv_list(wq,f"K2:K{qlast}",lref("D"),"Incluir en puntaje","Seleccione Sí o No.",warn=False,prompt="Sí = entra en el puntaje")
dv_date(wq,f"I2:I{qlast}")

# Controles marcos: puntaje
unlock(wcm,f"E2:E{clast}")

# RAT: filas 2..60, fórmula de completitud en Q
for r in range(2,61):
    wrat.cell(r,27,f'=IF(LEFT(A{r},4)="RAT-",COUNTA(B{r}:Z{r})/25,"")').number_format="0%"
unlock(wrat,"A2:P60")
dv_list(wrat,"D2:D60",lref("A"),"Base legal","Seleccione una base legal de la lista (Art. 13).",warn=True,prompt="Elija de la lista")
dv_list(wrat,"F2:F60",lref("D"),"Datos sensibles","Responda Sí o No. Si hay un detalle, escríbalo en la celda como texto.",warn=True,prompt="Sí o No")
dv_date(wrat,"P2:P60",warn=True)

# ARCOP: filas 2..60, estado de plazo en I
for r in range(2,61):
    warc.cell(r,9,f'=IF(LEFT(A{r},4)="ARC-",IF(ISNUMBER(F{r}),IF(F{r}<=E{r},"A tiempo","Fuera de plazo"),IF(TODAY()>E{r},"Vencida","En plazo")),"")')
unlock(warc,"A2:H60")
dv_list(warc,"C2:C60",lref("B"),"Tipo de derecho","Seleccione el tipo de derecho ARCOP de la lista.",warn=True,prompt="Elija de la lista")
dv_list(warc,"G2:G60",lref("C"),"Estado","Seleccione el estado de la solicitud de la lista.",warn=True,prompt="Elija de la lista")
dv_date(warc,"D2:F60",warn=True)

# Transferencias: filas 2..60, garantía en J
for r in range(2,61):
    wtr.cell(r,10,f'=IF(LEFT(A{r},4)="RAT-",IF(LEN(G{r})>0,"Sí","No"),"")')
unlock(wtr,"A2:I60")

# ---------- Hoja Cómo llenar ----------
wh=wb.create_sheet("Cómo llenar")
wh.column_dimensions["A"].width=6; wh.column_dimensions["B"].width=20
wh.column_dimensions["C"].width=60; wh.column_dimensions["D"].width=46; wh.column_dimensions["E"].width=40
wh["B1"]="CÓMO LLENAR ESTE CHECKLIST"; wh["B1"].font=Font(bold=True,size=15,color=NAVY)
wh["B2"]="Regla básica: las celdas AMARILLAS son las únicas que puede editar. El resto está bloqueado para evitar cambios por error."
wh["B2"].alignment=WR; wh.merge_cells("B2:E2"); wh.row_dimensions[2].height=32
wh["B3"]="Sin contraseña. Si necesita cambiar una fórmula: Revisar > Desproteger hoja. Vuelva a proteger al terminar."
wh["B3"].alignment=WR; wh.merge_cells("B3:E3"); wh.row_dimensions[3].height=28
header(wh,5,["Paso","Hoja","Qué hacer","Valores aceptados","Ejemplo"])
steps=[
("1","Portada","Escriba el nombre de la organización, el responsable, la fecha de evaluación, la próxima revisión y el evaluador.","Texto libre. Fechas en formato dd/mm/aaaa.","Fecha de evaluación: 19/10/2026"),
("2","RAT","Una fila por cada actividad de tratamiento de datos personales. Complete las columnas B a P. La columna Q (completitud) se calcula sola: 100% = las 15 columnas llenas.","Base legal: elija de la lista. Datos sensibles: Sí o No. Última revisión: dd/mm/aaaa. El resto, texto.","RAT-007 · Gestión de postventa · Lit. b — Ejecución de contrato"),
("3","ARCOP","Una fila por cada solicitud de derechos del titular. Tipo y Estado se eligen de la lista. La columna Estado de plazo se calcula sola (A tiempo, Fuera de plazo, En plazo o Vencida).","Tipo: lista. Estado: Pendiente, En proceso, Respondido o Rechazado. Fechas dd/mm/aaaa. Fecha límite: 30 días corridos desde la recepción (verifique el plazo en el texto oficial).","ARC-005 · Acceso · recibida 02/10/2026 · límite 01/11/2026"),
("4","Transferencias","Una fila por cada destinatario fuera de Chile. Use el mismo ID del RAT. La columna Garantía definida se calcula sola: Sí si escribió la garantía.","ID RAT: el mismo de la hoja RAT. Garantías: texto (por ejemplo, SCC + DPA).","RAT-003 · Mailchimp · EE.UU. · SCC + DPA"),
("5","Controles marcos","Para cada control, lea en la columna C qué debe levantar. El puntaje se calcula solo a partir de la hoja Registro entrevistas: registre allí cada entrevista y su evidencia verificada. Use N/A solo si el control no aplica, y justifique en la columna Evidencia (escriba la razón en esa celda).","1, 2, 3, 4, 5 o N/A. Guía de niveles en la hoja Escala.","ISO 8.7 · 4 (EDR en todos los endpoints, sin revisión trimestral)"),
("6","Requisitos Ley","No se puntúa aquí: el puntaje sale del promedio de los controles mapeados. Complete Evidencia, Brecha / plan de acción, Responsable y Fecha límite. Para incluir u omitir requisitos extra, cambie 'Incluir en puntaje' (columna K): Sí o No.","Texto libre. Fecha límite en dd/mm/aaaa.","Brecha: falta matriz de accesos. Responsable: CISO. Fecha: 30/11/2026"),
("7","Resumen","Ajuste el peso de cada dimensión si su organización lo requiere. Los pesos deben sumar 100%.","Porcentaje entre 0% y 100%. La celda de control debe decir 'Pesos correctos'.","D8 Seguridad: 11%"),
("8","Entrevistas","Prepare y registre cada entrevista por dimensión: quién consultar, qué preguntar y qué evidencia pedir. Lo que escriba aquí no cambia el puntaje; úselo para puntuar los controles.","Texto libre. Puntaje sugerido: 1 a 5 o N/A.","Gerencia de RRHH: contratos sin cláusulas de privacidad; expedientes en papel sin control de acceso."),
("9","Dashboard","No se edita. Revise el resultado después de completar los pasos anteriores.","—","Si aparece 'Cobertura insuficiente', faltan requisitos por evaluar."),
]
for i,row in enumerate(steps):
    r=6+i
    for j,v in enumerate(row,1):
        c=wh.cell(r,j,v); c.alignment=WR; c.border=B
    wh.row_dimensions[r].height=78
r=6+len(steps)+1
wh.cell(r,2,"ORDEN RECOMENDADO").font=Font(bold=True,color=NAVY); r+=1
wh.cell(r,2,"1 → 2 → 3 → 4 → 5 → 6 → 7 → revisar el Dashboard. Los pasos 2, 3 y 4 alimentan indicadores que afectan los requisitos de la ley.").alignment=WR
wh.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5); wh.row_dimensions[r].height=30
r+=2
wh.cell(r,2,"ERRORES COMUNES").font=Font(bold=True,color=NAVY); r+=1
for t in ["Una fecha escrita como texto no se reconoce en el cálculo de plazos. Use dd/mm/aaaa.",
"Si una celda muestra un aviso amarillo al escribir, revise el valor. Puede guardarlo, pero el cálculo puede no reflejarlo.",
"No borre fórmulas de las columnas Q (RAT), I (ARCOP) ni J (Transferencias). Están bloqueadas.",
"Un requisito sin controles puntuados aparece en blanco y no pesa en el total. Por eso el Dashboard muestra la cobertura."]:
    wh.cell(r,2,"• "+t).alignment=WR; wh.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5); wh.row_dimensions[r].height=30; r+=1


# ---------- Hoja Entrevistas ----------
we2=wb.create_sheet("Entrevistas")
we2.column_dimensions["A"].width=8; we2.column_dimensions["B"].width=36; we2.column_dimensions["C"].width=28
we2.column_dimensions["D"].width=70; we2.column_dimensions["E"].width=40; we2.column_dimensions["F"].width=14
we2.column_dimensions["G"].width=22; we2.column_dimensions["H"].width=44; we2.column_dimensions["I"].width=14
we2["A1"]="GUÍA DE ENTREVISTAS POR DIMENSIÓN"; we2["A1"].font=Font(bold=True,size=14,color=NAVY)
we2["A2"]="Guía de preguntas por dimensión para preparar cada entrevista. Para registrar lo que responde la persona, use la hoja Registro entrevistas: el puntaje se calcula desde allí."
we2["A2"].alignment=WR; we2.merge_cells("A2:I2"); we2.row_dimensions[2].height=30
header(we2,4,["Código","Dimensión","Entrevistado sugerido","Preguntas clave","Evidencia a solicitar","Fecha entrevista","Entrevistado real","Hallazgos","Puntaje sugerido (1-5)"])
GUIA={
"D1":("Gerencia General / Directorio","1. ¿Quién aprobó la política de privacidad y cuándo se revisó por última vez?\n2. ¿Existe un responsable de privacidad con recursos asignados?\n3. ¿Cómo demuestra el cumplimiento si la autoridad lo solicita?","Política aprobada, acta de directorio, organigrama con roles"),
"D2":("Gerencia de TI y responsable de datos","1. ¿Existe un inventario de sistemas que tratan datos personales?\n2. ¿Quién mantiene el RAT y cuándo se actualizó?","RAT vigente, inventario de sistemas"),
"D3":("Gerencia Comercial y Marketing","1. ¿Con qué base legal se trata cada base de datos?\n2. ¿Cómo se registra el consentimiento y cómo se revoca?","Formularios de consentimiento, registro de consentimientos"),
"D4":("Servicio al cliente","1. ¿Cómo llegan las solicitudes ARCOP y quién las responde?\n2. ¿Cuánto demoran en promedio y cómo verifican la identidad del solicitante?","Registro de solicitudes ARCOP, procedimiento"),
"D5":("Dueños de procesos (Comercial, RRHH, Finanzas)","1. ¿Qué datos piden y para qué los usan?\n2. ¿Cuánto tiempo guardan cada categoría y cómo la eliminan?","Política de retención, evidencia de eliminación"),
"D6":("RRHH y Seguridad","1. ¿Recolectan datos de salud, licencias médicas o biométricos?\n2. ¿Quién accede a ellos?","Inventario de datos sensibles, controles de acceso"),
"D7":("Gerencia de TI y de proyectos","1. ¿Se evalúa la privacidad antes de lanzar un sistema o campaña?\n2. ¿Hay EIPD documentada para tratamientos de alto riesgo (videovigilancia, biometría, perfilamiento)?","EIPD, checklist de privacidad en proyectos"),
"D8":("CISO / Jefe de TI","1. ¿Cómo se gestionan accesos y roles (RBAC)?\n2. ¿En qué sistemas hay MFA y qué antivirus o EDR usan?\n3. ¿Cuándo fue la última prueba de intrusión?","Matriz de accesos, reportes EDR, informe de pentest"),
"D9":("CISO y Gerencia de Operaciones","1. ¿Qué pasa si hay una brecha de datos un viernes en la noche?\n2. ¿Cuándo fue la última prueba de restauración y del plan de recuperación?","IRP, registro de simulacros, pruebas de restauración"),
"D10":("Compras / Administración","1. ¿Qué proveedores tratan datos personales?\n2. ¿Sus contratos tienen cláusulas de privacidad y eliminación al término? ¿Tienen sub-encargados?","Contratos, inventario de proveedores"),
"D11":("RRHH y Capacitación","1. ¿Cuántas personas completaron capacitación en privacidad el último año?\n2. ¿Hacen simulacros de phishing y con qué resultados?","Registro de capacitación, resultados de simulacros"),
"D12":("Auditoría interna / Gerencia General","1. ¿Qué indicadores de privacidad se reportan y con qué frecuencia?\n2. ¿Existe un plan de auditoría anual?","Plan de auditoría, informes, indicadores"),
"D13":("CISO","1. ¿Qué marcos utilizan y cuál es su alcance?\n2. ¿Cómo se delimita el entorno de pagos (tiendas, bodegas, ecommerce) para PCI DSS?","Alcance documentado, informes de cumplimiento"),
"D14":("Jefe de instalaciones y seguridad física","1. ¿Cómo se controla el acceso a bodegas, oficinas y centros de datos?\n2. ¿Hay CCTV y cuánto se retiene la grabación?","Políticas de acceso físico, retención de CCTV"),
"D15":("Jefe de infraestructura y DBA","1. ¿Qué bases de datos contienen datos personales y están cifradas?\n2. ¿Cómo está segmentada la red?\n3. ¿Se prueba la restauración de las bases?","Inventario de BD, diagramas de red, pruebas de restauración"),
"D16":("Gerencia de TI / PMO","1. ¿Cómo se aprueban los cambios en producción?\n2. ¿Existe un registro de configuración (CMDB)?","Procedimiento de cambios, CMDB"),
"D17":("Gerencia de RRHH","1. ¿Qué información piden en entrevistas y cuánto guardan los CV de candidatos no contratados?\n2. ¿El control de asistencia usa huella o reconocimiento facial? ¿Hay videovigilancia en tiendas o bodegas?\n3. ¿Cómo se revocan accesos y se recupera el equipo cuando sale un empleado?","Contratos tipo, política de retención, expediente, registro de accesos y asistencia"),
"D18":("Gerencia de Marketing y Ecommerce","1. ¿De dónde vienen las bases de datos de leads y prospectos? ¿Compran listas?\n2. ¿Usan cookies o seguimiento de terceros? ¿Hacen perfilamiento?\n3. ¿Cómo atienden el opt-out y cuánto conservan leads inactivos?","Formularios de consentimiento, política de cookies, contratos con proveedores de listas"),
}
er=5
for code,name,_ in DIMS:
    who,qs,ev=GUIA[code]
    vals=[code,name,who,qs,ev,None,None,None,None]
    for j,v in enumerate(vals,1):
        c=we2.cell(er,j,v); c.alignment=WR; c.border=B
    we2.row_dimensions[er].height=96
    er+=1
elast=er-1
unlock(we2,f"F5:I{elast}")
dv_list(we2,f"I5:I{elast}",lref("E"),"Puntaje","Use 1, 2, 3, 4, 5 o N/A.",warn=False)
dv_date(we2,f"F5:F{elast}",warn=True)
we2.freeze_panes="C5"


# ================= v10: Disclaimer y aceptación =================
from openpyxl.styles import Protection as _Prot
wdis=wb.create_sheet("Disclaimer")
wdis.sheet_view.showGridLines=False
wdis.column_dimensions["A"].width=3; wdis.column_dimensions["B"].width=34; wdis.column_dimensions["C"].width=90
wdis["B2"]="DISCLAIMER, CONFIDENCIALIDAD Y ACEPTACIÓN DE LA INFORMACIÓN"; wdis["B2"].font=Font(bold=True,size=14,color=NAVY)
wdis["B3"]="Levantamiento preliminar (discovery) — Ley N° 21.719 sobre protección de datos personales"; wdis["B3"].font=Font(italic=True)
secciones=[
("1. Naturaleza del documento","Este archivo es un levantamiento preliminar (discovery) para identificar brechas de cumplimiento y dimensionar la etapa de remediación. No constituye auditoría, certificación, dictamen ni asesoría legal. Para conclusiones jurídicas definitivas, consulte a un abogado especialista."),
("2. Base de la información","Los puntajes se calculan a partir de la información que entrega la organización y de las entrevistas realizadas. Salvo que se indique lo contrario, la información no ha sido verificada de forma independiente. Los puntajes son indicativos y pueden cambiar con evidencia adicional."),
("3. Marco normativo","Las referencias a la Ley 21.719 corresponden a la versión disponible al momento de la evaluación. La ley entra en vigencia el 1 de diciembre de 2026; existe un proyecto para postergarla. Los artículos y plazos citados deben validarse contra el texto oficial publicado en la Biblioteca del Congreso Nacional (BCN). Las sanciones no forman parte de este análisis."),
("4. Datos personales en este archivo","Este archivo debe contener la menor cantidad posible de datos personales. En los registros de actividades, solicitudes ARCOP y transferencias use identificadores internos o iniciales, nunca nombres completos, RUT, correos ni datos de salud de personas. Si la información sensible es necesaria para el diagnóstico, regístrela de forma agregada."),
("5. Clasificación y acceso","Este archivo y su contenido se consideran CONFIDENCIALES. Compártalo solo con personas autorizadas por la dirección de la organización. No lo reenvíe por canales no corporativos ni lo publique."),
("6. Entrevistas y aviso a entrevistados","Las entrevistas tienen por finalidad diagnosticar el estado de cumplimiento de la organización. Antes de iniciarlas, informe al entrevistado que sus respuestas se registran en este archivo, quién las usará y por cuánto tiempo se conservarán. Si se graban o se toman fotos o documentos personales, obtenga su consentimiento previo."),
("7. Responsabilidad","La organización es responsable de la veracidad y completitud de la información entregada. El evaluador no responde por decisiones tomadas con base en este archivo sin validación adicional."),
("8. Uso del resultado","El resultado sirve para priorizar la remediación. Antes de implementar cambios, valide cada brecha con su área responsable y con asesoría legal cuando corresponda."),
]
r=5
for titulo,texto in secciones:
    wdis.cell(r,2,titulo).font=Font(bold=True,color=NAVY)
    c=wdis.cell(r,3,texto); c.alignment=WR
    wdis.row_dimensions[r].height=max(48, 16*(len(texto)//95+1))
    wdis.cell(r,2).alignment=WR
    r+=1
r+=1
wdis.cell(r,2,"ACEPTACIÓN").font=Font(bold=True,size=12,color=NAVY); r+=1
wdis.cell(r,3,"Complete los campos amarillos. Al responder 'Sí', usted declara haber leído y aceptado las condiciones anteriores.").alignment=WR
wdis.row_dimensions[r].height=30; r+=1
campos=[("Organización evaluada",""),("Nombre de quien acepta",""),("Cargo",""),("Fecha de aceptación (dd/mm/aaaa)",""),("¿Acepta las condiciones? (Sí/No)",""),("Observaciones",""),("Firma (nombre escrito)","")]
accept_rows=[]
for lab,val in campos:
    wdis.cell(r,2,lab).font=Font(bold=True)
    c=wdis.cell(r,3,val); c.fill=INF; c.protection=_Prot(locked=False); c.alignment=WR; c.border=B
    wdis.row_dimensions[r].height=24 if lab!="Observaciones" else 48
    accept_rows.append(r); r+=1
dv_acc=DataValidation(type="list",formula1=lref("D"),allow_blank=True,showErrorMessage=True,errorTitle="Respuesta no válida",error="Seleccione Sí o No.")
wdis.add_data_validation(dv_acc); dv_acc.add(f"C{accept_rows[4]}")
dv_date_acc=DataValidation(type="date",operator="between",formula1="43831",formula2="49674",allow_blank=True,showErrorMessage=True,errorTitle="Fecha no válida",error="Ingrese una fecha válida dd/mm/aaaa.")
wdis.add_data_validation(dv_date_acc); dv_date_acc.add(f"C{accept_rows[3]}")


# ================= v13: Registro de entrevistas y consolidación automática =================
from openpyxl.styles import Protection as _P2
RG="'Registro entrevistas'"
R0,R1=30,229
wrg=wb.create_sheet("Registro entrevistas")
wrg.sheet_view.showGridLines=False
for col,w in zip("ABCDEFGHIJKL",[11,12,12,22,24,16,40,36,14,14,36,22]): wrg.column_dimensions[col].width=w
wrg["A1"]="REGISTRO DE ENTREVISTAS Y CONSOLIDACIÓN DEL PUNTAJE POR CONTROL"; wrg["A1"].font=Font(bold=True,size=14,color=NAVY)
expl=[
("CÓMO SE USA",""),
("Paso 1","Registre una fila por entrevista y por control evaluado. No registre una fila por dimensión: un control puede tener varias entrevistas (una por área)."),
("Paso 2","Escriba en 'Lo que afirma el entrevistado' la respuesta resumida, y en 'Evidencia solicitada' el documento o muestra que pidió."),
("Paso 3","Marque 'Evidencia verificada' en Sí solo si usted revisó el documento o la muestra. Si solo hay una afirmación verbal, marque No."),
("Paso 4","Asigne el 'Puntaje propuesto' (1 a 5 o N/A) con la escala de la hoja Escala."),
("Paso 5","El puntaje de cada control aparece solo en la hoja Controles marcos (columna E). No se escribe a mano."),
("POR QUÉ SE USA EL MÍNIMO VERIFICADO",""),
("Regla","El puntaje de un control es el MÍNIMO de los puntajes propuestos que tienen evidencia verificada."),
("Razón 1","Un control vale lo que vale su punto más débil. Si un área dice 4 y otra muestra que el control no opera (1), el promedio daría 2,5 y ocultaría la brecha. El mínimo la muestra."),
("Razón 2","El promedio premia a quien responde mejor. Con el mínimo, la discrepancia queda visible y se investiga."),
("Razón 3","Es el criterio con que un regulador o auditor evaluaría: la organización debe demostrar el control, no solo declararlo."),
("Razón 4","Evita inflar la madurez, que es el sesgo natural de un assessment."),
("REGLAS ADICIONALES",""),
("Sin evidencia","Si un control no tiene ninguna evidencia verificada, su puntaje máximo es 3 (Definido), porque no se puede demostrar que opera."),
("N/A y vacíos","Un N/A no cuenta como puntaje. Un control sin entrevistas queda en blanco y no pesa en el total."),
("Discrepancia","Si dos puntajes del mismo control difieren en 2 puntos o más, la columna Alerta marca 'Revisar'. Valide con el dueño del proceso antes de cerrar el control."),
("Límite del método","El mínimo puede ser severo si una entrevista está mal informada. Por eso la alerta exige validación. Una brecha identificada a tiempo es mejor que una madurez inflada."),
]
r=3
for k,v in expl:
    if v=="":
        wrg.cell(r,1,k).font=Font(bold=True,color=NAVY); r+=1; continue
    wrg.cell(r,1,k).font=Font(bold=True)
    c=wrg.cell(r,2,v); c.alignment=WR; wrg.merge_cells(start_row=r,start_column=2,end_row=r,end_column=12)
    wrg.row_dimensions[r].height=30; r+=1
assert r<=R0-2, r
hdr_r=R0-1
heads=["ID entrevista","Fecha","Dimensión (D1–D18)","Área","Cargo del entrevistado","Control (ID)","Lo que afirma el entrevistado","Evidencia solicitada","Evidencia verificada (Sí/No)","Puntaje propuesto (1-5 / N/A)","Observaciones / discrepancia","Alerta (automática)"]
for j,h in enumerate(heads,1):
    c=wrg.cell(hdr_r,j,h); c.font=H; c.fill=HF; c.alignment=CT; c.border=B
wrg.row_dimensions[hdr_r].height=42
for rr_ in range(R0,R1+1):
    for j in range(1,12):
        c=wrg.cell(rr_,j); c.protection=_P2(locked=False); c.fill=INF; c.border=B; c.alignment=WR
    wrg.cell(rr_,12,f'=IF(F{rr_}="","",IF(COUNTIF($F${R0}:$F${R1},F{rr_})<2,"",IF(_xlfn.MAXIFS($J${R0}:$J${R1},$F${R0}:$F${R1},F{rr_})-_xlfn.MINIFS($J${R0}:$J${R1},$F${R0}:$F${R1},F{rr_})>=2,"Revisar: discrepancia","")))').border=B
    wrg.cell(rr_,2).number_format="dd/mm/yyyy"
dv_d=DataValidation(type="list",formula1='"D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15,D16,D17,D18"',allow_blank=True,showErrorMessage=True,errorTitle="Dimensión",error="Seleccione D1 a D18.")
wrg.add_data_validation(dv_d); dv_d.add(f"C{R0}:C{R1}")
dv_c=DataValidation(type="list",formula1=f"'Controles marcos'!$B$2:$B${clast}",allow_blank=True,showErrorMessage=True,errorTitle="Control no válido",error="Seleccione un control de la hoja Controles marcos (columna B).")
wrg.add_data_validation(dv_c); dv_c.add(f"F{R0}:F{R1}")
dv_v=DataValidation(type="list",formula1=lref("D"),allow_blank=True,showErrorMessage=True,errorTitle="Respuesta no válida",error="Seleccione Sí o No.")
wrg.add_data_validation(dv_v); dv_v.add(f"I{R0}:I{R1}")
dv_p=DataValidation(type="list",formula1=lref("E"),allow_blank=True,showErrorMessage=True,errorTitle="Puntaje no válido",error="Use 1, 2, 3, 4, 5 o N/A.")
wrg.add_data_validation(dv_p); dv_p.add(f"J{R0}:J{R1}")
dv_dt=DataValidation(type="date",operator="between",formula1="43831",formula2="49674",allow_blank=True,showErrorMessage=True,errorTitle="Fecha no válida",error="Ingrese una fecha dd/mm/aaaa.")
wrg.add_data_validation(dv_dt); dv_dt.add(f"B{R0}:B{R1}")
wrg.freeze_panes=f"A{R0}"

wcm["E1"]="Puntaje consolidado (automático desde Registro entrevistas)"
wcm["F1"]="Referencia de evidencia (opcional)"
for rr_ in range(2,clast+1):
    cell=wcm.cell(rr_,5)
    cell.value=(f'=IF(COUNTIF({RG}!$F${R0}:$F${R1},B{rr_})=0,"",'
      f'IF(COUNTIFS({RG}!$F${R0}:$F${R1},B{rr_},{RG}!$I${R0}:$I${R1},"Sí")=0,'
      f'IF(_xlfn.MINIFS({RG}!$J${R0}:$J${R1},{RG}!$F${R0}:$F${R1},B{rr_})=0,"",MIN(3,_xlfn.MINIFS({RG}!$J${R0}:$J${R1},{RG}!$F${R0}:$F${R1},B{rr_}))),'
      f'IF(_xlfn.MINIFS({RG}!$J${R0}:$J${R1},{RG}!$F${R0}:$F${R1},B{rr_},{RG}!$I${R0}:$I${R1},"Sí")=0,"",_xlfn.MINIFS({RG}!$J${R0}:$J${R1},{RG}!$F${R0}:$F${R1},B{rr_},{RG}!$I${R0}:$I${R1},"Sí"))))')
    cell.fill=PatternFill(fill_type=None); cell.protection=_P2(locked=True); cell.alignment=CT
    cell.number_format="0"

# ---------- Protección de hojas ----------
for name in ["Portada","Disclaimer","Cómo llenar","Entrevistas","Registro entrevistas","Dashboard","Resumen","Requisitos Ley","Controles marcos","RAT","ARCOP","Transferencias","Escala","Notas","Listas"]:
    w_=wb[name]
    w_.protection.sheet=True
    w_.protection.autoFilter=False
    w_.protection.sort=False
    w_.protection.formatColumns=False
    w_.protection.formatRows=False

wn.cell(wn.max_row+1,1,"11. Meta objetivo: la escala 1 a 5 sigue la lógica de los modelos de capacidad y madurez (COBIT 5). El nivel 4 (Implementado) indica control operando con evidencia; el nivel 5 (Optimizado) se reserva para procesos críticos. La meta base es 4,0 y se ajusta según criterios de riesgo (impacto y exposición, enfoque ISO/IEC 27005): datos sensibles o de menores a gran escala, decisiones automatizadas, transferencias o muchos proveedores externos, y videovigilancia o biometría. Cada criterio que aplica suma 0,25, con tope de 4,5. Es un criterio de gestión propuesto, no un requisito numérico de la ley. Cualquier ajuste debe justificarse por escrito.").alignment=WR
wn.row_dimensions[wn.max_row].height=90
# orden de hojas
order=["Portada","Disclaimer","Cómo llenar","Entrevistas","Registro entrevistas","Dashboard","Resumen","Requisitos Ley","Controles marcos","RAT","ARCOP","Transferencias","Escala","Notas","Listas"]
wb._sheets=[wb[n] for n in order]
out="/mnt/user-data/outputs/checklist_ley21719_integrado.xlsx"
wb.save(out)
print("guardado", out, "requisitos", qlast-1, "controles", len(ctrls), "no usados", len(unused))
print("ARCOP estados:", [(warc.cell(r,1).value, warc.cell(r,5).value, warc.cell(r,6).value, warc.cell(r,7).value) for r in arc_rows])

# verificación de fórmulas si hay LibreOffice
if shutil.which("soffice") or shutil.which("libreoffice"):
    exe=shutil.which("soffice") or shutil.which("libreoffice")
    os.makedirs("/tmp/recalc",exist_ok=True)
    subprocess.run([exe,"--headless","--convert-to","xlsx","--outdir","/tmp/recalc",out],capture_output=True,timeout=120)
    rv=openpyxl.load_workbook("/tmp/recalc/checklist_ley21719_integrado.xlsx",data_only=True)
    errs=[(ws_.title,c.coordinate,c.value) for ws_ in rv.worksheets for row in ws_.iter_rows() for c in row if isinstance(c.value,str) and c.value.startswith("#")]
    print("errores de fórmula:", errs[:10], len(errs))
    print("indicadores:", [rv["Resumen"].cell(r,3).value for r in range(4,8)])
else:
    print("sin LibreOffice: verificación de valores no ejecutada")
