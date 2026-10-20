# Especificación de backend y base de datos — Assessment Ley 21.719

Versión: 0.1 (borrador para validar antes de construir)
Base: prototipo `checklist_ley21719_frontend.html` (estructura de datos y reglas de cálculo actuales).

---

## 1. Principios

1. **Un solo lugar para las reglas de cálculo.** Hoy el puntaje se calcula en el navegador. En el backend se implementa una única vez (módulo `scoring`) y el front solo muestra resultados. Así no hay discrepancias entre pantalla, Excel y PDF.
2. **Trazabilidad.** Cada puntaje debe poder explicarse: qué entrevista, qué evidencia y qué regla lo produjo.
3. **Datos mínimos.** La información no debe contener nombres completos, RUT ni datos de salud de personas. El sistema debe validar esto cuando sea posible.
4. **Multi-organización desde el inicio.** Todas las tablas de trabajo llevan `organizacion_id`, aunque al principio haya una sola.
5. **Auditoría.** Toda creación, modificación y borrado registra usuario, fecha y valor anterior.

---

## 2. Entidades

### 2.1 Catálogos (cambian poco, se cargan por script o administración)

| Tabla | Campos clave | Notas |
|---|---|---|
| `marco` | `codigo` (PK), `nombre` | ISO 27001:2022, NIST CSF 2.0, CIS v8.1, COBIT 5, PCI DSS v4.0.1, Normativa chilena |
| `control` | `id` (PK, ej. `5.34`, `NIST GV.OC-01`), `marco_codigo`, `nombre`, `descripcion_levantar`, `tema` | 171 registros en el prototipo. `id` es único globalmente |
| `dimension` | `codigo` (PK, D1–D18), `nombre`, `peso` (numeric 0–1) | Suma de pesos = 1 (validación obligatoria) |
| `requisito` | `id` (PK, ej. `1.1`), `dimension_codigo`, `texto`, `tipo` (Obligatorio Ley / Extra) | 191 registros |
| `requisito_control` | `requisito_id`, `control_id`, PK compuesta | Relación N:M. Reemplaza las listas del prototipo |
| `requisito_indicador` | `requisito_id`, `indicador` (`RAT`, `ARCOP`, `TRANSF`) | Indicadores automáticos |
| `actividad_catalogo` | `id`, `tema`, `nombre`, `finalidad`, `base_legal`, `categorias`, `sensibles`, `titulares`, `destinatarios`, `transferencia`, `pais`, `plazo`, `medidas` | Catálogo de 67 actividades. Versionado (`version`, `vigente_desde`) |
| `guia_dimension` | `dimension_codigo`, `entrevistado_sugerido`, `evidencia_solicitar` | Una fila por dimensión |
| `guia_pregunta` | `id`, `dimension_codigo`, `orden`, `texto` | Reemplaza el texto con saltos de línea |

### 2.2 Datos de la organización

| Tabla | Campos clave | Notas |
|---|---|---|
| `organizacion` | `id` (UUID), `nombre`, `area_principal`, `meta_objetivo` (numeric, default 4.0), `fecha_evaluacion`, `responsable`, `evaluador`, `disclaimer_aceptado_por`, `disclaimer_aceptado_en` | Cabecera de la evaluación |
| `usuario` | `id`, `organizacion_id`, `correo` (único), `nombre`, `estado` | Autenticación con proveedor (ver sección 5) |
| `usuario_rol` | `usuario_id`, `organizacion_id`, `rol` | Roles en sección 5 |

### 2.3 Levantamiento

| Tabla | Campos clave | Notas |
|---|---|---|
| `rat_registro` | `id` (`RAT-001`, generado por secuencia por organización), `organizacion_id`, `actividad_id` (FK a catálogo, nullable), `actividad_texto`, `finalidad`, `base_legal`, `categorias`, `sensibles` (Sí/No), `titulares`, `destinatarios`, `transferencia`, `pais`, `plazo`, `medidas`, `sistema`, `responsable`, `actualizado_en`, `actualizado_por` | Completitud se calcula en consulta (15 campos) |
| `arcop_solicitud` | `id` (`ARC-001`), `organizacion_id`, `titular_iniciales`, `tipo_derecho` (enum), `fecha_recepcion`, `fecha_limite` (calculada: recepción + 30 días corridos, configurable), `fecha_respuesta`, `estado` (enum), `observaciones` | Estado de plazo se calcula con la fecha del servidor |
| `transferencia` | `id`, `organizacion_id`, `rat_id` (FK), `destinatario`, `pais`, `categoria_datos`, `finalidad`, `base_legal`, `garantias`, `estado_adecuacion`, `observaciones` | Garantía definida = `garantias` no vacío |

### 2.4 Evaluación

| Tabla | Campos clave | Notas |
|---|---|---|
| `entrevista` | `id`, `organizacion_id`, `fecha`, `area`, `cargo`, `entrevistado` (rol o iniciales, sin nombre completo), `creado_por`, `creado_en`, `actualizado_en` | Una persona. Puede cubrir varias dimensiones |
| `entrevista_dimension` | `entrevista_id`, `dimension_codigo`, PK compuesta | Relación N:M |
| `entrevista_respuesta` | `id`, `entrevista_id`, `pregunta_id` (FK `guia_pregunta`), `respuesta`, `evidencia_presentada` (Sí/No/Parcial), `observaciones` | Una fila por pregunta respondida |
| `entrevista_control` | `id`, `entrevista_id`, `control_id`, `afirma`, `evidencia_verificada` (bool), `puntaje_propuesto` (1–5 o NULL, N/A → NULL con `aplica=false`), `observaciones` | Única por (`entrevista_id`, `control_id`) |
| `evidencia` | `id`, `organizacion_id`, `entrevista_control_id` o `rat_id`, `nombre_archivo`, `ruta_almacenamiento`, `hash_sha256`, `subido_por`, `subido_en`, `tipo` | Archivos en almacenamiento de objetos. Ver sección 6 |

### 2.5 Resultados (calculados, se guardan como caché y se recalculan al cambiar datos)

| Tabla | Campos | Notas |
|---|---|---|
| `puntaje_control` | `organizacion_id`, `control_id`, `valor` (numeric NULL), `n_entrevistas`, `n_verificadas`, `alerta_discrepancia`, `calculado_en` | Valor = regla de la sección 3 |
| `puntaje_requisito` | `organizacion_id`, `requisito_id`, `valor`, `incluido` (bool), `calculado_en` | Incluido = configuración por organización |
| `puntaje_dimension` | `organizacion_id`, `dimension_codigo`, `valor`, `evaluados`, `incluidos`, `calculado_en` | |
| `puntaje_total` | `organizacion_id`, `valor`, `nivel`, `cobertura`, `calculado_en` | |
| `requisito_incluido` | `organizacion_id`, `requisito_id`, `incluido` (bool), `actualizado_por`, `actualizado_en` | Decisión del usuario (Sí/No) |

### 2.6 Auditoría

| Tabla | Campos |
|---|---|
| `auditoria` | `id`, `organizacion_id`, `usuario_id`, `entidad`, `entidad_id`, `accion` (create/update/delete), `antes` (JSON), `despues` (JSON), `en` (timestamp con zona) |

---

## 3. Reglas de cálculo (implementar en el módulo `scoring`)

Son las mismas del prototipo. Cualquier cambio debe pasar por pruebas unitarias con casos de referencia.

1. **Puntaje de control**
   - Se toman los `puntaje_propuesto` no nulos de sus entrevistas.
   - Si hay al menos una entrevista con `evidencia_verificada = true`: valor = mínimo de esas.
   - Si no: valor = `min(3, mínimo de todas)`.
   - Sin puntajes: valor = NULL.
   - Alerta de discrepancia: diferencia máxima ≥ 2 entre puntajes propuestos (con 2 o más).
2. **Indicadores**
   - `RAT` = 1 + 4 × promedio de completitud (campos llenos / 15).
   - `ARCOP` = 1 + 4 × (solicitudes "a tiempo" o "en plazo") / total.
   - `TRANSF` = 1 + 4 × (transferencias con garantía) / total.
3. **Puntaje de requisito**: promedio de sus controles e indicadores con valor no nulo. NULL si no hay ninguno.
4. **Puntaje de dimensión**: promedio de sus requisitos incluidos con valor no nulo. NULL si no hay ninguno.
5. **Puntaje total**: Σ(valor × peso) / Σ(peso) sobre dimensiones con valor.
6. **Nivel**: Crítico < 2; Inicial < 3; En desarrollo < 4; Avanzado < 4,6; Optimizado ≥ 4,6. Cobertura < 50% se muestra como "Cobertura insuficiente".
7. **Brecha** = meta objetivo − valor (mínimo 0). **Prioridad**: Alta ≥ 2; Media ≥ 1; Baja > 0; Cumple meta si ≤ 0.
8. **Aporte a la brecha** = brecha × peso de la dimensión.

Casos de prueba: usar los valores ya verificados en el prototipo (total 4,30 con 140 requisitos incluidos y las 3 transferencias y 6 RAT de referencia; control 5.34 con entrevistas 4 y 1 verificadas = 1).

---

## 4. API (REST, JSON)

Prefijo: `/api/v1`. Todas las rutas requieren autenticación y filtran por `organizacion_id` del usuario.

| Recurso | Método | Ruta | Descripción |
|---|---|---|---|
| Organización | GET / PUT | `/organizacion` | Datos de cabecera, meta objetivo, área principal |
| Disclaimer | POST | `/organizacion/aceptacion` | Registra aceptación (quién, cuándo) |
| Catálogos | GET | `/catalogos/marcos`, `/controles`, `/dimensiones`, `/requisitos`, `/actividades`, `/guia` | Solo lectura para usuarios |
| Requisitos incluidos | PUT | `/requisitos/{id}/incluido` | Cambia Sí/No del requisito |
| RAT | GET / POST | `/rat` | Lista y alta (ID automático) |
| RAT | GET / PATCH / DELETE | `/rat/{id}` | Detalle, edición parcial, baja lógica |
| ARCOP | GET / POST | `/arcop` | |
| ARCOP | GET / PATCH / DELETE | `/arcop/{id}` | |
| Transferencias | GET / POST | `/transferencias` | |
| Transferencias | GET / PATCH / DELETE | `/transferencias/{id}` | |
| Entrevistas | GET / POST | `/entrevistas` | Lista y alta (con dimensiones, respuestas y controles en un solo payload) |
| Entrevistas | GET / PUT / DELETE | `/entrevistas/{id}` | Reemplaza la entrevista completa |
| Evidencias | POST | `/evidencias` | Subida de archivo (multipart). Devuelve `id` y hash |
| Resultados | GET | `/resultados` | Total, dimensiones, requisitos, controles, indicadores (calculados) |
| Diagnóstico | GET | `/diagnostico?meta=4.0` | Brechas por dimensión y requisito, prioridades |
| Exportación | GET | `/exportaciones/excel?alcance=...`, `/exportaciones/pdf` | Generación en servidor (ver sección 7) |
| Respaldo | GET / POST | `/respaldo/export`, `/respaldo/import` | JSON con `schemaVersion` (para migrar datos del prototipo) |

Formato de error: `{ "error": { "codigo": "VALIDACION", "mensaje": "...", "campos": { "fecha": "obligatorio" } } }`.
Concurrencia: cada recurso devuelve `version`; las escrituras exigen `If-Match` para evitar sobrescribir cambios de otra persona.

---

## 5. Usuarios, roles y permisos

| Rol | Puede |
|---|---|
| Administrador de evaluación | Todo, incluida configuración, meta y aceptación del disclaimer |
| Evaluador | Levantamiento, entrevistas, requisitos incluidos, exportaciones |
| Entrevistador (por área) | Registrar entrevistas de sus dimensiones; no ve otras áreas si se activa restricción |
| Lectura (dirección, auditoría externa) | Ver resultados, diagnóstico y exportaciones; sin edición |

Autenticación: proveedor de identidad corporativo (OIDC) o usuario y contraseña con MFA. Sesiones con expiración corta. Registro de accesos en `auditoria`.

---

## 6. Seguridad y tratamiento de datos (Ley 21.719)

- **Cifrado** en tránsito (TLS 1.2+) y en reposo (base de datos y almacenamiento de evidencias).
- **Minimización**: validar en el servidor que `entrevistado` y `titular_iniciales` no contengan patrones de RUT o nombres completos (regla de longitud y formato; revisión manual del resto).
- **Retención**: definir plazo de conservación del propio assessment y de las evidencias; job de eliminación con registro.
- **Evidencias**: almacenar en un bucket privado; nunca en la base de datos. Calcular hash SHA-256 al subir. Escaneo antimalware antes de habilitar la descarga.
- **Separación por organización**: consultas siempre filtradas por `organizacion_id`; pruebas automáticas que lo verifiquen.
- **Respaldos** cifrados, con prueba de restauración mensual.
- **Registros de auditoría** inmutables (solo inserción).

---

## 7. Exportaciones en servidor

- **Excel:** generar con una librería de hojas reales (no XML 2003) para que el archivo sea `.xlsx`. Mantener las hojas actuales: Resumen, Dimensiones, Marcos, Requisitos, Controles, Entrevistas, Respuestas, Controles entrevistados, RAT, ARCOP, Transferencias.
- **PDF:** generar con un motor de plantillas (HTML → PDF) en el servidor, con encabezado y pie de página, numeración y firma de cierre. Incluir la misma sección de diagnóstico.
- Cada exportación se registra en `auditoria` (quién, cuándo, alcance).

---

## 8. Migración desde el prototipo

1. El prototipo exporta JSON (botón "Exportar datos") con `schemaVersion` (agregar en la próxima versión del front).
2. Un script valida el JSON contra el esquema, carga catálogos y datos, y genera un informe de errores antes de confirmar.
3. Las filas del RAT, ARCOP y transferencias se reasignan a IDs con secuencia por organización.
4. Las entrevistas sin dimensiones explícitas (registros antiguos) se marcan como "migradas" y se revisan manualmente.
5. Después de migrar, se recalculan todos los puntajes y se comparan con los valores que mostraba el prototipo (tolerancia 0,001).

---

## 9. Decisiones pendientes (necesito tu respuesta para avanzar)

1. **Stack**: ¿prefieres Node.js/TypeScript, Python (FastAPI) u otro? ¿Algún lineamiento de la organización?
2. **Base de datos**: PostgreSQL recomendado (relacional, buena para auditoría y reportes). ¿Lo permite su infraestructura?
3. **Hosting**: nube (AWS, Azure, GCP) o servidores propios. Esto define el almacenamiento de evidencias y el cifrado.
4. **Autenticación**: ¿existe un proveedor corporativo (Microsoft Entra ID, Google Workspace u otro)?
5. **Multi-organización**: ¿el producto atenderá varios clientes (ShellTI como consultora) o una sola organización por instalación?
6. **Retención**: plazo de conservación del assessment y de las evidencias según la política del cliente.
7. **Restricción por área**: ¿los entrevistadores deben ver solo sus dimensiones?

---

## 10. Próximos pasos propuestos

1. Validar esta especificación y responder las decisiones de la sección 9.
2. Agregar `schemaVersion` y validación de importación al front (cambio pequeño).
3. Crear el esquema de base de datos (migraciones) y el módulo `scoring` con sus pruebas.
4. Construir la API por recursos y conectar el front, reemplazando el almacenamiento local.
5. Migrar datos del prototipo y comparar resultados.

---

## 11. Acceso multiempresa, solicitudes y administración (ampliación v0.2)

Esta sección recoge lo implementado en el prototipo: registro de solicitudes, autorización o revocación de acceso por empresa, y correos de notificación. En el front es una simulación; en el backend debe ser real y seguro.

### 11.1 Entidades nuevas

| Tabla | Campos clave | Notas |
|---|---|---|
| `empresa` | `id` (UUID), `razon_social`, `rut` (único, validado), `estado` (`pendiente`, `autorizada`, `revocada`, `rechazada`), `creada_en`, `revocada_en`, `revocada_por` | Es el **límite de aislamiento**: todos los datos de evaluación pertenecen a una empresa |
| `solicitud_acceso` | `id`, `empresa_id`, `solicitante_nombre`, `cargo`, `correo` (validado), `telefono`, `motivo`, `estado` (`pendiente`, `aprobada`, `rechazada`), `resuelta_por`, `resuelta_en`, `creada_en` | Una empresa puede tener varias solicitudes. Solo una `pendiente` por empresa y correo |
| `usuario` | `id`, `correo` (único), `nombre`, `empresa_id` (nullable para administradores de plataforma), `estado`, `mfa_habilitado`, `ultimo_acceso_en` | Un usuario pertenece a una empresa. Cambiar de empresa requiere otra cuenta |
| `rol_plataforma` | `usuario_id`, `rol` (`admin_plataforma`) | Administradores de ShellTI o del operador; no se mezclan con roles de empresa |
| `rol_empresa` | `usuario_id`, `empresa_id`, `rol` (`administrador`, `evaluador`, `entrevistador`, `lectura`) | Roles por empresa (sección 5) |
| `notificacion` | `id`, `destinatario`, `asunto`, `cuerpo`, `plantilla`, `estado` (`pendiente`, `enviada`, `fallida`), `intentos`, `creada_en`, `enviada_en`, `referencia_tipo`, `referencia_id` | Bandeja de salida (outbox). Un proceso envía y reintenta |
| `sesion` | `id`, `usuario_id`, `empresa_activa_id`, `creada_en`, `expira_en`, `revocada` | Sesión por usuario y empresa activa |

### 11.2 Flujo de acceso (estados)

```
Solicitud  pendiente ──autorizar──► aprobada   · empresa: autorizada
           pendiente ──rechazar───► rechazada  · empresa: rechazada (si no hay otras solicitudes pendientes)
Empresa    autorizada ──revocar──► revocada    · sesiones de sus usuarios se invalidan
           revocada   ──restablecer──► autorizada
```

Reglas:
1. Una empresa `revocada` o `rechazada` no puede iniciar sesión ni consultar datos. Sus datos **no se borran**: quedan retenidos según la política de retención (sección 6).
2. La revocación cierra las sesiones activas de la empresa en el mismo momento (lista de revocación o expiración forzada).
3. Toda decisión (autorizar, rechazar, revocar, restablecer) queda en `auditoria` con quién, cuándo y motivo.
4. Un administrador de plataforma no puede autorizarse a sí mismo para una empresa sin que quede registrado y, idealmente, lo apruebe un segundo administrador (cuatro ojos).

### 11.3 Aislamiento entre empresas (requisito central)

- **Cada consulta** a datos de evaluación debe incluir `empresa_id` derivado de la sesión, nunca del cliente.
- **Seguridad a nivel de fila** en PostgreSQL (`ROW LEVEL SECURITY`) con política por `empresa_id`, como segunda barrera.
- **Pruebas automáticas** que intenten leer datos de otra empresa con un usuario de prueba y deben fallar.
- **Almacenamiento de evidencias** con rutas que incluyan `empresa_id` y permisos por empresa.
- **Exportaciones** (Excel, PDF) generadas solo con datos de la empresa de la sesión.
- **Caché y reportes** segmentados por empresa para evitar fugas.

### 11.4 Notificaciones por correo

| Evento | Destinatario | Plantilla |
|---|---|---|
| Nueva solicitud | Administradores de plataforma (correo configurado) | `solicitud_nueva` |
| Confirmación de solicitud | Solicitante | `solicitud_recibida` |
| Acceso autorizado | Solicitante | `acceso_autorizado` |
| Solicitud no aprobada | Solicitante | `solicitud_rechazada` |
| Acceso revocado | Usuarios de la empresa | `acceso_revocado` |
| Acceso restablecido | Usuarios de la empresa | `acceso_restablecido` |

Requisitos: envío por proveedor transaccional (SMTP autenticado o API), SPF, DKIM y DMARC configurados en el dominio, plantillas sin datos sensibles, enlaces firmados y de vida corta, y registro de cada envío en `notificacion`.

### 11.5 Endpoints adicionales

| Recurso | Método | Ruta | Quién |
|---|---|---|---|
| Solicitud de acceso | POST | `/api/v1/solicitudes` | Público (con limitación de intentos y CAPTCHA) |
| Estado de solicitud | GET | `/api/v1/solicitudes/estado?correo=` | Público, respuesta genérica (no revela si el correo existe) |
| Solicitudes | GET | `/api/v1/admin/solicitudes?estado=` | Administrador de plataforma |
| Decidir solicitud | POST | `/api/v1/admin/solicitudes/{id}/decision` | Administrador (`aprobar` o `rechazar`, con motivo) |
| Empresas | GET | `/api/v1/admin/empresas` | Administrador |
| Cambiar estado de empresa | POST | `/api/v1/admin/empresas/{id}/estado` | Administrador (`revocar` o `restablecer`, con motivo) |
| Empresas disponibles | GET | `/api/v1/sesion/empresas` | Usuario autenticado (solo las autorizadas a él) |
| Elegir empresa activa | POST | `/api/v1/sesion/empresa` | Usuario autenticado |
| Correos | GET | `/api/v1/admin/notificaciones` | Administrador (solo lectura) |
| Configuración de administración | PUT | `/api/v1/admin/configuracion` | Administrador de plataforma |

### 11.6 Seguridad del acceso administrativo

- Autenticación fuerte con MFA obligatorio para administradores (el PIN del prototipo no sirve para producción).
- Sesiones de administrador con expiración corta y reautenticación para acciones sensibles (revocar, autorizar).
- Límite de intentos y bloqueo temporal; alertas de inicio de sesión inusual.
- Registro de cada acción administrativa en `auditoria`, con inmutabilidad.
- Separación de funciones: quien solicita no puede aprobar su propia solicitud.

### 11.7 Migración desde el prototipo

- El estado global (`empresas`, `solicitudes`, `buzon`, administrador) se exporta como JSON con `schemaVersion`.
- Las empresas del prototipo pasan a `empresa`. La empresa "demo" puede conservarse como empresa de prueba o eliminarse.
- Los datos de evaluación guardados por clave (`ley21719_assessment_v1_<empresa>`) se asocian a su `empresa_id`. Los que no tengan empresa se asignan a una empresa explícita antes de importar.
- El PIN del prototipo no se migra: el administrador de plataforma se da de alta con MFA.

### 11.8 Decisiones adicionales

8. **Dominio de correo** para envíos (por ejemplo, `no-responder@...`) y quién recibe las solicitudes.
9. **Verificación del solicitante**: ¿basta con correo corporativo y confirmación por enlace, o se exige validar el RUT de la empresa con un registro externo?
10. **Retención** de datos de una empresa revocada (¿cuánto tiempo se conservan?) y procedimiento para eliminarlos a solicitud.
11. **Cuatro ojos** para autorizar empresas: ¿un segundo administrador debe aprobar cada autorización?

---

## 12. Actualización v0.3 (alineada con el front y el Excel del 20/10/2026)

Esta sección reemplaza cualquier cifra anterior del documento. Las cifras vigentes son: 18 dimensiones, 171 controles, 198 requisitos, 67 actividades en el catálogo y 25 campos por actividad en el RAT.

### 12.1 Fuente única de requisitos

- El catálogo de requisitos y controles se carga desde una sola fuente. Hoy es el Excel de referencia, y en el backend será una tabla versionada.
- **Los IDs de requisito son únicos y estables.** Nunca se reutiliza un código aunque el requisito se elimine. Una auditoría de duplicados debe correr en cada carga del catálogo. Este punto surgió de un error real: el front agregaba requisitos por separado y reutilizaba códigos del Excel (4.7 y 5.7), lo que mezclaba puntajes. El front ya lee solo el catálogo.
- Requisitos incorporados en esta versión:

| Código | Dimensión | Tema | Controles que lo alimentan |
|---|---|---|---|
| 4.8 | D4 Derechos | Localizar los datos de un titular en todos los sistemas y proveedores | 5.33, 8.15 |
| 4.9 | D4 Derechos | Portabilidad en formato estructurado y de uso común | 5.34 |
| 4.10 | D4 Derechos | Supresión que llega a respaldos y proveedores | 8.13, 5.20 |
| 5.9 | D5 Principios | Tabla de retención por categoría, ejecutada y registrada | 5.33, 8.10 |
| 5.10 | D5 Principios | Evidencia de cada eliminación | 8.10 |
| 7.6 | D7 Evaluación de impacto | Inventario de decisiones automatizadas que afectan a personas | 5.34 |
| 7.7 | D7 Evaluación de impacto | Información sobre la lógica y revisión humana de decisiones automatizadas | 5.34, 5.8 |

Todos figuran como "Obligatorio Ley 21.719" y deben validarse con el área legal del cliente.

### 12.2 Registro de actividades (RAT): ciclo de vida y 25 campos

La tabla `rat_registro` agrega estos campos a los 15 existentes:

| Campo | Tipo | Valores |
|---|---|---|
| `fuente_datos` | enum | Directamente del titular; Formulario o canal digital; Sistema interno; Tercero o proveedor; Fuente pública |
| `como_se_recolectan` | texto | Canal habitual del proceso |
| `ubicacion_almacenamiento` | enum | Chile; Fuera de Chile; Chile y fuera de Chile; Por definir |
| `forma_eliminacion` | texto | Método de eliminación y registro |
| `encargados` | texto | Proveedores que tratan datos por cuenta de la empresa |
| `decisiones_automatizadas` | bool | Sí o No |
| `nivel_riesgo` | enum | Bajo; Medio; Alto |
| `requiere_eipd` | enum | No; Sí; Por evaluar |
| `derechos_arcop` | texto | Derechos que aplican a la actividad |
| `estado` | enum | Vigente; En revisión; Dada de baja |

La **completitud** se calcula sobre los 25 campos, no sobre 15: `completitud = campos_llenos / 25`. Las filas anteriores a esta versión quedan incompletas hasta que se completen los campos nuevos; el indicador RAT del Excel baja por esta razón y no por un error de cálculo.

Al elegir una actividad del catálogo, el sistema propone valores para estos campos (fuente según el tema, ubicación según transferencias, nivel de riesgo según datos sensibles y transferencias, EIPD según el nivel de riesgo). El usuario puede modificarlos; la propuesta no es una decisión legal.

### 12.3 Meta objetivo y perfil de riesgo

Campos nuevos en `organizacion`:

| Campo | Tipo | Regla |
|---|---|---|
| `meta_objetivo` | numeric(3,2) | Entre 1,00 y 5,00. Valor vigente |
| `meta_sugerida` | numeric(3,2) | Calculada: base 4,00 + 0,25 por cada criterio de riesgo que aplica, con tope de 4,50 |
| `meta_ajustada` | bool | Verdadero si la meta vigente difiere de la sugerida |
| `meta_justificacion` | texto | Obligatoria si `meta_ajustada` es verdadero |
| `criterio_datos_sensibles_gran_escala` | bool | Criterio de riesgo 1 |
| `criterio_decisiones_automatizadas` | bool | Criterio de riesgo 2 |
| `criterio_terceros_transferencias` | bool | Criterio de riesgo 3 |
| `criterio_videovigilancia_biometria` | bool | Criterio de riesgo 4 |

Reglas de negocio: el ajuste de la meta requiere justificación escrita y queda en `auditoria` con el valor anterior. El fundamento (escala COBIT 5 y enfoque ISO/IEC 27005) se muestra en el informe. Es un criterio de gestión propuesto; no es un requisito numérico de la ley.

### 12.4 Historial de evaluaciones

Nueva tabla `historial_evaluacion`:

| Campo | Tipo | Notas |
|---|---|---|
| `id` | UUID | |
| `organizacion_id` | UUID | |
| `fecha` | date | Único por organización: registrar de nuevo el mismo día reemplaza el registro |
| `puntaje_total` | numeric(4,3) | |
| `nivel` | texto | Calculado |
| `meta` | numeric(3,2) | Meta vigente en esa fecha |
| `cobertura` | numeric(4,3) | Requisitos evaluados / incluidos |
| `requisitos_incluidos` | int | |
| `requisitos_evaluados` | int | |
| `puntajes_dimension` | jsonb | Un valor por dimensión (D1 a D18) |
| `registrado_por` | UUID | |
| `registrado_en` | timestamp | |

El cambio entre dos registros se calcula en consulta, no se guarda. Las gráficas de evolución salen de esta tabla.

### 12.5 Entrevistas y registros de puntaje

Se mantiene el modelo de la sección 2.4 con dos precisiones tomadas del front:
- Cada entrevista puede cubrir varias dimensiones (relación N:M en `entrevista_dimension`).
- Los puntajes propuestos por control se guardan como registros ligados a la entrevista (`entrevista_control`). Al editar una entrevista, sus registros se reemplazan completos, para que nunca queden puntajes huérfanos.

### 12.6 Informe y exportaciones

- El informe PDF abre con el diagnóstico (semáforo, brechas, fuentes, mapa de calor y matriz de priorización), sigue con la evolución histórica y termina con el detalle. En el backend debe generarse en servidor con el mismo orden.
- La exportación a Excel por página exporta solo las hojas de esa página: resumen, dimensiones, marcos, requisitos, controles, entrevistas, respuestas, controles entrevistados o la tabla activa (RAT, ARCOP o transferencias).

### 12.7 Pendientes antes de la migración

1. Validación legal de los requisitos 4.8 a 4.10, 5.9, 5.10, 7.6 y 7.7, y de los plazos citados.
2. Completar los 10 campos nuevos en las 6 filas de RAT existentes.
3. Decidir si el Excel sigue siendo el catálogo de referencia o si la fuente será solo la base de datos.
