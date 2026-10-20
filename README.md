# Assessment de cumplimiento Ley 21.719

Herramienta de levantamiento, diagnóstico y scoring de cumplimiento de la Ley N° 21.719 sobre protección de datos personales (Chile). Cubre ISO/IEC 27001:2022, NIST CSF 2.0, CIS Controls v8.1, COBIT 5, PCI DSS v4.0.1 y normativa chilena complementaria.

## Estructura

| Carpeta | Contenido |
|---|---|
| `frontend/` | Aplicación en un solo archivo HTML (sin servidor, sin dependencias externas). |
| `frontend/src/plantilla.html` | Código fuente de la aplicación. El HTML final se genera desde aquí. |
| `frontend/build/` | Script que genera el HTML final a partir de la plantilla y del catálogo. |
| `catalogo/` | Catálogo de requisitos, controles y dimensiones en Excel (fuente de referencia). |
| `catalogo/generador/` | Scripts que construyen el Excel. |
| `catalogo/fuentes/` | Textos y datos de origen: actividades del RAT, descripciones de controles, textos de requisitos e ISO 27001. |
| `docs/` | Especificación del backend (modelo de datos, API, reglas de cálculo, seguridad). |

## Estado

- Front: prototipo funcional de una sola página. Los datos se guardan en el navegador. No hay backend ni autenticación real.
- Catálogo: 18 dimensiones, 171 controles, 198 requisitos, 67 actividades para el RAT.
- Pendiente de validación legal: artículos y plazos citados, requisitos marcados como obligatorios, plazo y autoridad de notificación de vulneraciones, bases legales sugeridas.
- Pendiente de decisión: fuente oficial del catálogo (Excel o base de datos) y proveedor de correo y autenticación para producción.

## Advertencias

- Los datos de ejemplo y los valores sugeridos (plazos, bases legales, criterios de riesgo para la meta objetivo) son propuestas de consultoría, no conclusiones jurídicas.
- El fundamento de la meta objetivo (escala 1 a 5, criterios de riesgo, tope de 4,5) es un criterio de gestión propuesto y debe validarse con el cliente.
- Los scripts de `catalogo/generador/` y `frontend/build/` usan rutas de trabajo de la sesión original (`/tmp`, `/mnt/project`). Antes de reproducir la generación, ajustar esas rutas.

## Cómo regenerar el front

```
python frontend/build/generar_front.py
```

Lee el catálogo desde `catalogo/checklist_ley21719_integrado.xlsx` y escribe `frontend/checklist_ley21719_frontend.html`. Ajustar las rutas antes de ejecutar.

## Uso

Abrir `frontend/checklist_ley21719_frontend.html` en el navegador.
