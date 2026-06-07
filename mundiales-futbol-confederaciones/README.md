# Mundiales de Fútbol por Confederaciones (1930–2022)
**Proyecto Final — Maestría en Ciencias de Datos**
**Autor:** Carlos Julio Marín Santana

---

## Descripción

Proyecto de visualización interactiva que analiza **92 años de historia mundialista** (1930–2022) desde la perspectiva de las confederaciones de fútbol. Incluye un dashboard interactivo en Tableau Cloud, una historia narrativa con 9 puntos, y un informe técnico completo.

## Tecnologías

| Herramienta | Uso |
|---|---|
| **Python 3.x** | ETL y generación de datos |
| **openpyxl / pandas** | Manipulación de Excel |
| **python-docx** | Generación del informe técnico |
| **Tableau Cloud** | Visualización interactiva |
| **Power BI** | Prototipo inicial |

## Estructura del Proyecto

```
├── crear_tableau_mundiales.py        # ETL: genera Mundiales_Datos.xlsx
├── generar_informe.py                # Genera informe técnico Word/PDF
├── Mundiales_Datos.xlsx              # Dataset procesado (5 hojas)
├── Mundiales_Confederaciones.twbx    # Tableau Packaged Workbook
├── Informe_Tecnico_...docx           # Informe técnico (Word)
├── Informe_Tecnico_...pdf            # Informe técnico (PDF)
└── README.md
```

## Dataset

`Mundiales_Datos.xlsx` contiene 5 hojas:

| Hoja | Filas | Descripción |
|---|---|---|
| **Participaciones** | 490 | Una fila por país × edición |
| **Ediciones** | 22 | Historial completo de cada Mundial |
| **Resumen_Confederación** | 6 | KPIs por confederación |
| **Resumen_País** | 84 | KPIs por país |
| **Datos_Mapa** | 84 | Datos georreferenciados (ISO Alpha-3) |

## Visualizaciones en Tableau Cloud

1. **Dashboard Mundiales** — Vista general interactiva
2. **Campeonatos x Confederación** — Dominancia histórica
3. **Participaciones x Confederación** — Presencia global
4. **Medallero x Confederación** — Podios acumulados
5. **Mapa de Participaciones** — Distribución geográfica
6. **Top Países** — Ranking por participaciones
7. **Evolución Histórica** — Serie temporal 1930–2022
8. **Historial de Ediciones** — Tabla completa por torneo

## Cómo reproducir el dataset

```bash
pip install openpyxl python-docx
python crear_tableau_mundiales.py
```

Esto genera `Mundiales_Datos.xlsx` listo para importar en Tableau.

## Informe Técnico

El informe cubre 8 secciones (100 pts rubrica):
- Descripción del problema
- Descripción y preprocesamiento de datos
- Escenario de uso
- Abstracción de datos y tareas
- Justificación de marcas y canales (Munzner, 2014)
- Conclusiones

## Fuentes

- FIFA — [https://www.fifa.com](https://www.fifa.com)
- Wikipedia — Historia de la Copa Mundial de Fútbol
- ISO 3166-1 Alpha-3 — Códigos de países
- Munzner, T. (2014). *Visualization Analysis and Design*. CRC Press.

---

*Proyecto desarrollado para el curso de Visualización de Datos — Maestría en Ciencias de Datos*
