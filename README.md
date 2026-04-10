# 🛡️ AQUASHIELD · ExportDesk

**Motor de Auditoría y Control de Comercio Exterior**

Dashboard interactivo para la gestión, auditoría y control de pedidos de exportación. Integra datos SAP, Bill of Ladings, Legalización DUS y directorio de analistas en una interfaz glassmorphism moderna.

---

## ✨ Características

| Módulo | Descripción |
|--------|-------------|
| **Auditoría Comex** | Cruce automático SAP ↔ BL con semáforo SLA y detección de responsables |
| **Legalización DUS** | Control de estado de DUS con clasificación automática por estado aduanero |
| **Memoria Persistente** | Historial acumulado en localStorage con export/import y migración automática |
| **Gráfico de Analistas** | Carga por analista con filtros por mes, grupo y status |
| **Directorio de Equipo** | Asociación código SAP ↔ nombre completo con grupos (Congelado/Fresco) |
| **Observaciones DUS** | Notas por pedido con carga masiva desde Excel |
| **Exportación Excel** | Reportes con estilos, colores y gráficos embebidos (ExcelJS) |

## 🔒 Seguridad

- Sanitización XSS completa (`escHtml()`) en todos los datos de Excel → DOM
- 37 correcciones verificadas en 8 pasadas de auditoría técnica
- CDNs pineados a versiones específicas

## 🚀 Uso

### Web Dashboard (sin instalación)
1. Abrir `web/index.html` en el navegador
2. Arrastrar archivos Excel a las zonas correspondientes
3. Ejecutar auditoría

### Motor ETL Python (opcional)
```bash
pip install -r requirements.txt
python acc_auditor.py
```

## 📁 Estructura

```
├── web/
│   ├── index.html      # Dashboard principal
│   ├── app_v2.js       # Lógica JS (2,744 líneas auditadas)
│   └── styles.css      # Estilos glassmorphism
├── acc_auditor.py      # Motor ETL Python (434 líneas)
├── requirements.txt    # Dependencias Python
└── PLANILLAS/          # (gitignored) Datos de trabajo
```

## 🛠️ Stack Tecnológico

- **Frontend:** HTML5 + Vanilla JS + CSS (glassmorphism)
- **Librerías:** Chart.js 4.4.1 · SheetJS · ExcelJS 4.3.0 · Lucide Icons 0.344.0
- **Backend ETL:** Python 3 + Pandas + openpyxl
- **Persistencia:** localStorage con QuotaExceeded handling

## 📋 Auditoría Técnica

Sistema validado con **8 pasadas de auditoría** (37 correcciones):

| Categoría | Correcciones |
|-----------|-------------|
| Seguridad XSS | 7 |
| Integridad de datos | 8 |
| Optimización | 3 |
| UX/UI | 4 |
| CDN/Dependencias | 2 |
| Lógica de negocio | 13 |

---

*Desarrollado por Equipo Comex · Aquachile*
