# AQUASHIELD · ExportDesk
## Resumen Técnico-Ejecutivo del Sistema de Auditoría Comex
**Versión:** 2.0 | **Fecha:** Abril 2026 | **Equipo:** Comercio Exterior — Aquachile S.A.

---

## 1. Resumen Ejecutivo

### El Problema
El equipo de Comercio Exterior gestionaba la auditoría de exportaciones en múltiples planillas Excel desconectadas entre sí. Cruzar manualmente los datos de SAP, los Bill of Lading marítimos y el estado de legalización DUS tomaba horas de trabajo repetitivo y generaba errores humanos recurrentes:

- **Sin visibilidad centralizada** del estado real de cada pedido de exportación.
- **Sin métricas de eficiencia** para medir tiempos de gestión por analista.
- **Sin trazabilidad** del ciclo completo: desde la factura SAP hasta el DUS legalizado.
- **Incapacidad de distinguir rápidamente** pedidos críticos (vencidos, con error de fecha) de los que avanzan normalmente.

### La Solución: AQUASHIELD · ExportDesk
Sistema propio de auditoría de exportaciones compuesto por:
1. **Motor ETL en Python** — procesa y cruza automáticamente todas las fuentes de datos.
2. **Dashboard Web interactivo** — panel visual en tiempo real con filtros, gráficos y exportación Excel.

### Impacto Estimado
| Indicador | Antes | Con ExportDesk |
|---|---|---|
| Tiempo para consolidar el reporte | ~2-3 horas semanales | < 5 minutos |
| Pedidos sin asignación de responsable | Frecuente | Automatizado (3 niveles de fallback) |
| Visibilidad DUS en tiempo real | Nula | Integrada y filtrable |
| Detección de errores de fecha SAP | Manual | Automática (💀 Rojo Crítico) |
| Seguimiento por analista | Sin gráfico | Gráfico de carga Procesados vs. Pendientes |
| Filtro por equipo (Congelado/Fresco) | Imposible | Integrado en todos los módulos |
| Trazabilidad del ciclo exportador | Parcial | Completa (Factura → BL → DUS) |
| **Tiempo de desarrollo del sistema** | Meses (sin apoyo) | **Días** con metodología correcta |

---

## 2. Riesgos Operacionales Mitigados

Esta es la dimensión más crítica del proyecto desde el punto de vista del negocio. La falta de un sistema de control centralizado exponía a la empresa a riesgos concretos y de alto impacto económico:

### 2.1 Cierre de Cuentas Navieras por No Pago
Las navieras (MSC, Hapag-Lloyd, Maersk, etc.) gestionan los fletes a través de cuentas de crédito. Un pedido cuyo BL no se gestiona a tiempo puede derivar en:
- **Facturas vencidas no identificadas** por falta de visibilidad del estado de los BL.
- **Bloqueo de cuenta naviera** — impide embarcar nueva carga hasta regularizar el saldo.
- **Sobrecostos de demurrage** (multa por días adicionales de uso del contenedor) que se acumulan silenciosamente cuando nadie hace seguimiento.

> Con ExportDesk, cualquier pedido sin BL después del plazo SLA aparece inmediatamente en rojo, con el analista responsable identificado.

### 2.2 No Legalización del DUS y Multas Aduaneras
El DUS (Declaración Única de Salida) **debe ser legalizado** dentro del plazo que establece el Servicio Nacional de Aduanas de Chile. Su incumplimiento conlleva:
- **Multas por infracción aduanera**, cuyo monto puede ser proporcional al valor de la exportación.
- **Bloqueo del RUT exportador** en casos reincidentes, lo que impide futuras operaciones de comercio exterior.
- **Retrasos en devolución de IVA exportador**, afectando directamente el flujo de caja de la empresa.

> Con el módulo de Legalización DUS, cada pedido sin legalizar es visible en tiempo real. Las observaciones documentan los casos especiales (traspasos entre analistas, situaciones pendientes de contraparte) para que ningún DUS quede sin gestionar.

### 2.3 Pérdida de Trazabilidad por Rotación de Personal
Sin un sistema centralizado, el conocimiento operacional ("este pedido lo tiene X analista", "este cliente siempre va por terrestre") existe **solo en la memoria del equipo**. Una salida de personal implica pérdida de información crítica.

> ExportDesk persiste toda la información en bases controladas: directorio de equipo, maestro de asignación, observaciones DUS, historial de memoria. El sistema no depende de ninguna persona en particular.

### 2.4 Errores de Digitación No Detectados en SAP
Fechas futuras mal ingresadas en SAP (ej: 2026 → 2062) hacían que pedidos quedaran perpetuamente como "vigentes" sin que nadie los identificara como errores.

> El semáforo 💀 **Rojo Crítico** marca automáticamente estos casos y los segrega para revisión, sin que el analista tenga que buscarlos manualmente.

### 2.5 Resumen de Riesgos Cubiertos

| Riesgo | Frecuencia Estimada | ¿Cubierto? |
|---|---|---|
| BL vencido sin detectar | Alta (sin sistema) | ✅ Semáforo + fecha creación BL |
| DUS no legalizado a tiempo | Media | ✅ Módulo DUS + observaciones |
| Bloqueo cuenta naviera | Media (consecuencia) | ✅ Prevención por visibilidad anticipada |
| Multa aduanera por DUS tardío | Baja-Media | ✅ T. Legalización + estado DUS en tiempo real |
| Pérdida info por rotación | Alta (sin sistema) | ✅ Persistencia en maestros y memoria |
| Error fecha SAP no detectado | Media | ✅ 💀 Rojo Crítico automático |
| Reporte manual incorrecto | Alta (sin sistema) | ✅ Automatizado con cruce de fuentes |

---

## 3. Alcance del Sistema

### Módulos Activos
| Módulo | Función |
|---|---|
| **Auditoría Comex** | Control de BL vs. SAP para pedidos de exportación |
| **Legalización DUS** | Seguimiento del estado de Declaración Única de Salida |
| **DUS Observaciones** | Notas por pedido para casos especiales o traspasos entre analistas |
| **Gráfico de Carga** | Distribución de trabajo por analista (Procesados / Pendientes / Sin Zarpe) |
| **Memoria Acumulada** | Historial persistente de pedidos procesados para métricas históricas |
| **Directorio de Equipo** | Asignación de analistas a grupos (Congelado / Fresco) |

### Fuentes de Datos Integradas
| Archivo | Origen | Contenido |
|---|---|---|
| `export.XLSX` | SAP | Pedidos, fechas de factura, folio, cliente, producto |
| `bills_of_lading-*.xlsx` | Naviera / Sistema | Número de BL, fecha de creación (uno o múltiples archivos) |
| `3.Asignacion Clientes.xlsx` | Comex | Mapeo cliente → analista responsable / Terrestre |
| `4.Procesados Manualmente.xlsx` | Comex | Excepciones validadas con motivo documentado |
| `5.Analista que facturo.XLSX` | SAP | N° Factura + Pedido → código SAP del analista real |
| `7. Legalización DUS_*.xlsx` | Aduana / Agente | Estado de legalización DUS por pedido |

---

## 3. Arquitectura del Sistema

```
FUENTES DE DATOS (Excel)
┌─────────────┐   ┌──────────────────────┐   ┌────────────────────────┐
│ export.XLSX │   │bills_of_lading-*.xlsx│   │7. Legalización DUS.xlsx│
│   (SAP)     │   │ (uno o múltiples BL) │   │  (reporte DUS / estado)│
└──────┬──────┘   └──────────┬───────────┘   └────────────┬───────────┘
       │                     │                             │
       ▼                     ▼                             ▼
┌──────────────────────────────────────────────────────────────────────┐
│            acc_auditor.py — Motor ETL Python                         │
│  Filtro @ # → Limpieza → Cruce → Reglas Negocio → Export Excel      │
└──────────────────────────┬───────────────────────────────────────────┘
                           │  Auditoria_Comex_Final.xlsx
                           ▼
┌──────────────────────────────────────────────────────────────────────┐
│        AQUASHIELD · ExportDesk — Dashboard Web (app_v2.js)          │
│  Carga online → KPIs → Tabla filtrable → Gráficos → Export grupal  │
└──────────────────────────────────────────────────────────────────────┘
```

---

## 4. Lógica de Carga y Cruce de Datos

### 4.1 Normalización de Llaves
Todos los N° de Pedido se limpian antes de cualquier cruce para garantizar coincidencia exacta:
```
"40569129.0" → "40569129"    (elimina decimal de Excel)
"040567890"  → "40567890"    (elimina ceros iniciales)
```

### 4.2 Unificación Multi-BL
Se procesan **todos** los archivos BL disponibles en la carpeta automáticamente. Si un mismo pedido aparece en más de un BL (por correcciones), se conserva el **registro más reciente** (útil cuando se sube una versión corregida).

### 4.3 Jerarquía de Cruces
Los cruces se realizan en cascada, siempre conservando todos los pedidos SAP (left join):

```
Pedidos SAP (base)
  ├── + BL               → por N° Pedido
  ├── + Analista SAP     → por N° Factura (prioridad 1: quien realmente facturó)
  ├── + Analista SAP     → por N° Pedido  (prioridad 2: respaldo)
  ├── + Maestro Clientes → por nombre cliente normalizado (prioridad 3: fallback)
  ├── + Procesados Manuales → por N° Pedido (excepciones documentadas)
  └── + DUS Legalizado   → por N° Pedido (estado de cierre aduanero)
```

### 4.4 Regla de Prioridad DUS
Si el reporte DUS indica que un pedido está **legalizado** (`✅`), ese estado tiene prioridad sobre cualquier otra clasificación interna. Un pedido puede aparecer como "Pendiente BL" pero estar correctamente cerrado si su DUS ya fue legalizado.

---

## 5. Filtros de Seguridad

Aplicados sobre la columna de **Descripción del producto** en SAP, antes de cualquier cruce:

| Marcador | Línea de Producto | Acción |
|---|---|---|
| `@` | Congelado | ✅ **Incluir** (dentro del alcance de auditoría) |
| `#` | Fresco puro | ❌ **Excluir** (fuera del alcance de este módulo) |
| Ninguno | Sin clasificar | ❌ Excluido por defecto |

Esto garantiza que solo se auditen los pedidos de la línea **Congelado** y elimina el ruido de otros productos.

---

## 6. Reglas de Negocio y Semáforo SLA

### 6.1 Semáforo de Vencimiento
Basado en los días transcurridos desde la Fecha de Factura hasta hoy:

| Estado | Días | Significado |
|---|---|---|
| 💀 Rojo Crítico | Fecha futura | Error de digitación en SAP (ej: 2026 → 2062) |
| 🟢 Verde | 0 – 3 días | Dentro del plazo estándar |
| 🟠 Naranja | 4 – 7 días | En riesgo de incumplimiento SLA |
| 🔴 Rojo | > 7 días | SLA vencido — requiere gestión urgente |
| ⚪ N/A | Sin fecha | No facturado aún |

### 6.2 Árbol de Decisión — Estado Final del Pedido

```
¿Tiene folio y fecha de factura?
│
├── NO  →  📦 Pendiente Facturación
│
└── SÍ  ──→ ¿Fecha futura (error SAP)?
            │
            ├── SÍ  →  ⚠️ Error Fecha
            │
            └── NO  ──→ ¿DUS Legalizado?
                        │
                        ├── SÍ  →  ✅ DUS Legalizado   ← cierre total del ciclo
                        │
                        └── NO  ──→ ¿Tiene BL subido?
                                    │
                                    ├── SÍ  →  ✅ Procesado
                                    │
                                    └── NO  ──→ ¿Es cliente Terrestre?
                                                │
                                                ├── SÍ  →  🚚 Procesado Terrestre
                                                │
                                                └── NO  ──→ ¿Tiene motivo manual?
                                                            │
                                                            ├── SÍ  →  ⚠️ Validado Manualmente
                                                            │
                                                            └── NO  →  ❌ Pendiente
```

---

## 7. Métricas de Eficiencia del Ciclo Exportador

### 7.1 DEMORA
**¿Cuántos días lleva este pedido esperando?**
```
DEMORA = Hoy − Fecha Factura
```
Permite identificar pedidos que llevan semanas sin avanzar.

### 7.2 T. GESTIÓN — Eficiencia Documental Interna
**¿Cuántos días tardó el analista en subir el BL desde que se emitió la factura?**
```
T. GESTIÓN = Fecha Creación BL − Fecha Factura
```
- `0` días → máxima eficiencia (mismo día)
- Valores negativos → BL pre-registrado antes de la factura
- Valores altos → demora en gestión documental interna

### 7.3 T. LEGALIZACIÓN — Ciclo Completo Exportador *(nuevo)*
**¿Cuántos días tardó el ciclo completo desde la factura hasta el cierre aduanero?**
```
T. LEGALIZACIÓN = Fecha DUS − Fecha Factura
```
Solo calculado para pedidos con DUS legalizado. Mide la **eficiencia total del proceso exportador** de extremo a extremo. Es el indicador estratégico más importante para negociación con aduanas y navieras.

---

## 8. Funcionalidades del Dashboard Web

### KPIs Principales (tarjetas superiores)
| KPI | Descripción |
|---|---|
| **Total** | Total de pedidos cargados en la sesión |
| **Pendientes** | Pedidos con `❌ Pendiente`, con desglose ❄️ Congelado / 🌿 Fresco |
| **% Cumplimiento** | `Procesados / Total` — indicador verde/amarillo/rojo |
| **En Riesgo** | Pedidos en semáforo 🟠 Naranja o 🔴 Rojo |

### Filtros Disponibles
- **Por Estado:** multiselección de estados (Procesado, Pendiente, DUS Legalizado, etc.)
- **Por Grupo:** ❄️ Congelado | 🌿 Fresco | 👥 Todos
- **Por Búsqueda:** texto libre sobre cualquier campo (pedido, cliente, responsable)
- **Combinado:** todos los filtros actúan en intersección

### Exportación Inteligente
- El botón **Exportar** descarga **exactamente lo que se ve en pantalla** (respeta búsqueda + grupo + estado)
- El nombre del archivo refleja los filtros activos: `Legalización DUS_Congelado_busqueda_em_2026-04-06.xlsx`
- La columna **OBSERVACIONES** se incluye automáticamente para pedidos no legalizados

### Sistema de Grupos (Congelado / Fresco)
- Cada analista se asigna a un grupo desde el modal de Equipo
- El filtro de grupo aplica globalmente: tabla, KPIs, gráfico de carga y exportaciones
- Las pills `❄️ N | 🌿 N` en la tarjeta de Pendientes muestran siempre el desglose de ambos grupos, independientemente del filtro activo

### Módulo DUS Observaciones
- Widget exclusivo del módulo DUS
- Carga observaciones desde el mismo reporte Excel (columna "Observaciones")
- Ingreso manual por pedido para casos especiales (ej: "LCONEJEROSM facturó, VVENEGAS tramita")
- La observación desaparece automáticamente del reporte cuando el DUS queda legalizado
- Exportación incluye la columna OBSERVACIONES solo para pedidos pendientes

---

## 9. Estructura del Output — Columnas Finales

| # | Columna | Descripción |
|---|---|---|
| 1 | `ESTATUS_FINAL` | Estado con emoji según árbol de decisión |
| 2 | `DEMORA` | Días desde factura hasta hoy |
| 3 | `T. GESTIÓN` | Días factura → subida BL |
| 4 | `T. LEGALIZACIÓN` | Días factura → DUS legalizado (ciclo total) |
| 5 | `N° PEDIDO` | Número de pedido limpiado |
| 6 | `NOMBRE SOLICITANTE` | Cliente / consignatario |
| 7 | `RESPONSABLE` | Código SAP del analista |
| 8 | `FECHA FACTURA` | Fecha de emisión de factura SAP |
| 9 | `FECHA CREACIÓN BL` | Fecha de subida del Bill of Lading |
| 10 | `FECHA DUS` | Fecha de legalización DUS |
| 11 | `ESTADO DUS` | Estado crudo del DUS (para referencia) |
| 12 | `MOTIVO` | Justificación de excepciones manuales |

---

## 10. Script Python ETL Completo

```python
"""
AQUASHIELD · ExportDesk — Motor ETL de Auditoría Comex v2.0
Integra: SAP | Bill of Lading (multi-archivo) | Legalización DUS
"""
import pandas as pd, glob, os
from datetime import datetime

FOLDER_PLANILLAS   = 'PLANILLAS'
SAP_FILE           = os.path.join(FOLDER_PLANILLAS, 'export.XLSX')
ASIGNACION_FILE    = os.path.join(FOLDER_PLANILLAS, '3.Asignacion Clientes.xlsx')
PROCESADOS_FILE    = os.path.join(FOLDER_PLANILLAS, '4.Procesados Manualmente.xlsx')
ANALISTAS_SAP_FILE = os.path.join(FOLDER_PLANILLAS, '5.Analista que facturo.XLSX')
DUS_FILE_PATTERN   = os.path.join(FOLDER_PLANILLAS, '7. Legalización DUS*.xlsx')
OUTPUT_FILE        = 'Auditoria_Comex_Final.xlsx'
FECHA_HOY          = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)

def limpiar_pedido(val):
    if pd.isna(val): return ""
    return str(val).split('.')[0].strip().lstrip('0')

def calcular_semaforo(row):
    fecha = row.get('FECHA FACTURA')
    if pd.isna(fecha): return "⚪ N/A"
    if fecha > FECHA_HOY: return "💀 Rojo Crítico"
    dias = (FECHA_HOY - fecha).days
    return "🟢 Verde" if dias<=3 else "🟠 Naranja" if dias<=7 else "🔴 Rojo"

def determinar_estatus(row):
    if pd.isna(row.get('Folio Factura')) or pd.isna(row.get('FECHA FACTURA')):
        return "📦 Pendiente Facturación"
    if row.get('SEMÁFORO') == "💀 Rojo Crítico": return "⚠️ Error Fecha"
    if row.get('DUS_LEGALIZADO', False):          return "✅ DUS Legalizado"
    if pd.notna(row.get('BL')):                   return "✅ Procesado"
    if 'Terrestre' in str(row.get('RESPONSABLE','')): return "🚚 Procesado Terrestre"
    if pd.notna(row.get('MOTIVO')):               return "⚠️ Validado Manualmente"
    return "❌ Pendiente"

def procesar_auditoria():
    print(f"AQUASHIELD · ExportDesk — {FECHA_HOY.strftime('%d/%m/%Y')}")

    # SAP + Filtros @ # (Congelado / Fresco)
    df_sap = pd.read_excel(SAP_FILE)
    df_sap = df_sap.rename(columns={c: c.encode('ascii','ignore').decode('ascii') for c in df_sap.columns})
    df_sap = df_sap[df_sap['Descripcin'].str.contains('@', na=False)]
    df_sap = df_sap[~df_sap['Descripcin'].str.contains('#', na=False)]
    df_sap['PEDIDO_CLEAN'] = df_sap['NPedido'].apply(limpiar_pedido)
    df_sap = df_sap.drop_duplicates(subset=['PEDIDO_CLEAN'], keep='first')

    # BL multi-archivo (última entrada por pedido)
    bl_files = glob.glob(os.path.join(FOLDER_PLANILLAS, 'bills_of_lading-*.xlsx'))
    if bl_files:
        df_bl = pd.concat([pd.read_excel(f) for f in bl_files], ignore_index=True)
        df_bl['PEDIDO_CLEAN'] = df_bl['Pedido'].apply(limpiar_pedido)
        df_bl = df_bl.drop_duplicates(subset=['PEDIDO_CLEAN'], keep='last')
    else:
        df_bl = pd.DataFrame(columns=['PEDIDO_CLEAN','BL','FechaCreacion'])

    # Maestros
    df_asignacion = pd.read_excel(ASIGNACION_FILE)
    df_manual     = pd.read_excel(PROCESADOS_FILE)
    df_manual['PEDIDO_CLEAN'] = df_manual['PEDIDO FLUJO'].apply(limpiar_pedido)
    df_analistas = pd.DataFrame()
    if os.path.exists(ANALISTAS_SAP_FILE):
        df_analistas = pd.read_excel(ANALISTAS_SAP_FILE)
        fc = next((c for c in df_analistas.columns if 'Factura' in c or 'Folio' in c), None)
        pc = next((c for c in df_analistas.columns if 'Pedido' in c or 'NPedido' in c), None)
        if fc: df_analistas['FOLIO_JOIN']  = df_analistas[fc].apply(limpiar_pedido)
        if pc: df_analistas['PEDIDO_JOIN'] = df_analistas[pc].apply(limpiar_pedido)

    # DUS Legalizado (archivo más reciente)
    dus_files = sorted(glob.glob(DUS_FILE_PATTERN), key=os.path.getmtime, reverse=True)
    if dus_files:
        df_dus = pd.read_excel(dus_files[0])
        df_dus['PEDIDO_CLEAN']  = df_dus['N° PEDIDO'].apply(limpiar_pedido)
        ec = next((c for c in df_dus.columns if 'ESTADO' in c.upper() and 'GENERAL' in c.upper()), None)
        df_dus['DUS_LEGALIZADO'] = df_dus[ec].str.contains('✅', na=False) if ec else False
        fc2 = next((c for c in df_dus.columns if 'FECHA' in c.upper()), None)
        df_dus['FECHA_DUS'] = pd.to_datetime(df_dus[fc2], errors='coerce') if fc2 else pd.NaT
        dc = next((c for c in df_dus.columns if 'DETALLE' in c.upper()), None)
        df_dus['ESTADO_DUS_RAW'] = df_dus[dc] if dc else ''
        df_dus = df_dus.sort_values('DUS_LEGALIZADO', ascending=False).drop_duplicates('PEDIDO_CLEAN', keep='first')
    else:
        df_dus = pd.DataFrame(columns=['PEDIDO_CLEAN','DUS_LEGALIZADO','FECHA_DUS','ESTADO_DUS_RAW'])

    # Cruces en cascada
    df_sap['CLIENTE_NORM'] = df_sap['Nombre Solicitante'].str.lower().str.strip()
    df_asignacion['CLIENTE_NORM'] = df_asignacion['Cliente'].str.lower().str.strip()
    terrestres = df_asignacion[df_asignacion['Responsable'].str.contains('Terrestre',na=False,case=False)]['CLIENTE_NORM'].tolist()
    sfc = next((c for c in df_sap.columns if 'Factura' in c or 'Folio' in c), 'Folio Factura')
    df_sap['FOLIO_JOIN'] = df_sap[sfc].apply(limpiar_pedido)

    df = df_sap.merge(df_bl[['PEDIDO_CLEAN','BL','FechaCreacion']], on='PEDIDO_CLEAN', how='left')
    if not df_analistas.empty and 'FOLIO_JOIN' in df_analistas.columns:
        df = df.merge(df_analistas[['FOLIO_JOIN','Creado por']], on='FOLIO_JOIN', how='left')
    else: df['Creado por'] = pd.NA
    if not df_analistas.empty and 'PEDIDO_JOIN' in df_analistas.columns:
        da2 = df_analistas[['PEDIDO_JOIN','Creado por']].rename(columns={'Creado por':'Creado_por_Ped'}).drop_duplicates('PEDIDO_JOIN')
        df = df.merge(da2, left_on='PEDIDO_CLEAN', right_on='PEDIDO_JOIN', how='left')
    else: df['Creado_por_Ped'] = pd.NA
    df = df.merge(df_asignacion[['CLIENTE_NORM','Responsable']], on='CLIENTE_NORM', how='left')
    df = df.merge(df_manual[['PEDIDO_CLEAN','MOTIVO']], on='PEDIDO_CLEAN', how='left')
    df = df.merge(df_dus[['PEDIDO_CLEAN','DUS_LEGALIZADO','FECHA_DUS','ESTADO_DUS_RAW']], on='PEDIDO_CLEAN', how='left')
    df['DUS_LEGALIZADO'] = df['DUS_LEGALIZADO'].fillna(False)

    def responsable(r):
        if r['CLIENTE_NORM'] in terrestres: return "🚚 Terrestre"
        if pd.notna(r.get('Creado por')):   return r['Creado por']
        if pd.notna(r.get('Creado_por_Ped')): return r['Creado_por_Ped']
        if pd.notna(r.get('Responsable')):  return r['Responsable']
        return "Sin Asignar"
    df['RESPONSABLE_FINAL'] = df.apply(responsable, axis=1)

    # Renombres y KPIs
    df = df.rename(columns={'NPedido':'N° PEDIDO','Nombre Solicitante':'NOMBRE SOLICITANTE',
                             'Fecha Factura':'FECHA FACTURA','FechaCreacion':'FECHA CREACIÓN BL',
                             'RESPONSABLE_FINAL':'RESPONSABLE','FECHA_DUS':'FECHA DUS','ESTADO_DUS_RAW':'ESTADO DUS'})
    df['FECHA FACTURA']     = pd.to_datetime(df['FECHA FACTURA'], errors='coerce')
    df['FECHA CREACIÓN BL'] = pd.to_datetime(df['FECHA CREACIÓN BL'], errors='coerce')
    df['SEMÁFORO']          = df.apply(calcular_semaforo, axis=1)
    df['ESTATUS_FINAL']     = df.apply(determinar_estatus, axis=1)
    df['DEMORA']            = (FECHA_HOY - df['FECHA FACTURA']).dt.days
    df.loc[df['ESTATUS_FINAL']=='📦 Pendiente Facturación','DEMORA'] = 0
    df.loc[df['DEMORA']>10000,'DEMORA'] = 0
    df['T. GESTIÓN']       = (df['FECHA CREACIÓN BL'] - df['FECHA FACTURA']).dt.days
    df['T. LEGALIZACIÓN']  = (df['FECHA DUS'] - df['FECHA FACTURA']).dt.days
    df.loc[~df['DUS_LEGALIZADO'], 'T. LEGALIZACIÓN'] = pd.NA

    cols = ['ESTATUS_FINAL','DEMORA','T. GESTIÓN','T. LEGALIZACIÓN','N° PEDIDO',
            'NOMBRE SOLICITANTE','RESPONSABLE','FECHA FACTURA','FECHA CREACIÓN BL',
            'FECHA DUS','ESTADO DUS','MOTIVO']
    df[[c for c in cols if c in df.columns]].to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Reporte generado: {OUTPUT_FILE}")

if __name__ == "__main__":
    try: procesar_auditoria()
    except Exception as e:
        import traceback; print(f"❌ ERROR: {e}"); traceback.print_exc()
```

---

## 11. Hoja de Ruta — Próximos Pasos

| Prioridad | Mejora | Beneficio |
|---|---|---|
| Alta | Alertas automáticas por correo para pedidos en 🔴 Rojo > 7 días | Gestión proactiva sin revisar el dashboard |
| Alta | API de conexión directa a SAP (reemplazar Excel manual) | Eliminar el paso de exportación manual |
| Media | Histórico de T. Legalización por naviera | Identificar navieras con mejor performance |
| Media | Dashboard público (SharePoint/Teams) | Visibilidad para analistas sin acceso a la carpeta |
| Baja | Agregar nuevo grupo "Secos" si se incorpora esa línea | Solo crear el grupo en el modal de equipo |
| Baja | Notificación cuando el porcentaje de cumplimiento cae bajo el 80% | Alerta temprana de deterioro del SLA del equipo |

---

## 12. Glosario de Términos

| Término | Significado |
|---|---|
| **SAP** | Sistema ERP corporativo. Fuente maestra de pedidos y facturas |
| **BL (Bill of Lading)** | Documento de embarque marítimo. Confirma que la carga zarpó |
| **DUS** | Declaración Única de Salida. Documento aduanero que cierra oficialmente la exportación |
| **Legalizar DUS** | Proceso por el cual la aduana valida y cierra el DUS. Hito final del ciclo exportador |
| **SLA** | Service Level Agreement — plazo máximo acordado para gestionar un pedido |
| **Semáforo** | Indicador visual (🟢🟠🔴) que refleja el cumplimiento del SLA por pedido |
| **T. Gestión** | Tiempo en días entre la factura y la subida del BL (eficiencia interna) |
| **T. Legalización** | Tiempo en días entre la factura y el DUS legalizado (ciclo exportador completo) |
| **Memória** | Historial acumulado de pedidos procesados en sesiones anteriores del dashboard |
| **Congelado / Fresco** | Líneas de producto. Solo Congelado es auditado por este sistema (filtro `@`) |
