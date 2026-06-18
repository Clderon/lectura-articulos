# Configuración de la Plantilla Excel

Guía para crear o preparar el Excel para que funcione con el pipeline.

---

## 📋 Estructura Requerida

El Excel debe tener **exactamente 2 hojas**:
1. "Revisión de Artículos"
2. "Lectura Selectiva"

---

## 🔧 Hoja 1: "Revisión de Artículos"

### Encabezado (Fila 3)
```
A          B                      C          D      E             F           G       H                I        J       K          L           M
N°         Artículo (Título)      Autor(es)  Año    Pertinencia   Actualidad  Rigor   Fuente (Calidad) Impacto Ética   Puntaje     Veredicto   Justificación
```

### Filas de Datos (4 en adelante)
```
Fila 4:  [número] [título]        [autor]   [año]  [0-2]         [0-2]       [0-2]   [0-2]            [0-2]   [0-2]   [fórmula]   [sí/no]     [texto]
Fila 5:  [número] [título]        [autor]   [año]  [0-2]         [0-2]       [0-2]   [0-2]            [0-2]   [0-2]   [fórmula]   [sí/no]     [texto]
...
```

### Fórmulas Importantes

#### Columna A (N°)
```excel
=SI.ERROR(INDICE(TablaArticulos[N°];FILA()-3);"")
```
- Trae el número automáticamente de una tabla nombrada
- **NO LA EDITES** (el script respeta esta fórmula)

#### Columna K (Puntaje Total)
```excel
=SUMA(E:J)
```
- Suma automáticamente E+F+G+H+I+J (máximo 12)
- **NO LA EDITES** (el script respeta esta fórmula)

#### Columna L (Veredicto)
```excel
Aprobado   (si IaC es objeto de estudio)
Excluido   (si IaC es solo herramienta)
```

---

## 📖 Hoja 2: "Lectura Selectiva"

### Encabezado (Fila 3)
```
A      B     C          D        E          F            G                   H              I                  J                K            L             M               N               O
N°     [vacío] Objetivo Método   Hallazgo   Limitación   Ref. textual/Secc.  Paradigma IaC  Herramienta IaC    Tipo de estudio  Variables def. Dimensiones    Indicadores med. Trabajo futuro  ¿Generalizable?
```

### Filas de Datos (4 en adelante)
```
Fila 4:  [auto] [vacío] [texto]    [texto]   [texto]     [texto]         [texto]            [texto]        [texto]          [texto]         [texto]        [texto]        [texto]         [texto]        [texto]
Fila 5:  [auto] [vacío] [texto]    [texto]   [texto]     [texto]         [texto]            [texto]        [texto]          [texto]         [texto]        [texto]        [texto]         [texto]        [texto]
...
```

### Formato Recomendado

**Para todas las celdas de contenido (C3:O43):**
- ✅ **Word Wrap activado:** Formato → Celdas → Alineación → "Ajustar texto"
- ✅ **Altura de fila:** 60 puntos (para que textos largos sean legibles)
- ✅ **Alineación:** Superior izquierda
- ✅ **Fuente:** Calibri 11

**Ejemplo (Excel):**
```
1. Selecciona C4:O43
2. Click derecho → Formato de celdas
3. Alineación → 
   - Ajuste de texto: ☑
   - Alineación vertical: Superior
4. OK
5. Formato → Altura de fila → 60
```

---

## ✅ Checklist de Setup

Antes de ejecutar el pipeline:

- [ ] Excel tiene 2 hojas: "Revisión de Artículos" y "Lectura Selectiva"
- [ ] Fila 3 en ambas hojas tiene los encabezados
- [ ] Columna A en "Revisión de Artículos" tiene fórmula `=SI.ERROR(...)`
- [ ] Columna K en "Revisión de Artículos" tiene fórmula `=SUMA(E:J)`
- [ ] Filas 4-43 están vacías (listas para datos)
- [ ] Columnas C-P en "Lectura Selectiva" tienen Word Wrap activado
- [ ] Altura de fila es 60 puntos en ambas hojas
- [ ] `.env` apunta al Excel correcto (ruta completa)

---

## 🎨 Formato Visual (Opcional)

Para mejorar la legibilidad:

### Colores de encabezado
- Fila 3: Fondo gris claro (RGB 217,217,217), texto negrita
- Resultado: Encabezados destacados

### Bordes
- Celdas de datos: Bordes delgados (1pt, gris)
- Resultado: Tabla clara

### Ancho de columnas
- A: 5
- B: 30
- C-P: 20
- Q+: 15

---

## 🚨 Problemas Comunes

### "El script no encuentra Excel"
- Verifica que `EXCEL_PATH` en `.env` tenga ruta **completa** (no relativa)
- Windows: `C:\Users\...` (usa `\` o `/`, ambos funcionan)
- Ejemplo: `C:/Users/usuario/Desktop/Revision.xlsx`

### "Datos aparecen en fila equivocada"
- Verifica que encabezado esté en **fila 3** (no fila 1 o 2)
- El script suma 3 para calcular dónde escribir: fila_encabezado + 1 = fila_datos

### "El Excel se daña / muestra errores de fórmula"
- Asegúrate de que columnas A y K **no tengan contenido manual**
- El script respeta las fórmulas, no las sobrescribe
- Si las dañaste: Edita el Excel manualmente y reinserta las fórmulas

### "Excel abierto en Excel Desktop vs OneDrive vs LibreOffice"
- Cierra el Excel mientras ejecutas el script
- OneDrive puede causar conflictos (guarda localmente primero)
- LibreOffice: Usa OpenOffice Calc (compatible pero no probado)

---

## 🔄 Actualización de Plantilla

Si necesitas cambiar la estructura:

1. **Agregar columnas:** Edita `deepseek_analyzer.py` línea ~386 para mapear nuevos campos
2. **Cambiar nombres de hojas:** Actualiza `SHEET1_NAME` y `SHEET2_NAME` en `.env`
3. **Mover encabezado:** Edita `deepseek_analyzer.py` línea ~462 (rango de búsqueda)

---

## 📥 Descargar Plantilla

No hay plantilla descargable, pero puedes crear una fácilmente:

1. Abre Excel
2. Crea 2 hojas: "Revisión de Artículos" y "Lectura Selectiva"
3. Copia los encabezados de arriba
4. Agrega las fórmulas en columnas A y K
5. Formatea (Word Wrap, altura 60)
6. Guarda como `xlsx` (no `xls`)
7. Actualiza `EXCEL_PATH` en `.env`

---

**Última actualización:** 2026-06-17
