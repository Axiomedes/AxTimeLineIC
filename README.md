# AxTimeLineIC (AxTimeLine v2.0 or Improved Control) - Control Gráfico de Línea de Tiempo para Visual Basic 6

**AxTimeLineIC** es un control de usuario ActiveX (`UserControl` / `.ctl` / `.ocx`) diseñado para Visual Basic 6.0. Proporciona una representación visual moderna, fluida y altamente personalizable de líneas de tiempo, procesos por pasos, diagramas de flujo interactivos, bitácoras y wizards.

Construido sobre un motor gráfico vectorial **GDI+** con soporte de suavizado (*Anti-Aliasing*), sombreado difuso (*Soft Drop Shadows*), doble búfer en memoria (*DIB Section*) sin parpadeos (*Flicker-Free*) y compatibilidad nativa con fuentes de iconos e imágenes con canal alfa.

---

## 🌟 Características Principales

* **Renderizado Vectorial de Alta Calidad (GDI+):** Trazos suaves, bordes anti-aliasing y tipografías con subpixel grid fitting.
* **Cero Parpadeo (Flicker-Free):** Doble búfer en memoria mediante DIB Sections de 32 bits y volcado atómico vía `BitBlt`.
* **Múltiples Modos de Orientación (`tlOrientation`):**
  * `tlVertical` (0): Vista vertical clásica con columna de fecha y hora.
  * `tlHorizontal` (1): Vista horizontal tipo *stepper*, flujo de proceso o asistente *wizard*.
  * `tlVerticalAlternating` (2): Disposición moderna en zig-zag / split timeline con tarjetas a izquierda y derecha.
* **Formas Geométricas Configurables en Nodos (`NodeShape`):**
  * Círculo (`tlShapeCircle`)
  * Rectángulo Redondeado (`tlShapeRoundedRect`)
  * Cuadrado (`tlShapeSquare`)
  * Píldora / Cápsula (`tlShapePill`)
  * Diamante / Rombo (`tlShapeDiamond`)
  * **Triángulo Direccional Hacia Arriba** (`tlShapeTriangleUp`)
  * **Triángulo Direccional Hacia Abajo** (`tlShapeTriangleDown`)
* **Soporte de Imágenes y Avatares:** Carga directa de imágenes (PNG con transparencia alfa, JPG, BMP o `StdPicture`) con recorte circular automático (*Avatar Clip*).
* **Fuentes de Iconos Vectoriales:** Soporte nativo para glifos Unicode/Hexadecimales (ej. *Segoe MDL2 Assets*, *Font Awesome*, *Material Icons*).
* **Tarjetas de Información (Cards) Modernas:**
  * Sombras suaves multicapa difuminadas (*Soft Drop Shadows*).
  * Flecha puntero de conexión (*Callout Arrow*).
  * Insignias o etiquetas de estado en forma de píldora (*Badges*).
  * Formato de Título y Subtítulo multilineal.
* **Temas Visuales Integrados (`tlTheme`):**
  * `tlThemeLight`: Estilo moderno claro y minimalista.
  * `tlThemeDark`: Modo oscuro profesional de alto contraste.
  * `tlThemeCyberpunk`: Estilo neón futurista de alto impacto visual.
  * `tlThemeCorporate`: Azul corporativo ejecutivo.
  * `tlThemeCustom`: Control total de colores individuales.
* **Barra de Desplazamiento Fluida (`ucScrollbar`):** Control integrado de scroll con renderizado suave, arrastre continuo y cambio de color dinámico según el tema.
* **Arquitectura Orientada a Objetos 100% Segura en VB6:** Clases `clsTimePoint` y `clsTimePoints` sin dependencias de UDTs públicas ni memory thunks propensos a cuelgues del IDE.

---

## 📁 Estructura del Proyecto

| Archivo | Descripción |
| :--- | :--- |
| `AxTimeLineIC.ctl` | Control de Usuario ActiveX principal (Canvas GDI+, renderizado y eventos). |
| `clsTimePoint.cls` | Clase de datos de cada hito individual (Propiedades, colores, estado, icono e imagen). |
| `clsTimePoints.cls`| Colección tipada de hitos (`Add`, `Insert`, `Remove`, `Item`, `Count`, `Clear`). |
| `ucScrollbar.ctl`   | Control interno de barra de desplazamiento GDI+ con diseño moderno. |
| `mGDIPlus.bas`      | Módulo de soporte de APIs Win32 / GDI+, trazado de geometrías y sombras. |
| Projecto de prueba |
| `mdlTestData.bas`   | Generador de datos de prueba y escenarios de ejemplo. |
| `FormDemo.frm`      | Formulario de demostración interactiva (*Showcase*). |
| `AxTimeLineIC.vbp`  | Proyecto OCX / Control ActiveX. |
| `TestDemo.vbp`      | Proyecto ejecutable de prueba. |
| `Desarrollo.vbg`    | Grupo de proyectos de Visual Basic para depuración en vivo. |

---

## 🛠️ Guía de Uso y Ejemplos de Código

### 1. Inicialización y Carga de Hitos

```vb
' Limpiar hitos previos
AxTimeLine1.Points.Clear

' 1. Hito Completado con Icono
Dim pt1 As clsTimePoint
Set pt1 = AxTimeLine1.AddTimePoint("Requerimientos y Alcance", _
                         "Especificaciones aprobadas por el comite directivo.", _
                         "E73E", "09:00", "01/10/2026", True, , tlStatusCompleted)
pt1.BadgeText = "Completado"
pt1.BadgeColor = &H43A047

' 2. Hito En Progreso (Activo)
Dim pt2 As clsTimePoint
Set pt2 = AxTimeLine1.AddTimePoint("Desarrollo de AxTimeLineIC v2", _
                         "Implementacion del motor GDI+, sombras y figuras.", _
                         "E7BE", "16:45", "08/10/2026", True, , tlStatusInProgress)
pt2.BadgeText = "En Progreso"
pt2.BadgeColor = &HE65100

' 3. Hito con Forma de Triangulo Direccional
Dim pt3 As clsTimePoint
Set pt3 = AxTimeLine1.AddTimePoint("Validacion QA & Seguridad", _
                         "Pruebas de estres y renderizado sin fugas de memoria.", _
                         "E814", "10:00", "10/10/2026", True, , tlStatusPending)
pt3.NodeShape = tlShapeTriangleUp
pt3.BadgeText = "Importante"
pt3.BadgeColor = &H9C27B0

' 4. Hito con Imagen / Avatar
Dim pt4 As clsTimePoint
Set pt4 = AxTimeLine1.AddTimePoint("Despliegue a Produccion", _
                         "Lanzamiento de la nueva version del componente.", _
                         "", "15:00", "15/10/2026", True, LoadPicture(App.Path & "\avatar.jpg"), tlStatusPending)
pt4.ImageCircular = True

' Establecer seccion activa y refrescar
AxTimeLine1.ActiveSection = 2
AxTimeLine1.Refresh
```

---

### 2. Cambio de Temas y Orientación

```vb
' Cambiar Tema Visual
AxTimeLine1.Theme = tlThemeDark        ' Modo Oscuro
' AxTimeLine1.Theme = tlThemeLight     ' Modo Claro
' AxTimeLine1.Theme = tlThemeCyberpunk ' Modo Neon
' AxTimeLine1.Theme = tlThemeCorporate ' Azul Corporativo

' Cambiar Orientacion
AxTimeLine1.Orientation = tlVertical             ' Vertical clasico
' AxTimeLine1.Orientation = tlHorizontal         ' Horizontal (Stepper)
' AxTimeLine1.Orientation = tlVerticalAlternating' Zig-Zag / Split

' Cambiar Forma Global de los Nodos
AxTimeLine1.NodeShape = tlShapeCircle         ' Circulos
' AxTimeLine1.NodeShape = tlShapeRoundedRect  ' Rectangulos redondeados
' AxTimeLine1.NodeShape = tlShapeDiamond      ' Rombos
' AxTimeLine1.NodeShape = tlShapeTriangleUp   ' Triangulos apuntando arriba
```

---

### 3. Captura de Eventos

```vb
Private Sub AxTimeLine1_ItemClick(ByVal Index As Long, ByVal Point As clsTimePoint)
    MsgBox "Clic en el Hito #" & Index & ": " & Point.Title & vbCrLf & _
           "Estado: " & Point.BadgeText, vbInformation, "Detalle de Seleccion"
End Sub

Private Sub AxTimeLine1_ItemDblClick(ByVal Index As Long, ByVal Point As clsTimePoint)
    MsgBox "Doble clic en: " & Point.Title & vbCrLf & Point.Subtitle, vbInformation, "Edicion de Hito"
End Sub

Private Sub AxTimeLine1_ItemHover(ByVal Index As Long, ByVal Point As clsTimePoint)
    ' Feedback visual en barra de estado
    lblStatus.Caption = "Apuntando a: " & Point.Title
End Sub
```

---

## 📖 Referencia de la API

### Enumeraciones Públicas

#### `tlOrientation`
* `tlVertical = 0`: Disposición vertical estándar con columna de fecha.
* `tlHorizontal = 1`: Disposición horizontal tipo asistente / barra de pasos.
* `tlVerticalAlternating = 2`: Disposición vertical alternada en zig-zag / split timeline.

#### `tlNodeShape`
* `tlShapeDefault = -1`: Hereda la forma global del control.
* `tlShapeCircle = 0`: Círculo perfecto.
* `tlShapeRoundedRect = 1`: Rectángulo con curvatura configurable (`CornerCurve`).
* `tlShapeSquare = 2`: Cuadrado plano.
* `tlShapePill = 3`: Cápsula / Píldora horizontal.
* `tlShapeDiamond = 4`: Diamante / Rombo.
* `tlShapeTriangleUp = 5`: Triángulo apuntando hacia arriba.
* `tlShapeTriangleDown = 6`: Triángulo apuntando hacia abajo.

#### `tlItemStatus`
* `tlStatusNone = 0`: Sin estado especial.
* `tlStatusCompleted = 1`: Completado (Icono verde / tilde automático `✓`).
* `tlStatusInProgress = 2`: En curso (Borde naranja con brillo activo).
* `tlStatusPending = 3`: Pendiente (Colores atenuados).
* `tlStatusWarning = 4`: Advertencia o revisión requerida (Amarillo / Turquesa).
* `tlStatusFailed = 5`: Error o fallo (Rojo intenso).

#### `tlTheme`
* `tlThemeCustom = 0`: Colores personalizados por propiedad.
* `tlThemeLight = 1`: Tema claro moderno.
* `tlThemeDark = 2`: Tema oscuro.
* `tlThemeCyberpunk = 3`: Tema neón futurista.
* `tlThemeCorporate = 4`: Tema azul corporativo.

---

### Propiedades Principales de `AxTimeLineIC`

| Propiedad | Tipo | Descripción |
| :--- | :--- | :--- |
| `Points` | `clsTimePoints` | Colección de hitos / elementos de la línea de tiempo. |
| `ActiveSection` | `Long` | Índice (1-based) del hito actualmente activo o seleccionado. |
| `Orientation` | `tlOrientation` | Orientación del control (`tlVertical`, `tlHorizontal`, `tlVerticalAlternating`). |
| `Theme` | `tlTheme` | Tema visual aplicado globalmente. |
| `NodeShape` | `tlNodeShape` | Forma geométrica por defecto para todos los nodos. |
| `NodeSize` | `Long` | Diámetro/tamaño en píxeles de cada nodo (default: `38`). |
| `SectionSpace` | `Long` | Distancia en píxeles entre centros de hitos consecutivos (default: `90`). |
| `CardVisible` | `Boolean` | Determina si se renderizan las tarjetas de información descriptivas. |
| `CardRadius` | `Long` | Radio de curvatura de las esquinas de las tarjetas (default: `8`). |
| `CardShadowVisible` | `Boolean` | Activa o desactiva la sombra difuminada (*Soft Drop Shadow*). |
| `CalloutPointer` | `Boolean` | Muestra u oculta la flecha triangular conectora entre tarjeta y nodo. |
| `TitleFont` | `StdFont` | Tipografía para el título principal de las tarjetas. |
| `SubtitleFont` | `StdFont` | Tipografía para el texto secundario / descripción. |
| `IconFont` | `StdFont` | Fuente tipográfica utilizada para iconos vectoriales. |
| `DateVisible` | `Boolean` | Muestra u oculta la columna de fecha en orientación vertical. |
| `TimeVisible` | `Boolean` | Muestra u oculta la hora en la columna de fecha. |

---

### Métodos de `AxTimeLineIC`

| Método | Descripción |
| :--- | :--- |
| `AddTimePoint(...)` | Crea y agrega un nuevo hito al final de la colección, retornando el objeto `clsTimePoint`. |
| `UpdateTimePoint(...)` | Actualiza los datos de un hito existente según su índice. |
| `EnsureVisible(Index)` | Realiza un desplazamiento automático del scroll para asegurar que el hito indicado sea visible en pantalla. |
| `ApplyTheme(Theme)` | Aplica instantáneamente una paleta de colores preconfigurada a todos los elementos y al scrollbar. |
| `Refresh()` | Recalcula las dimensiones del scroll y repinta el control completo. |

---

### Propiedades de Cada Hito (`clsTimePoint`)

| Propiedad | Tipo | Descripción |
| :--- | :--- | :--- |
| `Title` | `String` | Texto del título principal. |
| `Subtitle` | `String` | Descripción o texto secundario multilineal. |
| `Timestamp` | `String` | Hora asociada al hito (ej. `"14:30"`). |
| `DateText` | `String` | Fecha asociada al hito (ej. `"10/10/2026"`). |
| `IconChar` | `Variant` | Carácter o código hexadecimal del icono (ej. `"E73E"` o `ChrW$(&HE73E)`). |
| `Picture` | `Variant` | Imagen asignada (ruta de archivo, `StdPicture` o handle `HBITMAP`). |
| `ImageCircular` | `Boolean` | Si es `True`, recorta la imagen como un avatar circular con anti-aliasing. |
| `NodeShape` | `tlNodeShape` | Permite anular la forma global para asignar una forma específica a este nodo. |
| `Status` | `tlItemStatus` | Estado del hito (`Completed`, `InProgress`, `Pending`, `Warning`, `Failed`). |
| `BadgeText` | `String` | Texto de la etiqueta de estado de la tarjeta (ej. `"Completado"`, `"Urgente"`). |
| `BadgeColor` | `OLE_COLOR` | Color de fondo de la etiqueta de estado. |
| `NodeColor` / `CardColor` | `OLE_COLOR` | Colores de fondo individuales (sobrescriben los del tema). |
| `Visible` | `Boolean` | Visibilidad del hito en la línea de tiempo. |
| `Tag` / `UserData` | `Variant` | Datos personalizados adjuntos al hito para lógica de negocio. |

---

## 💻 Requisitos del Sistema y Compatibilidad

* **Entorno de Desarrollo:** Microsoft Visual Basic 6.0 (SP6 recomendado).
* **Sistemas Operativos Compatibles:** Windows 7, Windows 8/8.1, Windows 10 y Windows 11 (32 bits y 64 bits).
* **Librerías del Sistema:** `gdiplus.dll` (incluida de forma nativa en todas las versiones modernas de Windows).
* **Codificación de Archivos:** Windows-1252 (ANSI) / UTF-8 con terminación de línea `CR+LF`.

---

## 📄 Licencia

Este componente forma parte de la suite de controles evolucionados para Visual Basic 6.0. Código abierto para integración en proyectos comerciales y educativos.
