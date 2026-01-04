# Master-Habit-Tracker-VBA-Analytics
Tracker de Hábitos 2026 en Excel potenciado con VBA. Sistema automatizado con inserción dinámica de filas, gestión de Checkboxes mediante macros y Dashboard de indicadores (Engagement y Eficiencia) para el análisis de disciplina de alto rendimiento.
# 🚀 **Master Tracker de Hábitos de Alto Rendimiento 2026**
 ![image alt](https://github.com/ramirezdavidge/Master-Habit-Tracker-VBA-Analytics/blob/083897f90a4eae1a04e8883f211e979862d7b6bb/images/portada.PNG)
## 📌 **Descripción del Proyecto**
Este sistema es una solución integral de **Business Intelligence aplicada al desarrollo personal**.  
Desarrollado íntegramente en **Excel** y potenciado con **VBA (Visual Basic for Applications)**, este Tracker permite no solo **registrar hábitos**, sino también **analizar el comportamiento, la constancia y la disciplina del usuario** mediante indicadores dinámicos y automatizaciones avanzadas.

---
 ![image alt](https://github.com/ramirezdavidge/Master-Habit-Tracker-VBA-Analytics/blob/083897f90a4eae1a04e8883f211e979862d7b6bb/images/Mes_Enero.PNG)
 ![image alt]( https://github.com/ramirezdavidge/Master-Habit-Tracker-VBA-Analytics/blob/083897f90a4eae1a04e8883f211e979862d7b6bb/images/Mes_Setiembre.PNG)
## 📊 **Dashboard de Indicadores y Análisis**
El proyecto cuenta con un **panel de control superior** que transforma los datos diarios en **métricas accionables**.

### 🔑 **Métricas de Rendimiento (KPIs)**

- **🏆 Hábito con Mayor Engagement**  
  Identificado mediante la fórmula:  
=INDICE(B11:B25; COINCIDIR(MAX(AI11:AI25); AI11:AI25; 0))
Permite conocer el hábito con mayor nivel de cumplimiento.

- **📈 Promedio de Efectividad Diaria**  
Calcula el porcentaje global de éxito considerando todas las tareas del mes.

- **⚠️ Hábito en Desarrollo**  
Detecta el hábito con menor nivel de cumplimiento histórico para enfocar esfuerzos de mejora.

- **🚨 Cantidad de Hábitos (< 50% de realización)**  
Indicador crítico que utiliza `CONTAR.SI` para mostrar cuántos hábitos se encuentran en la *zona de riesgo*.

- **📊 Análisis Gráfico**  
Gráfico de barras dinámico que muestra los **Días de Cumplimiento por cada Hábito**, facilitando la comparación visual entre metas.

---
![image alt](https://github.com/ramirezdavidge/Master-Habit-Tracker-VBA-Analytics/blob/083897f90a4eae1a04e8883f211e979862d7b6bb/images/indicadores.PNG)
## 📅 **Seguimiento Mensual y Estructura de Datos**
El archivo está organizado en **hojas individuales para cada mes (Enero a Diciembre)**, permitiendo un seguimiento completo durante todo el año **2026**.

### 🧩 **Características Técnicas de la Hoja**

- **☑️ Interactividad con Checkboxes**  
Cada día del mes cuenta con casillas de verificación vinculadas a celdas ocultas que alimentan las fórmulas de análisis.

- **🔥 Cálculo de Racha Actual**  
Columna dedicada a contar los días consecutivos de cumplimiento exitoso.

- **📉 Visualización de Progreso**  
Uso de **barras de datos (Formato Condicional / Sparklines)** en la columna *“% Realización Fin del Mes”* para lectura rápida del avance.

- **📆 Control de Carga Diaria**  
Al final de cada columna se muestra el **% de tareas completadas por día**, permitiendo identificar los días más productivos de la semana.

---

## 💻 **Arquitectura del Código (VBA)**
El backend del proyecto está diseñado con un **enfoque modular**, optimizado para manejar grandes cantidades de **Checkboxes** sin afectar el rendimiento.

### 🧱 **Estructura del VBAProject**
![image alt](https://github.com/ramirezdavidge/Master-Habit-Tracker-VBA-Analytics/blob/083897f90a4eae1a04e8883f211e979862d7b6bb/images/Estructura_VBA.PNG)
- **📋 Formularios (UserForm1)**  
Interfaz gráfica para la captura limpia y controlada de nuevos hábitos.

- **📦 Módulos (1–4)**  
Contienen la lógica de negocio para la **inserción**, **eliminación** y **limpieza** de datos.

---

## ⚙️ **Funciones Principales de Automatización**

### ➕ Inserción Dinámica de Hábitos
- Captura el nombre del hábito mediante un formulario.
- La macro identifica la última fila ocupada en la columna **B**.
- Inserta una nueva fila y **desplaza automáticamente los totales**, evitando romper fórmulas existentes.

### ☑️ Gestión de Checkboxes
- La macro `Checkboxes_por_Habitos` recorre las columnas de la **C a la AG**.
- Genera automáticamente casillas solo para los **días válidos del mes**.
![image alt]( https://github.com/ramirezdavidge/Master-Habit-Tracker-VBA-Analytics/blob/083897f90a4eae1a04e8883f211e979862d7b6bb/images/checkboxes_habitos.PNG)

### 🧹 Limpieza y Mantenimiento
- **Eliminación Inteligente:**  
La macro `EliminarFilasYCasillasVacias` borra filas innecesarias y Checkboxes huérfanos para mantener el archivo ligero.
![image alt](https://github.com/ramirezdavidge/Master-Habit-Tracker-VBA-Analytics/blob/083897f90a4eae1a04e8883f211e979862d7b6bb/images/eliminar_vacias.PNG )
- **Reinicio Total:**  
La función `LimpiarRangoHabitos` resetea el tablero eliminando formatos invisibles (`;;;`) y contenidos de las filas **11 a 30**.
![image alt]( https://github.com/ramirezdavidge/Master-Habit-Tracker-VBA-Analytics/blob/083897f90a4eae1a04e8883f211e979862d7b6bb/images/reiniciar_total.PNG)
- **Llamada al formulario:**
Esta se encarga de llamar al formulario para poder registrar un nuevo Habito. 
![image alt](https://github.com/ramirezdavidge/Master-Habit-Tracker-VBA-Analytics/blob/083897f90a4eae1a04e8883f211e979862d7b6bb/images/abrir_formulario.PNG )
---

## 📖 **Guía de Usuario e Instrucciones**

**Paso 1️⃣**  
Presionar el botón **“Insertar un nuevo Hábito”** para abrir el formulario.

**Paso 2️⃣**  
Registrar el hábito. El sistema lo insertará al final de la lista y ajustará los totales automáticamente.

**Paso 3️⃣**  
Si se utilizan menos de 20 hábitos, presionar **“Eliminar fila de hábitos vacío”** para optimizar el espacio del tablero.
![image alt](https://github.com/ramirezdavidge/Master-Habit-Tracker-VBA-Analytics/blob/083897f90a4eae1a04e8883f211e979862d7b6bb/images/instrucciones.PNG)
