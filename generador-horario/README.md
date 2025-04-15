# Calculadora de Horas Laborales en Excel

Este proyecto es una **plantilla de Excel** que permite calcular de forma automática las horas trabajadas durante un mes, teniendo en cuenta el horario laboral habitual y los días festivos o de vacaciones.

---

## ¿Qué hace?

- Genera una hoja mensual con todas las jornadas laborales del mes seleccionado.
- Calcula automáticamente las **horas trabajadas por día**.
- Permite indicar **horarios diferentes por día de la semana (lunes a viernes)**.
- Admite **pausas para comer** por separado.
- Identifica los **días festivos o vacaciones** para excluirlos del cálculo.
- Muestra el **total de horas trabajadas al mes**.
- Si ya existe una hoja con el nombre del mes, la elimina para evitar duplicados.

---

## ¿Cómo funciona?

1. **Introduce el año y el número del mes** (por ejemplo: Año: `2025`, Mes: `1`).
2. **Define los horarios laborales** por día de la semana (entrada, salida y duración de la pausa).
3. **Indica los días festivos o de vacaciones** en la columna correspondiente (por ejemplo: `1`, `6`, `22`).
4. Pulsa el botón **"Generar mes"**.
5. Automáticamente se creará una hoja con el nombre del mes y todas las fechas laborales del mes.
6. El sistema identificará los días festivos y calculará las **horas reales trabajadas por jornada**.

---

## Ejemplo visual

### Configuración inicial:

![Configuración](https://github.com/tuusuario/tu-repo/blob/main/assets/configuracion.png)

### Resultado generado:

![Resultado](https://github.com/tuusuario/tu-repo/blob/main/assets/resultado.png)

---

## ⚠️ Notas importantes

- El archivo sobrescribirá cualquier hoja existente con el nombre del mes que se va a generar.
- Los días festivos deben indicarse con el número del día del mes, uno por celda.
- Funciona únicamente con macros habilitadas (asegúrate de habilitar contenido en Excel al abrir el archivo).

---

## 🛠 Requisitos

- Microsoft Excel con soporte para macros (VBA).
- Habilitar macros al abrir el archivo `.xlsm`.

---

## ✨ Créditos

Este proyecto ha sido creado por [Tu nombre o alias] como una solución simple para automatizar el cálculo de jornadas laborales en Excel.
¡Puedes modificarlo y adaptarlo a tus necesidades!

---

¿Quieres que te genere también los archivos `configuracion.png` y `resultado.png` si aún no los tienes? ¿O prefieres que los suba como parte del README solo con texto?