# Pruebas: Calculadora de embalaje

**Versión probada:** v1.0 · **Fecha:** 17/10/2026, repetida el 18/10/2026 · **Quién probó:** Claude (asistente de IA), en el PC de casa de David. **Falta que la pruebe una persona de KIMAK** (apartado 4)

---

## 1. Casos probados

Probada en el navegador, sirviendo la aplicación en local. Los resultados se
comprobaron a mano con las fórmulas de `DOCUMENTACION_TECNICA.md`, con los
coeficientes de ejemplo (espesor 10 mm, merma 20 %, 7,50 €/m², 22 €/h, 15 min
fijos, 4 min/m², 3 € de accesorios). Los casos 1, 2 y 4 se repitieron el
18/10/2026 con el mismo resultado.

| # | Entrada | Resultado esperado | Resultado obtenido | ¿Pasa? |
|---|---|---|---|---|
| 1 | 1200 x 800 x 600 mm, 1 bulto | Exterior 1220x820x620; 4,53 m²; 5,44 m² con merma; 0,620 m³; 37 min; 40,77 € material; 13,47 € mano de obra; 57,25 € total | Idéntico | Sí |
| 2 | 2400 x 800 x 600 mm, 6 bultos | 7,99 m²; 9,58 m² con merma; 53 min; 94,43 € por caja; 57,50 m² totales; 7,382 m³; 5,3 h; 566,60 € total | Idéntico | Sí |
| 3 | Referencia escrita ("Obra 17321 Munich") | Aparece en el desglose y en la impresión | Correcto | Sí |
| 4 | Medida vacía o cero | Mensaje de error claro, sin resultado a medias | Correcto | Sí |

Diferencias de céntimo entre el total por caja y la suma de las partidas
mostradas: vienen de redondear a dos decimales solo al mostrar. El cálculo
interno no redondea.

## 2. Comparación con el método antiguo

**Pendiente.** David indica que quizá no exista un Excel de referencia y que el
cálculo se haga a ojo. Rubinos debe confirmar el método y comparar varios
bultos conocidos con la aplicación. Cualquier diferencia hay que explicarla
antes de dar la aplicación por buena.

## 3. Lo que NO se ha probado

- No se ha contrastado con el método real de Presupuestos (apartado 2).
- No se ha probado dentro de un contenedor: el `Dockerfile` está escrito pero no
  se ha construido.
- No se ha probado en el portal con Authentik, porque todavía no hay acceso.
- No se ha probado la impresión en papel, solo la vista de impresión.
- No se ha probado en móvil, aunque la pantalla está preparada para adaptarse.
- No se ha probado con un navegador que tenga bloqueado el almacenamiento local:
  el código lo contempla, pero no se ha comprobado en vivo.

## 3 bis. Cambio visual local del 19/10/2026

Se comprobó la estructura HTML, la existencia del logo local, la sintaxis del
script, el cálculo inicial y los eventos de apertura y cierre con un entorno
simulado. Falta comprobar el nuevo aspecto y el uso del menú en un navegador
tras recargar la página.

## 4. Usuarios de prueba

Pendiente de la lista de cinco personas que tiene que pasar Javier Rubinos.

| Persona | Departamento | Desde | Qué reportó |
|---|---|---|---|
| | | | |

## 5. Fallos encontrados y qué se hizo

| Fecha | Quién lo vio | Qué pasaba | Estado |
|---|---|---|---|
| | | | |
