# Historial de versiones: Calculadora de embalaje

| Versión | Fecha | Qué cambió | Autor | Aprobado por | Dónde está |
|---|---|---|---|---|---|
| v1.1 | 20/10/2026 | Cálculo alineado con el Excel de embalajes (rev. oct. 2026): material MDF/aglomerado, espesores 16/19/30 con precio por m², mano de obra fija y variable, tacos, consumibles y generales; peso estimado; precios y tiempos editables en ajustes; botones independientes Imprimir, Guardar PDF y Descargar CSV (PDF con logo KIMAK y tablas, sin librerías; CSV con columnas concepto/valor/unidad; datos generales separados del desglose); resultado resumido con desglose desplegable; diseño renovado y adaptado a móvil. **Cambian los resultados** respecto a v1.0 | David Vaquero | - | Sin desplegar |
| v1.0 | 17/10/2026 | Primera versión: cálculo de tablero, merma, volumen, tiempo y coste; ajustes del responsable; impresión en PDF | David Vaquero | - | Sin desplegar |

## Cambios locales anteriores

- 19/10/2026: barra superior negra con logo KIMAK y ajustes accesibles desde su
  botón. No cambian las fórmulas ni los coeficientes por defecto.

## Pendiente antes de la v1.1

- Cargar los coeficientes validados por Presupuestos (cambia resultados: hay que
  anotarlo aquí cuando se haga).
- Fijar los coeficientes en el archivo para que no dependan del navegador.

## Cómo se rellena

- Sube el segundo número para cambios pequeños y el primero cuando cambie la
  forma de usarla o los resultados.
- Si cambia un cálculo o un coeficiente, dilo claramente: afecta a presupuestos
  ya emitidos.
- Cada fila que diga "producción" tiene que tener su etiqueta en GitLab.
