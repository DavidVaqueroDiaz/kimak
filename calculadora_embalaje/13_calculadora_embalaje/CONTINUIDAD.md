# Continuidad: Calculadora de embalaje

**Fecha:** 17/10/2026 · **Autor:** David Vaquero

---

## 1. Si mañana no estoy, lo primero que hay que saber

Es una calculadora de coste de embalaje. La usa Presupuestos para valorar lo que
cuesta la caja de madera de un bulto. Si se apaga, Presupuestos vuelve a su
método habitual de valoración. La aplicación no está en producción.

**Criticidad: baja.** Es la aplicación más sencilla del portal, elegida a
propósito como piloto para probar el circuito completo.

## 2. Quién sabe de esto, aparte de mí

| Persona | Qué sabe |
|---|---|
| Javier Rubinos | Debe confirmar el método y los coeficientes. Es quien propuso la app |
| Presupuestos | Usa su método habitual de valoración |

## 3. Las reglas de negocio que hay detrás

- Las medidas que se meten son las **interiores**: lo que necesita el mueble por
  dentro. La caja por fuera crece dos espesores de tablero en cada medida.
- La **merma** cubre recortes y desperdicio. El valor de ejemplo es un 20 %,
  que es el criterio que Rubinos mencionó para las creces de material en la
  reunión del 15/10. Hay que confirmarlo con Rubinos.
- El tiempo de montaje se reparte en una parte fija por caja y otra según los
  metros de tablero.
- Quien puede cambiar esos coeficientes es el propietario funcional, no el
  desarrollador.

## 4. Lo que hay que mirar antes de tocar nada

- Todo el cálculo está en la función `calcular()` de `src/index.html`. Está
  comentada y es corta.
- Los coeficientes están en `AJUSTES_POR_DEFECTO`, justo encima.
- Cualquier cambio en una fórmula o en un coeficiente **cambia presupuestos que
  ya se han dado**. Eso se anota en el `CHANGELOG.md` con fecha, porque alguien
  puede tener que explicar por qué el mismo bulto costaba otra cosa el mes
  pasado.

## 5. Cómo probar que sigue funcionando

Con los coeficientes de ejemplo (botón "Volver a los valores de ejemplo"):

1. 1200 x 800 x 600 mm, 1 bulto, tiene que dar **5,44 m²** con merma y
   **57,25 €**.
2. 2400 x 800 x 600 mm, 6 bultos, tiene que dar **94,43 €** por caja y
   **566,60 €** de total.

Si alguno cambia sin que se haya tocado a propósito una fórmula, el cambio está
mal aunque la aplicación no dé ningún error.

## 6. Dónde está todo

| Qué | Dónde |
|---|---|
| Código y documentación | GitLab: `kimak-calculadora-embalaje` (pendiente de crear) |
| Origen de la idea | Acta de la reunión del 15/10/2026 |
| Método de partida | Pendiente de confirmar con Presupuestos; no consta un Excel de referencia |

## 7. Lo que dejaría a medias

- Falta contrastar las fórmulas y los coeficientes con Presupuestos.
- Falta fijar los coeficientes para que no dependan del navegador de cada uno.
- Faltan los tipos de embalaje que no son caja cerrada.
