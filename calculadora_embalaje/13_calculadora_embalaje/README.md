# Calculadora de embalaje

**Qué es:** metes las tres medidas interiores de un bulto y te devuelve los
metros cuadrados de tablero, el tiempo de montaje y el coste del embalaje.

**Para quién:** Presupuestos, y más adelante la persona de embalaje que tenga
que valorar el coste de cada bulto.

**Estado:** prototipo local en revisión; pendiente de validar los coeficientes y el proceso de alta.
**Versión actual:** v1.1 · **Última revisión local:** 20/10/2026

---

## Para qué sirve

Presupuestos valora el embalaje con su criterio de trabajo. Esta aplicación
calcula metros de tablero, minutos según la superficie calculada y un coste
orientativo. Desde la v1.1 reproduce el Excel de embalajes de Rubinos (rev. oct. 2026).

Se propone como **aplicación piloto del portal** por lo que se habló en la
reunión del 15/10: es la más sencilla posible, no usa IA, no se integra con nada
y sirve para recorrer todo el circuito (contenedor, acceso por Authentik,
permisos, usuarios de prueba) antes de meter aplicaciones serias.

## Cómo se usa

Ver [`MANUAL_USUARIO.md`](MANUAL_USUARIO.md). Resumen: tres medidas, número de
bultos, botón Calcular, y botón de imprimir o guardar en PDF. Los ajustes se
abren desde el botón de la barra superior.

## Quién responde de ella

Ver [`FICHA_APP.md`](FICHA_APP.md). Resumen:

| Papel | Persona |
|---|---|
| Desarrollador responsable | David Vaquero |
| Propietario funcional | Javier Rubinos (pendiente de confirmar) |
| Sistemas | Guillermo Paredes |

## Soporte

Pendiente del correo que cree Sistemas en GLPI. Mientras tanto, avisar al
propietario funcional. Por Teams no, que se pierde el rastro.

## Qué NO hace esta aplicación

- No se integra con SIG, Libra ni MES.
- No llama a internet: no carga ninguna librería ni tipografía de fuera.
- No guarda nada en ningún servidor. Los coeficientes que cambia el responsable
  se quedan en su propio navegador.
- No decide el precio de venta: calcula el coste con los coeficientes que le
  carguen.

## Cómo está hecha

Una página en `src/index.html`, con estructura, estilos y cálculos dentro, más
el logo local en `src/assets/`. Sin librerías ni llamadas a internet. Detalle en
[`DOCUMENTACION_TECNICA.md`](DOCUMENTACION_TECNICA.md).

## Importante antes de usarla de verdad

**Los coeficientes cargados vienen del Excel de Rubinos** (las densidades para el peso son estimadas).
Rubinos debe confirmar los valores definitivos y comprobar resultados conocidos
antes de usarla para presupuestar. El aviso puede desaparecer al guardar
ajustes, pero eso no demuestra que los valores hayan sido validados.

## Documentos de este repositorio

| Documento | Para qué |
|---|---|
| `README.md` | Este resumen |
| `FICHA_APP.md` | Roles, alta de usuarios y soporte |
| `MANUAL_USUARIO.md` | Cómo se usa |
| `DOCUMENTACION_TECNICA.md` | Cómo funciona por dentro y las fórmulas |
| `CONTINUIDAD.md` | Cómo seguir con esto sin el autor |
| `CHANGELOG.md` | Historial de versiones |
| `PRUEBAS.md` | Qué se probó y qué no |
