# Documentación técnica: Calculadora de embalaje

**Versión:** v1.1 · **Fecha:** 20/10/2026

---

## 1. Qué hace, en términos técnicos

Lee tres medidas y una cantidad de un formulario, aplica las fórmulas del
apartado 4 con los coeficientes configurables, y pinta el desglose en pantalla.
No hay servidor, no hay base de datos y no hay peticiones de red.

## 2. De qué está hecha

| Pieza | Qué se usa | Por qué |
|---|---|---|
| Interfaz | HTML y CSS dentro del propio archivo | No necesita instalación ni descargas |
| Cálculos | JavaScript en el mismo archivo | Se ejecuta en el navegador del usuario |
| Almacenamiento | `localStorage` del navegador, solo para los coeficientes | No hace falta base de datos para una calculadora |
| Librerías externas | Ninguna | Menos superficie de ataque, no depende de internet y Sistemas puede revisar un solo archivo |

Toda lectura y escritura de `localStorage` va dentro de `try/catch`: si el
navegador lo tiene bloqueado, la aplicación sigue funcionando con los valores de
ejemplo en lugar de romperse.

## 3. Archivos

```
src/
├── index.html      Estructura, estilos y cálculos
└── assets/
    └── kimak-logo-blanco.png  Logo local de KIMAK para la barra superior
```

## 4. Las fórmulas

Origen: hoja EMBALAJES del Excel `embalajes_rev_oct_2026.xlsx` (Rubinos).
Medidas en mm: ancho `A`, fondo `F`, alto `H`. Si en ajustes las medidas son
"interiores", cada una crece 2 x espesor; por defecto se usan tal cual (Excel).

| Qué se calcula | Cómo (celda del Excel) |
|---|---|
| Precio del m² | `precio_tablero(material, espesor) x (1 + merma/100) + recargo` (R20:R22; 20 % y +1 €) |
| Superficie por pares de piezas | `2·A·H`, `2·F·H`, `2·F·A` (filas 31-33), en m² |
| Coste de material | `m² x precio del m²` por cada par |
| Volumen (m³) | `A·F·H / 1e9` (B18) |
| Mano de obra fija | `minutos_fijos/60 x €/h` (H34: 30 min) |
| Mano de obra variable | `m³ x horas_por_m³ x €/h` (H35: 0,45 h/m³) |
| Tacos inferiores | `precio_taco x tacos_por_tramo x ROUNDUP(A / tramo)` (H36: 0,75 x 2 x ROUNDUP(A/800)) |
| Consumibles | importe fijo por caja (H37: 2 €) |
| Gastos generales | `% x subtotal` (H38: 0 %) |
| Total por caja | suma de todo lo anterior (H39) |
| Peso estimado (añadido, no está en el Excel) | `m² x espesor(m) x densidad` del material, sin merma |

Comprobación de referencia: 2000 x 1200 x 500, aglomerado 19 mm, 1 bulto =
**125,12 €** (coincide con el Excel).

### Coeficientes configurables

Están en `AJUSTES_POR_DEFECTO` (inicio del JavaScript) y se editan desde el
botón de ajustes: tabla de precios por material y espesor, merma, recargo por
m², densidades (750 kg/m³ MDF y 650 aglomerado son **estimaciones mías**, no del
Excel), €/hora, minutos fijos, horas por m³, tacos, consumibles, generales y
tipo de medidas. Se guardan en `localStorage` (clave `kimak_embalaje_ajustes_v2`;
los ajustes de la v1.0 se ignoran porque tenían otra estructura).

## 5. Decisiones tomadas y por qué

| Decisión | Alternativa descartada | Motivo |
|---|---|---|
| Un solo archivo HTML | Framework con varios archivos y compilación | Guillermo dijo que para él la app son ficheros que carga en un contenedor. Un archivo se revisa, se despliega y se entiende sin herramientas |
| Sin librerías externas | Cargar estilos o librerías de un CDN | La regla del portal es no depender de internet, y así no hay componentes de terceros que auditar |
| Coeficientes en el navegador | Base de datos | Una calculadora no justifica una base de datos. Cuando haga falta que todos vean los mismos valores, se fijan en el archivo |
| Clave escrita en el archivo para los ajustes | Sistema de usuarios propio | El control de acceso real lo da Authentik. Duplicarlo aquí sería justo lo que IT no quiere |

## 6. Límites conocidos

- Calcula una **caja cerrada de seis caras**. No contempla jaulas, bases de
  palet, embalajes abiertos ni refuerzos interiores.
- No contempla solapes ni despiece real del tablero: la merma configurable es la
  que absorbe esa diferencia.
- Los coeficientes son por navegador, así que dos personas pueden calcular
  distinto hasta que Sistemas los fije en el archivo.
- No guarda historial de cálculos.

## 7. Cómo se construye y se despliega

```bash
docker build -t kimak-calculadora-embalaje .
docker run -p 8080:80 kimak-calculadora-embalaje
# abrir http://localhost:8080
```

Para probarla sin Docker basta con abrir `src/index.html` en el navegador.

## 8. Qué habría que hacer en la próxima versión

1. Contrastar fórmulas y coeficientes con el criterio real de Presupuestos.
2. Fijar los coeficientes en el archivo, para que todos calculen igual.
3. Añadir los tipos de embalaje que falten (jaula, base de palet).
4. Si Presupuestos lo pide, exportar el resultado en una línea para pegar en su
   hoja de cálculo.
