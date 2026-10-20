# Ficha de la aplicación: Calculadora de embalaje

**Identificador interno:** MINI-13
**Fecha de la ficha:** 17/10/2026 · **Revisión local:** 19/10/2026

---

## 1. Identificación

| Dato | Valor |
|---|---|
| Nombre de la aplicación | Calculadora de embalaje |
| Nombre técnico (repositorio y contenedor) | `kimak-calculadora-embalaje` |
| Dirección prevista | La que asigne Sistemas dentro del portal |
| Departamento que la pide | Presupuestos (propuesta de Javier Rubinos en la reunión del 15/10/2026) |
| Fecha de puesta en laboratorio | Pendiente |
| Fecha de pase a producción | Pendiente |

## 2. Roles

| Papel | Persona | Qué le toca |
|---|---|---|
| Desarrollador responsable | David Vaquero | Código y documentación |
| Propietario funcional | Javier Rubinos (a confirmar) | Validar los coeficientes y el resultado; decidir cambios |
| Sistemas | Guillermo Paredes | Contenedor, permisos y despliegue |
| Responsable de auditoría | Pendiente | Revisar la documentación |

## 3. Quién puede entrar

| Grupo o persona | Nivel | Para qué |
|---|---|---|
| Usuarios de prueba del piloto (5 personas, lista de Rubinos) | Usuario | Calcular |
| Propietario funcional | Responsable | Cambiar los coeficientes de coste |

El apartado de coeficientes se abre desde el botón de ajustes de la barra
superior y pide una clave escrita en el propio archivo. **Eso no es seguridad
real**: solo evita cambios sin querer. Authentik controlaría la entrada al
portal, pero no distinguiría por sí solo quién puede editar estos valores.

## 4. Datos que utiliza

| Dato | De dónde sale | Lectura o escritura |
|---|---|---|
| Medidas del bulto y número de bultos | Los teclea el usuario | No se guardan |
| Referencia de pedido u obra (opcional) | La teclea el usuario | No se guarda |
| Coeficientes de coste | Los carga el responsable | Se guardan en el navegador de esa persona |

**Datos personales o de cliente:** el campo de referencia es libre y podría
contenerlos. No se envía al servidor, pero aparece en pantalla y puede quedar
en una impresión o PDF guardado por el usuario.

## 5. Integraciones

| Sistema | ¿Se integra? |
|---|---|
| SIG | No |
| Libra | No |
| MES | No |
| Internet | No |

## 6. Soporte

| Dato | Valor |
|---|---|
| Correo de soporte | Pendiente de que Sistemas lo cree en GLPI |
| Horario de atención | Horario de oficina |
| Qué hacer si no responde | Avisar al propietario funcional, que avisa a Sistemas |

## 7. Riesgos conocidos

| Riesgo | Qué puede pasar | Cómo se reduce |
|---|---|---|
| Coeficientes de ejemplo sin validar | Presupuestos mal calculados | La app avisa en pantalla hasta que se guardan ajustes; ese cambio no demuestra validación. Rubinos debe confirmar los valores |
| Coeficientes distintos en cada ordenador | Dos personas calculan distinto el mismo bulto | Se resuelve cuando Sistemas deje los valores fijos en la aplicación en vez de en cada navegador. Está anotado como mejora para la v1.1 |
| Fórmula de superficie simplificada | Se calcula como caja cerrada de seis caras; un embalaje con solapes o refuerzos consumirá algo más | Hay que contrastar el cálculo y la merma con Presupuestos |

## 8. Para el alta en el portal

- Nombre técnico: `kimak-calculadora-embalaje`
- Archivos que componen la app: `src/index.html` y `src/assets/kimak-logo-blanco.png`
- ¿Necesita guardar datos? No
- ¿Necesita salir a internet? No
- Grupos del dominio con acceso: pendiente de la lista de cinco usuarios de prueba
- Correo de soporte solicitado: `soporte-embalaje@` (propuesta)
