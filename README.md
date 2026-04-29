# Horas Extra v2

Automatización en Google Apps Script para consolidar, validar y desglosar horas extra mensuales a partir de datos operativos, nómina, novedades, feriados y valores hora.

El proyecto procesa legajos, calcula horas teóricas y reales, ajusta topes mensuales, clasifica horas pagables y genera una salida final tanto en hoja resumen como en hoja operativa con desglose por tipo de hora.

## Tabla de contenidos

- [Objetivo](#objetivo)
- [Alcance funcional](#alcance-funcional)
- [Estructura del repositorio](#estructura-del-repositorio)
- [Arquitectura general](#arquitectura-general)
- [Requisitos](#requisitos)
- [Configuración esperada en Google Sheets](#configuración-esperada-en-google-sheets)
- [Flujo de ejecución](#flujo-de-ejecución)
- [Reglas de negocio principales](#reglas-de-negocio-principales)
- [Configuración funcional](#configuración-funcional)
- [Permisos y seguridad](#permisos-y-seguridad)
- [Instalación y actualización](#instalación-y-actualización)
- [Uso](#uso)
- [Entradas y salidas](#entradas-y-salidas)
- [Mantenimiento](#mantenimiento)
- [Pruebas y validación](#pruebas-y-validación)
- [Troubleshooting](#troubleshooting)
- [Roadmap sugerido](#roadmap-sugerido)
- [Contribución](#contribución)
- [Autores y mantenimiento](#autores-y-mantenimiento)
- [Licencia](#licencia)

## Objetivo

Este repositorio centraliza la lógica de cálculo de horas extra para una operatoria mensual basada en Google Sheets. Su propósito es:

- tomar datos operativos por legajo;
- cruzarlos con nómina y novedades;
- aplicar reglas por tipo de servicio;
- considerar feriados y ausencias;
- calcular horas teóricas, horas reales y control operativo;
- desglosar las horas pagables en normales, 50% y 100%;
- escribir un resumen final y una hoja operativa desglosada.

## Alcance funcional

El script hoy contempla, entre otras, las siguientes capacidades:

- Lectura de configuración desde la hoja `DATOS SCRIPT`.
- Lectura de valores hora desde la hoja `VALORES HORAS`.
- Lectura de datos de nómina desde un archivo externo de Google Sheets.
- Resolución de tipo de servicio a partir de un catálogo externo.
- Tratamiento de feriados por tipo de servicio y por servicio puntual.
- Tratamiento de ausencias en base a la hoja de novedades.
- Cálculo de topes mensuales a partir de una base de 192 horas.
- Desglose por categoría de hora para la hoja operativa.
- Control de permisos para ejecución manual.

## Estructura del repositorio

```text
.
├── desglose.gs   # Reglas de clasificación y reparto de horas por tipo
├── horas.gs      # Orquestación principal, lectura de datos y cálculo mensual
└── README.md     # Documentación del proyecto
```

## Arquitectura general

### `horas.gs`

Archivo principal del proceso. Contiene:

- configuración global del script;
- menú de Google Sheets;
- validación de configuración;
- carga de contexto y datos base;
- cálculo de horas teóricas, reales y topes;
- escritura de resultados;
- soporte para feriados, servicios y nómina.

### `desglose.gs`

Contiene la lógica de negocio que clasifica y distribuye horas extra en:

- horas normales;
- horas al 50%;
- horas al 100%.

También contempla reglas especiales por servicio y por evento.

## Requisitos

- Una planilla de Google Sheets que actúe como archivo principal del proceso.
- Un proyecto de Google Apps Script vinculado a esa planilla.
- Acceso al archivo externo de nómina configurado en `DATOS SCRIPT`.
- Acceso al catálogo de servicios utilizado para resolver el tipo de servicio.
- Permisos suficientes para ejecutar el script manualmente.

## Configuración esperada en Google Sheets

El proceso asume la existencia de estas hojas:

- `DATOS SCRIPT`
- `VALORES HORAS`
- `Novedades <Mes> <AA>` o `Novedades`
- `Horas Extra <Mes> <AA>` o `Horas Extra`
- `Resultados <Mes> <AA>` o `Resultados`

### Hoja `DATOS SCRIPT`

Debe contener, como mínimo, la configuración base del proceso:

- URL nómina
- Hoja nómina
- Mes
- Año
- Desde fila
- Hasta fila
- Última ejecución
- Estado
- Mensaje

Además, desde esta hoja también se leen:

- feriados por tipo de servicio;
- feriados por servicio puntual;
- cierres extraordinarios por servicio.

### Hoja `VALORES HORAS`

Debe proveer, por categoría:

- valor hora normal;
- valor hora extra al 50%;
- valor hora extra al 100%.

### Archivo externo de nómina

La hoja configurada debe permitir identificar, como mínimo:

- legajo;
- apellido;
- nombre;
- jornada;
- categoría;
- servicio.

## Flujo de ejecución

1. El usuario abre la planilla y el script agrega el menú `Horas Extra`.
2. Se valida la configuración y la existencia de hojas necesarias.
3. Se cargan datos operativos, novedades, nómina, valores y horas reales.
4. Se resuelve el tipo de servicio de cada legajo a partir del catálogo externo.
5. Se calculan horas teóricas base y finales.
6. Se calcula el tope mensual.
7. Se determina el control operativo (`min(hs solicitadas, hs reales - hs teóricas)`).
8. Se desglosan las horas pagables.
9. Se escriben resultados en la hoja de resultados y en la hoja operativa.
10. Se registra estado, mensaje y fecha de última ejecución.

## Reglas de negocio principales

### Tope mensual

El script parte de una base mensual de `192` horas y ajusta el tope según la reducción entre horas teóricas base y horas teóricas finales.

### Tipos de servicio

La lógica actual contempla formas de trabajo tales como:

- `Colegio`
- `Hospital`
- `Supermercado`
- `Lunes a Sábado`

La resolución del tipo depende del catálogo externo, con coincidencia exacta, simplificada o parcial.

### Feriados

Los feriados se pueden configurar de dos maneras:

- por tipo de servicio;
- por servicio puntual.

El resultado final combina ambas fuentes.

### Ausencias

Las novedades reducen los días considerados para el cálculo de horas teóricas finales.

### Jornadas

El cálculo de horas teóricas depende tanto del tipo de servicio como de la jornada semanal.

Caso relevante actualmente implementado para `Supermercado`:

- jornada `24`: se calcula con regla específica de `4` horas por día considerado;
- jornada `36`: se calcula con regla específica de `6` horas por día considerado;
- jornadas distintas de `44`: mantienen descuento de `8` francos;
- jornada `44`: mantiene descuento de `6` francos.

Las reglas de ausencias y feriados siguen vigentes sobre esos mismos días considerados.

### Desglose de horas

Las horas se distribuyen entre:

- normales;
- 50%;
- 100%.

La clasificación depende del servicio, la fecha, el día, el contexto del pedido y si el registro está marcado como evento.

## Configuración funcional

### Permisos de ejecución

El script restringe la ejecución manual a una lista de usuarios autorizados definida en código.

### Zona horaria

La zona horaria utilizada es la del propio proyecto de Apps Script. Si no está disponible, se usa `America/Argentina/Cordoba`.

### Catálogo de servicios

El proyecto consulta un archivo externo para mapear el nombre del servicio al tipo de servicio funcional que usa la lógica de cálculo.

## Permisos y seguridad

- La ejecución está controlada por una lista explícita de correos autorizados.
- El acceso a la nómina y al catálogo de servicios depende de permisos de Google Sheets.
- La información procesada puede contener datos personales y operativos, por lo que se recomienda limitar el acceso al proyecto y a la planilla.
- No exponer públicamente URLs, credenciales ni archivos con datos sensibles.

## Instalación y actualización

1. Abrir el proyecto de Google Apps Script vinculado a la planilla.
2. Crear o actualizar los archivos `horas.gs` y `desglose.gs` con el contenido del repositorio.
3. Guardar los cambios.
4. Recargar la planilla para regenerar el menú.

## Uso

1. Completar la hoja `DATOS SCRIPT`.
2. Verificar la hoja `VALORES HORAS`.
3. Confirmar que las hojas mensuales existan con el nombre esperado.
4. Ejecutar `Validar configuración` desde el menú `Horas Extra`.
5. Ejecutar `Generar resumen final`.
6. Revisar las hojas de salida y el estado final del proceso.

## Entradas y salidas

### Entradas

- hoja operativa mensual;
- hoja de novedades mensual;
- hoja de valores hora;
- hoja de configuración;
- hoja externa de nómina;
- catálogo externo de servicios.

### Salidas

#### Hoja de resultados

Genera un resumen por legajo con columnas como:

- legajo;
- nombre;
- apellido;
- tope;
- horas solicitadas;
- horas teóricas;
- horas reales;
- control operativo;
- diferencia;
- horas al 100%;
- horas al 50%;
- horas normales.

#### Hoja operativa

Reescribe la hoja operativa con desglose por registro, incluyendo:

- tipo de hora asignada;
- tope aplicado;
- horas asignadas;
- horas denegadas o parcialmente denegadas.

Además aplica color de fondo según:

- novedad;
- denegación;
- evento.

## Mantenimiento

Se recomienda mantener actualizados los siguientes puntos cada vez que cambie la operatoria:

- catálogo de servicios;
- reglas de jornada;
- valores hora por categoría;
- estructura de la nómina externa;
- nombres de hojas;
- reglas de feriados y cierres por servicio.

## Pruebas y validación

Antes de usar cambios en producción, validar al menos:

- legajos con y sin novedades;
- servicios con y sin feriados configurados;
- jornadas estándar y jornadas especiales;
- casos con horas reales menores a las teóricas;
- registros con evento;
- servicios con reglas particulares de desglose.

Se recomienda conservar una planilla de prueba con casos controlados para regresión funcional.

## Troubleshooting

### El menú no aparece

- Recargar la planilla.
- Verificar que el proyecto esté vinculado correctamente.
- Confirmar que no haya errores de guardado en Apps Script.

### Falla la ejecución por permisos

- Verificar que el usuario esté en la lista de ejecutores autorizados.
- Confirmar acceso al archivo externo de nómina y al catálogo de servicios.

### No encuentra hojas

- Revisar los nombres configurados en `DATOS SCRIPT`.
- Confirmar que existan las variantes mensuales o genéricas esperadas.

### Resultados inesperados en horas teóricas

- Revisar jornada del legajo en nómina.
- Revisar tipo de servicio resuelto por catálogo.
- Revisar feriados configurados y novedades del mes.

## Roadmap sugerido

- Versionar reglas funcionales por período o convenio.
- Incorporar logging más detallado para auditoría operativa.

## Contribución

Para mantener trazabilidad y reducir regresiones:

- trabajar con cambios pequeños y revisables;
- documentar toda modificación de regla de negocio;
- validar escenarios reales antes de publicar;
- actualizar este README cuando cambie el comportamiento funcional.

## Autores y mantenimiento

### Autoría funcional

Este proyecto refleja reglas operativas del proceso de horas extra de Kazaro y debe mantenerse en coordinación con las áreas dueñas del proceso.
Ha sido desarrollado por el área ed Mejorea e Innovación para el uso, prueba y testeo del área operativa.

### Mantenimiento técnico

Repositorio mantenido por el equipo de Mejora e Innovación.

## Licencia

Repositorio de uso interno

Mejora e Innovación - Abril 2026

