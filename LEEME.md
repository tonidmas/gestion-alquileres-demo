# Susalquia — Demo

Esta es la versión de demostración de Susalquia: **funciona sin ninguna base de datos ni cuenta de
Supabase**. Todos los datos son de ejemplo y viven solo en la memoria del navegador — al recargar la
página, o al pulsar "Reiniciar demo", todo vuelve a su estado inicial.

Pensada para enseñar la app a un cliente potencial sin necesidad de darle acceso a datos reales, y sin
ningún riesgo de que alguien manipule información real durante la demostración.

## Qué incluye

- Datos de ejemplo ya cargados: 12 habitaciones (9 ocupadas, 3 libres), gastos de varios meses, alertas
  activas y un inquilino en el histórico.
- Todas las funciones son operativas de verdad dentro de la sesión: editar fichas, marcar pagos, subir
  documentos (contrato, DNI, facturas — se guardan solo en el navegador mediante `URL.createObjectURL`,
  no en ningún servidor), generar un contrato en Word real con la función "Generar contrato", configurar
  el modelo de gestión, etc.
- Botón **"Reiniciar demo"** en la barra lateral (y flotante en móvil) para volver a los datos originales
  en cualquier momento.
- Banda superior permanente que recuerda que se trata de una demo.

## Cómo desplegarla

Igual que el proyecto principal: es un proyecto Vite + React independiente.

1. Sube todo el contenido de esta carpeta a un repositorio de GitHub (puede ser uno nuevo, distinto del
   proyecto real — así puedes compartir la demo sin dar acceso al código de producción).
2. Impórtalo en Vercel como un nuevo proyecto.
3. No hace falta configurar ninguna variable de entorno — a diferencia del proyecto principal, esta
   versión no usa Supabase.
4. Despliega. Ya está lista para compartir el enlace con quien quieras.

## Diferencias con la versión real (conectada a Supabase)

| | Demo | Versión real |
|---|---|---|
| Guarda los datos | No, solo en memoria del navegador | Sí, en la nube (Supabase) |
| Requiere iniciar sesión | No | Sí |
| Documentos y facturas | Se pueden subir y ver, pero se pierden al recargar | Se guardan de forma permanente |
| Generar contrato en Word | Sí, funciona igual | Sí, funciona igual |
| Coste de mantenimiento | Ninguno (no usa Supabase) | El de tu plan de Supabase/Vercel |

Cuando quieras actualizar la demo con cambios nuevos de la app real, dímelo y la vuelvo a sincronizar.
