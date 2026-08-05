# Gestión de Alquileres — versión DEMO

Esta es una versión especial de la app pensada para **enseñar a posibles clientes**, no para uso real.

## Qué la diferencia de la versión completa

- **No tiene base de datos.** No usa Supabase ni ningún otro servicio externo — todo funciona con datos que viven solo en la memoria del navegador mientras la página está abierta.
- **Viene precargada con inquilinos de ejemplo** (habitaciones A1, A2, B1, E2 ocupadas, y una en el Histórico), gastos del mes ya rellenados, y algún pago pendiente a propósito, para que el Dashboard, las Alertas y el Histórico se vean con contenido real desde el primer segundo.
- **Nada se guarda de forma permanente.** La persona que la pruebe puede añadir inquilinos, marcar pagos, liberar habitaciones, cambiar gastos... y al recargar la página (F5) todo vuelve a los datos de ejemplo originales.
- Se ve un aviso permanente de **"MODO DEMO"** en la barra lateral y un banner en la parte superior para que quede claro en todo momento que es una demostración.
- Hay un botón **"Reiniciar demo"** en la barra lateral para volver a los datos de ejemplo en cualquier momento, sin tener que recargar la página.

## Cómo desplegarla (idéntico al proceso que ya conoces)

Es exactamente el mismo proceso que usaste para la app completa:

1. Sube todos los archivos de esta carpeta a un repositorio de GitHub (puedes llamarlo, por ejemplo, `gestion-alquileres-demo`)
2. Entra en vercel.com → Add New → Project → importa ese repositorio
3. Deja la configuración por defecto y pulsa **Deploy**

**Diferencia importante:** esta vez **no hace falta configurar ninguna variable de entorno** (no hay `VITE_SUPABASE_URL` ni `VITE_SUPABASE_ANON_KEY` que añadir), porque la demo no se conecta a ninguna base de datos. Así que el despliegue es todavía más sencillo que el de la app real.

Si tienes dudas con cualquier paso del proceso de GitHub/Vercel, el archivo `GUIA_VERCEL.md` incluido en esta misma carpeta lo explica con más detalle paso a paso.

## Sugerencia de uso comercial

Puedes compartir la URL de esta demo (`tu-demo.vercel.app`) directamente con posibles clientes para que la prueben por su cuenta, sin ningún riesgo de que estropeen datos reales ni de que un cliente vea los datos de otro. Cuando quieran contratar el servicio de verdad, ahí sí se les daría de alta su propia instancia con base de datos real (ver el documento del plan técnico de comercialización).
