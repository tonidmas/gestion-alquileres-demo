# Cómo subir tu app a Vercel (paso a paso)

No necesitas saber programar. Solo sigue estos pasos con calma.

## Antes de empezar
Descomprime el archivo `.zip` que te he dado. Dentro verás una carpeta con varios archivos (`package.json`, `index.html`, una carpeta `src`, etc.). Esa carpeta completa es tu proyecto.

---

## Paso 1: Crear una cuenta en GitHub
GitHub es donde vamos a guardar el código de tu app (es gratis).

1. Ve a **github.com**
2. Haz clic en **Sign up** (crear cuenta)
3. Sigue los pasos con tu correo y una contraseña

---

## Paso 2: Crear un repositorio nuevo
Un "repositorio" es como una carpeta donde vivirá tu proyecto.

1. Ya dentro de GitHub, haz clic en el botón verde **New** (o el símbolo **+** arriba a la derecha → "New repository")
2. Ponle un nombre, por ejemplo: `gestion-alquileres`
3. Déjalo como **Public** o **Private** (como prefieras)
4. Haz clic en **Create repository**
5. En la página que aparece, busca el enlace que dice **uploading an existing file** (subir un archivo existente)

---

## Paso 3: Subir tu proyecto
1. Abre la carpeta que descomprimiste en tu ordenador
2. Selecciona **todos los archivos y carpetas de dentro** (no la carpeta en sí, su contenido)
3. Arrástralos a la página de GitHub donde dice "Drag files here"
4. Espera a que termine de cargar
5. Abajo, haz clic en **Commit changes** (guardar cambios)

Ya tienes tu proyecto en GitHub. ✅

---

## Paso 4: Crear una cuenta en Vercel
Vercel es el servicio gratuito que va a "publicar" tu app en internet.

1. Ve a **vercel.com**
2. Haz clic en **Sign Up**
3. Elige la opción **Continue with GitHub** (así se conectan automáticamente)
4. Autoriza el acceso cuando te lo pida

---

## Paso 5: Publicar tu app
1. Dentro de Vercel, haz clic en **Add New...** → **Project**
2. Verás una lista de tus repositorios de GitHub. Busca `gestion-alquileres` y haz clic en **Import**
3. Vercel detecta automáticamente que es un proyecto Vite/React. No cambies nada.
4. Haz clic en el botón **Deploy**
5. Espera 1-2 minutos mientras Vercel prepara todo

---

## Paso 6: ¡Listo!
Cuando termine, Vercel te mostrará una URL como:

`https://gestion-alquileres-tu-usuario.vercel.app`

Esa es la dirección de tu aplicación. Puedes entrar desde cualquier ordenador o móvil.

Cada vez que quieras hacer un cambio en el futuro, solo tendrás que subir el archivo modificado a GitHub, y Vercel actualizará la app sola, automáticamente.

---

## Algo importante que debes saber
Esta app guarda los datos **en el navegador de cada dispositivo** (no en un servidor). Esto significa:

- Si abres la app desde el ordenador de tu oficina, los datos se guardan ahí.
- Si la abres desde tu móvil, empezará vacía, con sus propios datos.
- Si borras el historial/caché del navegador, los datos de la app se perderán.

👉 Por eso, te recomiendo usar siempre **el mismo navegador, en el mismo ordenador**, y hacer de vez en cuando una copia de seguridad exportando tus inquilinos a Excel desde el botón "Exportar" de la app.

---

## Nota: esta versión ya no tiene subcarpetas

He simplificado la estructura del proyecto: ahora **todos los archivos están sueltos, sin ninguna carpeta dentro**. Esto significa que, subas los archivos por donde los subas (GitHub o directamente en Vercel), puedes arrastrarlos todos de una sola vez, sin tener que entrar a ninguna subcarpeta después. Así se evita el error que tuviste la vez anterior (Vercel no encontraba `src/main.jsx` porque esa carpeta no se creó bien).

Si ya tienes un repositorio en GitHub de un intento anterior, lo más simple es:
1. Borra todos los archivos que subiste antes (o borra el repositorio entero y crea uno nuevo)
2. Sube de una sola vez **todos** los archivos de esta nueva carpeta descomprimida (selecciona todo y arrastra)
3. Vercel detectará el cambio y volverá a desplegar automáticamente en 1-2 minutos

