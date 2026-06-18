/**
 * Módulo para la comunicación con la API de Google Sheets y gestión de caché local.
 * ACTUALIZADO: Para trabajar con un solo libro y múltiples pestañas.
 */
const GoogleSheet = {
    // --- URL ÚNICA DE LA APP WEB ---
    // ¡¡¡IMPORTANTE: PEGA AQUÍ LA NUEVA URL QUE OBTUVISTE AL IMPLEMENTAR EL SCRIPT!!!
    //estadistica API_URL: 'https://script.google.com/macros/s/AKfycbw71yCt8jVIIivLRjbRjDNSgixxYt4eiUG5ShNtsgLbXohYAY-Ze0ngwLEydPvoBeVW/exec', 
    API_URL: 'https://script.google.com/macros/s/AKfycbwAYn2vsPsuR2c6ixfj4epCgVaVf-9jLrDTPGDoaUmXYyIklnzm4kQzqGbtNDlcMjktBw/exec', 

    // ---------------------------------

    CLAVE_BASE_CACHE: 'nombresJugadoresCache',
    TIEMPO_EXPIRACION: 86400000, // 24 horas

    /**
     * Obtiene la lista de nombres enviando el grupo como parámetro GET.
     */
    obtenerNombres: async function(group) {
        if (!this.API_URL || this.API_URL.includes("TU_NUEVA_URL")) {
            console.error("Error: No has configurado la API_URL en comunicacion.js");
            return [[], null];
        }
        
        const cacheKey = `${this.CLAVE_BASE_CACHE}-${group}`;
        const ahora = new Date().getTime();
        const datosCache = localStorage.getItem(cacheKey);

        // 1. INTENTAR CARGAR DESDE CACHÉ
        if (datosCache) {
            try {
                const cache = JSON.parse(datosCache);
                if (ahora < cache.expiracion) {
                    console.log(`Cargando nombres del grupo ${group} desde LocalStorage (RÁPIDO).`);
                    return [cache.nombres, this.API_URL];
                } else {
                    localStorage.removeItem(cacheKey);
                }
            } catch (e) {
                localStorage.removeItem(cacheKey);
            }
        }

        // 2. CONSTRUIR URL CON PARÁMETRO
        // Añadimos ?grupo=8-3 al final de la URL para que el doGet sepa qué pestaña leer
        const fetchUrl = `${this.API_URL}?grupo=${group}`;

        // 3. CARGAR DESDE GOOGLE APPS SCRIPT
        console.log(`Cargando nombres del grupo ${group} desde Google Apps Script. URL: ${fetchUrl}`);
        try {
            const response = await fetch(fetchUrl);
            if (!response.ok) {
                throw new Error("Error en la respuesta del servidor.");
            }
            const nombresArray = await response.json();

            // 4. GUARDAR EN CACHÉ
            if (nombresArray && Array.isArray(nombresArray) && nombresArray.length > 0) {
                const nuevaExpiracion = ahora + this.TIEMPO_EXPIRACION;
                const nuevoCache = {
                    nombres: nombresArray,
                    expiracion: nuevaExpiracion
                };
                localStorage.setItem(cacheKey, JSON.stringify(nuevoCache));
            } else {
                // Si viene un error o array vacío
                console.warn("La API devolvió datos vacíos o error:", nombresArray);
                return [[], this.API_URL];
            }

            return [nombresArray, this.API_URL]; 
        } catch (error) {
            console.error(`Error al obtener los nombres del grupo ${group}:`, error);
            return [[], this.API_URL];
        }
    },

    /**
     * Guarda el resultado. Ahora incluye el grupo en el cuerpo de los datos.
     */
    guardarResultado: async function(url, datosJuego) {
        if (!url) {
            return { status: "error", message: "URL no definida." };
        }
        
        console.log("Enviando datos a Google Sheet (Libro Unificado):", datosJuego);

        try {
            const response = await fetch(url, {
                method: 'POST',
                mode: 'no-cors', 
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8',
                },
                body: JSON.stringify(datosJuego), // datosJuego ya debe incluir la propiedad 'grupo'
                redirect: 'follow'
            });

            console.log("Datos enviados (modo no-cors).");
            return { status: "success", message: `Puntaje guardado correctamente en la lista del grupo ${datosJuego.grupo}.` };

        } catch (error) {
            console.error("Error al guardar el resultado:", error);
            return { status: "error", message: error.message };
        }
    }
};