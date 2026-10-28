# mayITa — Retroalimentación de procesos

Réplica de mayITa Proyectos con la encuesta **Proceso Retroalimentación**.
Se mantiene el inicio de sesión con Microsoft, los juegos (Flappy mayITa y Torre de Bloques)
y el guardado vía n8n. Se reemplazó la encuesta Likert por las preguntas nuevas.

## Pendiente antes de publicar

1. **`N8N_WEBHOOK_URL`** en `assets/js/app.js` está en `PENDIENTE_URL_WEBHOOK_N8N`.
   Hay que crear un flujo nuevo en n8n (ver "Datos que recibe n8n"). En `localhost` el guardado se
   simula para poder probar la interfaz; en producción mostrará un error hasta que se configure.
2. **Redirect URI**: ahora se toma de la URL donde se publique el sitio. Debe registrarse exactamente
   igual como *Single-page application (SPA)* en la app de Entra ID (`MSAL_CONFIG.auth.clientId`).
3. **Ranking de Flappy**: `N8N_RANKING_SAVE_URL` / `N8N_RANKING_TOP_URL` siguen apuntando a los del
   proyecto original, por lo que el ranking sería compartido. Cámbialos si quieres uno separado.
4. **Torre de Bloques** no cambió: sigue usando su propio banco de preguntas (`TOWER_API_BASE` en
   `tower-blocks.js` y la carpeta `server/`). Son preguntas del juego, no de esta encuesta.

## Cómo editar las preguntas

Todo está en `SURVEY_QUESTIONS` (`assets/js/app.js`). Tipos: `yesno`, `single`, `multi`, `text`.
Una pregunta puede mostrarse solo según otra respuesta con `showIf: { id, equals }`.
Los `id` son los nombres de campo al guardar; no los cambies si ya hay respuestas guardadas.

## Datos que recibe n8n (POST JSON)

```json
{
  "usuario": { "nombre": "", "correo": "", "entraId": "" },
  "encuesta": "Retroalimentación de procesos",
  "respuestas": [
    { "numeroPregunta": 1, "id": "induccion", "pregunta": "¿Recibiste inducción…?",
      "tipoRespuesta": "yesno", "respuesta": "Sí" }
  ],
  "respuestasPorId": { "induccion": "Sí", "siguiente_nivel": "…", "tipo_modificacion": "Frecuencia; Fecha" },
  "puntajeJuego": 0,
  "origen": "mayITa Retroalimentación",
  "versionApp": "2.0"
}
```

- `respuestas`: solo las preguntas que se mostraron (una fila por respuesta, si prefieres ese formato).
- `respuestasPorId`: todas las preguntas, con `""` en las que no aplicaron (una fila por encuesta).
- Las de selección múltiple van como arreglo en `respuestas` y unidas con `; ` en `respuestasPorId`.

n8n debe responder `{ "ok": true, "codigoEncuesta": "…", "respuestasGuardadas": N }`
(los dos últimos son opcionales). Si `ok` no es `true`, la encuesta muestra el error con opción de reintentar.

## Prueba local

```bash
python3 -m http.server 8080   # y abrir http://localhost:8080
```
