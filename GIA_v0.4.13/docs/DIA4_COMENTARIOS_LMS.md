# Día 4 — Comentarios de evidencia estilo LMS

## Regla funcional

Los comentarios forman una conversación lineal contextual de la actividad/evidencia. La conversación permanece después de una reentrega porque la entrega (`EvidenceSubmission`) conserva su identidad y sus intentos se mantienen en `EvidenceSubmissionAttempt`.

- Aprendiz: puede comentar y ver comentarios del instructor.
- Instructor de seguimiento / líder: puede comentar y ver comentarios del aprendiz.
- Soporte y certificador: gestionan el dominio, pero no son autores de esta conversación educativa.
- Un comentario normal no cambia el estado.
- Un comentario marcado como `is_correction_request` solo lo puede crear un instructor y, cuando la evidencia está pendiente de revisión, cambia el estado a `requiere_correccion` y dispara la notificación correspondiente.
- Los comentarios nuevos guardan referencia al intento vigente (`attempt_id`) cuando existe; los comentarios históricos pueden quedar sin esa referencia sin perderse.
- Las notificaciones informan del evento, pero no sustituyen la conversación.
