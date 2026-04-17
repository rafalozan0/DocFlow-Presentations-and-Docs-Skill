# Security | Seguridad

## English

This project is local-first:
- No automatic uploads to external services.
- Document processing runs locally.

Sensitive data guidance:
- Never hardcode credentials in source files.
- Use environment variables (`OFFICE_EMAIL_PASSWORD`).
- Keep `.env` out of git.
- Use least-privilege permissions on config files.

Email sending notes:
- SMTP credentials are sensitive.
- Prefer app passwords/tokens over account passwords.

Current security posture summary:
- No dynamic `eval` or `exec` pathways.
- No `os.system` and no `shell=True` subprocess pattern.
- Network calls are limited to explicit SMTP usage in `send_email`.
- Conversion command uses fixed executable/args (`libreoffice --headless`).

## Español

Este proyecto es local-first:
- No sube archivos automáticamente a servicios externos.
- El procesamiento de documentos ocurre localmente.

Guía para datos sensibles:
- Nunca dejes credenciales en código fuente.
- Usa variables de entorno (`OFFICE_EMAIL_PASSWORD`).
- Mantén `.env` fuera de git.
- Usa permisos mínimos necesarios en archivos de configuración.

Notas sobre correo:
- Las credenciales SMTP son sensibles.
- Prefiere contraseñas de aplicación/tokens sobre contraseña de cuenta.

Resumen de seguridad actual:
- No hay uso de `eval` ni `exec`.
- No hay `os.system` ni uso de `shell=True` en subprocess.
- No hay llamadas de red excepto uso explícito de SMTP en `send_email`.
- La conversión usa comando fijo (`libreoffice --headless`).
