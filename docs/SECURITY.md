# Security | Seguridad

## English

This project is local-first:
- No automatic uploads to external services.
- Document and template processing runs locally.

Sensitive data guidance:
- Never hardcode credentials in source files.
- Keep `.env` out of git.
- Use least-privilege permissions on config files.

Current security posture summary:
- No dynamic `eval` or `exec` pathways.
- No `os.system` and no `shell=True` subprocess pattern.
- No built-in email sending functionality.
- Conversion command uses fixed executable/args (`libreoffice --headless`).
- HTML rendering uses fixed headless Chrome args and supports `OFFICE_CHROME_BIN` override.

## Español

Este proyecto es local-first:
- No sube archivos automáticamente a servicios externos.
- El procesamiento de documentos y templates ocurre localmente.

Guía para datos sensibles:
- Nunca dejes credenciales en código fuente.
- Mantén `.env` fuera de git.
- Usa permisos mínimos necesarios en archivos de configuración.

Resumen de seguridad actual:
- No hay uso de `eval` ni `exec`.
- No hay `os.system` ni uso de `shell=True` en subprocess.
- No incluye funcionalidad de envío de email integrada.
- La conversión usa comando fijo (`libreoffice --headless`).
- El render HTML usa argumentos fijos de Chrome headless y permite override con `OFFICE_CHROME_BIN`.
