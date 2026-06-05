# Docker / Cloud Deployment

This project is prepared to run in a Docker container with LibreOffice and fonts installed inside the same runtime.

## Why Docker

DOCX to PDF conversion is sensitive to:

- LibreOffice version
- Chinese and English fonts
- operating system font fallback
- headless profile state

Docker keeps these stable so the PDF layout on Alibaba Cloud should match the test environment more closely.

## Included Runtime

The `Dockerfile` installs:

- Python 3.12
- LibreOffice Writer
- Noto CJK fonts for Chinese text
- Noto core fonts and Liberation fonts for English text
- Poppler utilities for optional PDF to image checks

The app uses `/usr/bin/soffice` in Docker. On Windows local development it can still use `SOFFICE_PATH` if needed.

## Local Docker Test

Build and start:

```powershell
docker compose up --build
```

Open:

```text
http://127.0.0.1:8088
```

Default login:

```text
admin / admin123
```

Stop:

```powershell
docker compose down
```

## Persistent Data

`docker-compose.yml` keeps these folders as Docker volumes:

- `/app/app/data`: users, sessions, common people, rules, template registry
- `/app/app/uploads`: uploaded Excel files
- `/app/app/generated`: generated PDF ZIP files
- `/app/app/doc_templates`: active Word templates

Template uploads from the admin backend are written into `/app/app/doc_templates`, and old templates are backed up under `/app/app/data/template_versions`.

## Alibaba Cloud Notes

Recommended first deployment:

1. Install Docker and Docker Compose on the ECS server.
2. Copy the project folder to the server.
3. Run `docker compose up --build -d`.
4. Put Nginx in front of port `8088`.
5. Enable HTTPS.
6. Change the default admin password immediately.
7. Restrict access by IP, VPN, or private network if this is only for internal staff.

The site should not be exposed publicly without HTTPS, password changes, backups, and access control.

## Useful Checks

Check LibreOffice inside the container:

```powershell
docker compose exec secretary-files soffice --version
```

Check installed fonts:

```powershell
docker compose exec secretary-files fc-list | findstr /i "Noto"
```

Check app logs:

```powershell
docker compose logs -f secretary-files
```
