# Teams Live Chat Bridge

Production bridge between WordPress website chat widgets and Microsoft Teams.

## Endpoints

- `POST /api/messages` — Microsoft Teams bot endpoint
- `POST /wp/message` — WordPress to Teams
- `GET /health` — health check

## Required environment variables

- `CLIENT_ID`
- `CLIENT_SECRET`
- `TENANT_ID`
- `APP_NAME`
- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`

Do not commit real secrets to this repository.

## Agent commands in Teams

- `/online`
- `/away`
- `/offline`
