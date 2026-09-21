# Sponti CRM

CRM starter application for Sponti.

## Run locally

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

The production build is created in `dist/`.

## Netlify

The project is Vite-based.

- Build command: `npm run build`
- Publish directory: `dist`

You can connect the GitHub repository directly in Netlify.

## Excel

The Excel importer currently accepts common German/English column names:
- Vorname / First Name
- Nachname / Last Name
- E-Mail / Email
- Telefon / Phone
- Ort / City
- Dienstleistung / Service
- Status
- Anmeldedatum / Signup Date

This version keeps data in browser memory. For a real multi-user CRM, connect the app to a hosted database such as Supabase before using it in production.
