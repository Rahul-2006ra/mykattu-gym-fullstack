# MYKATTU Gym Full Stack

AI-powered gym website for MYKATTU Gym, Hesaraghatta, Bengaluru.

Live demo: https://rahul-2006ra.github.io/gym/

## What is included

- Static frontend in `index.html`
- Express backend in `backend/`
- Lead capture API with Excel storage and email notification support
- AI coach chat proxy with a built-in fallback so the demo still answers useful gym questions if the AI provider is unavailable
- BMI calculator, workout generator, pricing modal, contact form, and chat widget

## Run backend locally

```bash
cd backend
npm install
npm start
```

The backend runs on `http://localhost:3001`.

## Backend environment variables

Create `backend/.env` from `backend/.env.example`.

```env
PORT=3001
OWNER_EMAIL=your-email@example.com
GMAIL_USER=your-gmail@example.com
GMAIL_APP_PASS=your-gmail-app-password
ANTHROPIC_API_KEY=your-anthropic-api-key
ANTHROPIC_MODEL=claude-haiku-4-5
```

The site still works without the email/API keys for demo purposes, but production lead emails and live AI responses need those values.

## Frontend

Open `index.html` directly or serve it with any static server:

```bash
python -m http.server 5500
```

Then open `http://localhost:5500`.

## Verification

Tested:

- Public frontend loads from GitHub Pages
- BMI calculator works
- Workout generator works
- Pricing modal works
- AI feature modal works
- Chat widget returns a useful fallback answer when the deployed AI proxy fails
- Local backend health, chat, and lead APIs work
