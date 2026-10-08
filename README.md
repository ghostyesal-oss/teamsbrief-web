# TeamsBrief Web

Version web hébergée — 3 fichiers client + proxy API.

| Fichier | Rôle |
|---------|------|
| `index.html` | Interface |
| `app.js` | Logique (import VTT, résumé Groq, agent chat) |
| `style.css` | Design |
| `api/groq.js` | Proxy serveur (Vercel, pour éviter CORS) |

## Héberger sur Vercel (gratuit)

```bash
cd teamsbrief-web
npx vercel --prod
```

## Utilisation

1. Ouvrez l'URL hébergée
2. **Paramètres** → clé Groq `gsk_...`
3. **Importer .vtt** → résumé + agent

Données stockées dans le navigateur (localStorage).
