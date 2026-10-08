# TeamsBrief Web

Version web hébergée — 3 fichiers client + proxy API.

| Fichier | Rôle |
|---------|------|
| `index.html` | Interface |
| `app.js` | Logique (import VTT, résumé Groq, agent chat) |
| `style.css` | Design |
| `api/groq.js` | Proxy serveur Vercel |
| `netlify/functions/groq.js` | Proxy serveur Netlify |
| `functions/api/groq.js` | Proxy serveur Cloudflare Pages |

## Hébergements

| Plateforme | URL |
|------------|-----|
| **Vercel** | https://teamsbrief-web.vercel.app |
| **GitHub Pages** | https://ghostyesal-oss.github.io/teamsbrief-web/ |
| **Cloudflare Pages** | _(optionnel, voir ci-dessous)_ |
| **Netlify** | _(optionnel, voir ci-dessous)_ |

### Vercel

```bash
cd teamsbrief-web
npx vercel --prod
```

### Cloudflare Pages

```bash
cd teamsbrief-web
npx wrangler pages deploy . --project-name teamsbrief-web
```

### Netlify

```bash
cd teamsbrief-web
npx netlify login
npx netlify deploy --prod --dir .
```

## Utilisation

1. Ouvrez l'URL hébergée
2. **Paramètres** → clé Groq `gsk_...`
3. **Importer .vtt** → résumé + agent

Données stockées dans le navigateur (localStorage).
