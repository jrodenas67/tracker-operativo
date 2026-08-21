// Cloudflare Pages advanced mode: /trigger-update dispara el workflow; el resto sirve estático.
export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === '/trigger-update') {
      const json = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { 'Content-Type': 'application/json' } });
      if (request.method !== 'POST') return json({ error: 'Usa POST' }, 405);
      const token = env.GITHUB_TOKEN;
      if (!token) return json({ error: 'Token no configurado' }, 500);
      const res = await fetch('https://api.github.com/repos/jrodenas67/tracker-operativo/actions/workflows/update.yml/dispatches', {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json', 'Content-Type': 'application/json', 'User-Agent': 'tracker-operativo-taperia' },
        body: JSON.stringify({ ref: 'main' }),
      });
      return res.status === 204 ? json({ ok: true }) : json({ error: await res.text() }, res.status);
    }
    return env.ASSETS.fetch(request);   // estático (index.html, tracker-operativo.html)
  },
};
