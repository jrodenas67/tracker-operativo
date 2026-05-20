exports.handler = async () => {
  const token = process.env.GITHUB_TOKEN;
  if (!token) {
    return { statusCode: 500, body: JSON.stringify({ error: "Token no configurado" }) };
  }

  const res = await fetch(
    "https://api.github.com/repos/jrodenas67/tracker-operativo/actions/workflows/update.yml/dispatches",
    {
      method: "POST",
      headers: {
        Authorization: `token ${token}`,
        Accept: "application/vnd.github+json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ ref: "main" }),
    }
  );

  if (res.status === 204) {
    return { statusCode: 200, body: JSON.stringify({ ok: true }) };
  }
  const text = await res.text();
  return { statusCode: res.status, body: JSON.stringify({ error: text }) };
};
