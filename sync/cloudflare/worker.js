export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // Health check
    if (request.method === 'GET' && url.pathname === '/health') {
      return json({ ok: true, service: 'ntn-sync-worker' });
    }

    // Get companies
    if (request.method === 'GET' && url.pathname === '/companies') {
      const result = await env.DB.prepare(
        "SELECT * FROM companies LIMIT 100"
      ).all();

      return json(result.results);
    }

    // Add company
    if (request.method === 'POST' && url.pathname === '/companies') {
      const body = await request.json();

      const company = body.company;
      const ntn = body.ntn;
      const cnic = body.cnic || "";
      const reff = body.reff || "";

      await env.DB.prepare(
        "INSERT INTO companies (company, ntn, cnic, reff) VALUES (?, ?, ?, ?)"
      ).bind(company, ntn, cnic, reff).run();

      return json({ ok: true });
    }

    return json({ ok: false, error: 'Not found' }, 404);
  }
};

function json(data, status = 200) {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: {
      'content-type': 'application/json',
      'access-control-allow-origin': '*'
    }
  });
}
