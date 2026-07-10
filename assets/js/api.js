window.Api = class Api {
  constructor(baseUrl, tokenProvider) {
    this.baseUrl = baseUrl.replace(/\/$/, '');
    this.tokenProvider = tokenProvider;
  }

  async request(path, options = {}) {
    const token = await this.tokenProvider();
    const headers = new Headers(options.headers || {});
    headers.set('Authorization', `Bearer ${token}`);
    if (options.body && !headers.has('Content-Type')) headers.set('Content-Type', 'application/json');
    const url = `${this.baseUrl}${path}`;
    let response;
    try {
      response = await fetch(url, { ...options, headers });
    } catch (error) {
      throw new Error(`Could not reach the API (${url}). Check the browser console for the blocked request details.`);
    }
    let data = {};
    try { data = await response.json(); } catch { data = { error: 'The server returned an unreadable response.' }; }
    if (!response.ok) {
      const error = new Error(data.error || `Request failed (${response.status})`);
      error.status = response.status;
      throw error;
    }
    return data;
  }

  get(path) { return this.request(path); }
  post(path, body) { return this.request(path, { method: 'POST', body: JSON.stringify(body || {}) }); }
  put(path, body) { return this.request(path, { method: 'PUT', body: JSON.stringify(body || {}) }); }

  async postFormWithToken(path, body) {
    const token = await this.tokenProvider();
    const url = `${this.baseUrl}${path}`;
    const form = new URLSearchParams();
    form.set('accessToken', token);
    Object.entries(body || {}).forEach(([key, value]) => {
      form.set(key, Array.isArray(value) ? JSON.stringify(value) : String(value ?? ''));
    });

    let response;
    try {
      response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: form.toString()
      });
    } catch {
      throw new Error(`Could not reach the API (${url}) from ${location.origin}.`);
    }

    let data = {};
    try { data = await response.json(); } catch { data = { error: 'The server returned an unreadable response.' }; }
    if (!response.ok) {
      const error = new Error(data.error || `Request failed (${response.status})`);
      error.status = response.status;
      throw error;
    }
    return data;
  }
};
