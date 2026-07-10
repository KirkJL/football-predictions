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
    const response = await fetch(`${this.baseUrl}${path}`, { ...options, headers });
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
};
