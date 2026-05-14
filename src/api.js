// Thin API client. Falls back gracefully if backend not reachable.
const BASE = import.meta.env.VITE_API_BASE || "/api";

async function req(path, opts = {}) {
  const token = localStorage.getItem("ntpas_token");
  const res = await fetch(`${BASE}${path}`, {
    headers: {
      "Content-Type": "application/json",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...(opts.headers || {}),
    },
    ...opts,
  });
  if (!res.ok) throw new Error(`API ${path} ${res.status}`);
  return res.json();
}

export const api = {
  health: () => req("/health"),
  login: (email, password) =>
    req("/auth/login", { method: "POST", body: JSON.stringify({ email, password }) }),
  me: () => req("/auth/me"),
  // generic CRUD
  list: (resource) => req(`/${resource}`),
  create: (resource, body) => req(`/${resource}`, { method: "POST", body: JSON.stringify(body) }),
  update: (resource, id, body) => req(`/${resource}/${id}`, { method: "PUT", body: JSON.stringify(body) }),
  remove: (resource, id) => req(`/${resource}/${id}`, { method: "DELETE" }),
};

export default api;
