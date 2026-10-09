// Server endpoint. The task pane is served from https://localhost:5173 and the
// API runs on https://localhost:3001 (both using `office-addin-dev-certs`).
// Override with VITE_API_BASE in a .env.local file if you deploy elsewhere.
export const API_BASE =
  (import.meta.env.VITE_API_BASE as string | undefined) ?? "https://localhost:3001";
