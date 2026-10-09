import React from "react";
import ReactDOM from "react-dom/client";
import { App } from "./App.js";
import { applyTheme, readStoredTheme } from "./uiStore.js";
import "./styles.css";

applyTheme(readStoredTheme());

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
