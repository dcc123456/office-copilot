import React from "react";
import ReactDOM from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { App } from "./App.js";
import "./styles.css";

// Office.onReady resolves both in the real host and when office.js is loaded
// outside of Office (it resolves with host: null). We mount in both cases so
// the UI remains usable for local development in a browser.
Office.onReady(() => {
  ReactDOM.createRoot(document.getElementById("root")!).render(
    <React.StrictMode>
      <FluentProvider theme={webLightTheme}>
        <App />
      </FluentProvider>
    </React.StrictMode>,
  );
});
