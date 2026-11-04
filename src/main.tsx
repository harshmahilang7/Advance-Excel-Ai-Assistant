import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { App } from "./ui/App";

const mount = () => {
  const root = document.getElementById("root");
  if (!root) throw new Error("#root not found");
  createRoot(root).render(
    <StrictMode>
      <FluentProvider theme={webLightTheme} style={{ minHeight: "100vh" }}>
        <App />
      </FluentProvider>
    </StrictMode>,
  );
};

// Office.js may be missing when opened in a plain browser (dev). Mount either way.
if (typeof Office !== "undefined") {
  void Office.onReady().then(mount);
} else {
  mount();
}
