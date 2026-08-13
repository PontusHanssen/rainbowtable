import { createRoot } from "react-dom/client";
import { App } from "./components/App";

/* global document */

/**
 * The authoring dialog.
 *
 * It has no access to the document — that is the pane's job — so nothing here touches
 * Office.js beyond the message channel in `usePane`.
 */
const host = document.getElementById("app");
if (host) {
  createRoot(host).render(<App />);
}
