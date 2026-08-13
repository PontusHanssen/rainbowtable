import { setUpFindingsPanel } from "./findingsPanel";
import { setUpCvssPanel } from "./cvssPanel";
import { setUpHttpPanel } from "./httpPanel";
import { setUpMarkdownPanel } from "./markdownPanel";
import { setUpLimitsPanel } from "./limitsPanel";
import { byId, show } from "./dom";

/* global document, Office, HTMLButtonElement, NodeListOf, location */

const PANELS = ["findings", "cvss", "http", "markdown"];

function setUpTabs(): void {
  const tabs = document.querySelectorAll(".tab") as NodeListOf<HTMLButtonElement>;

  tabs.forEach((tab) =>
    tab.addEventListener("click", () => {
      const chosen = tab.dataset.panel;

      tabs.forEach((other) =>
        other.setAttribute("aria-selected", String(other.dataset.panel === chosen))
      );
      PANELS.forEach((panel) => show(byId(`panel-${panel}`), panel === chosen));
    })
  );
}

/**
 * Mark the pane when it is served from the dev server. The development and production
 * add-ins can be installed side by side, and two identical panes are easy to confuse.
 */
function markDevelopmentBuild(): void {
  if (location.hostname !== "localhost") {
    return;
  }
  const badge = document.createElement("div");
  badge.textContent = "dev";
  badge.style.cssText =
    "position:fixed;top:0;right:0;padding:1px 6px;font-size:10px;font-weight:600;" +
    "background:#8a5300;color:#fff;border-bottom-left-radius:4px;z-index:1";
  document.body.appendChild(badge);
}

/* Wire the pane once Office is ready; every panel talks to the document. */
Office.onReady(() => {
  markDevelopmentBuild();
  setUpTabs();
  setUpFindingsPanel();
  setUpCvssPanel();
  setUpHttpPanel();
  setUpMarkdownPanel();
  setUpLimitsPanel();
});
