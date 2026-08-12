import { setUpFindingsPanel } from "./findingsPanel";
import { setUpCvssPanel } from "./cvssPanel";
import { setUpHttpPanel } from "./httpPanel";
import { byId, show } from "./dom";

/* global document, Office, HTMLButtonElement, NodeListOf */

const PANELS = ["findings", "cvss", "http"];

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

/* Wire the pane once Office is ready; every panel talks to the document. */
Office.onReady(() => {
  setUpTabs();
  setUpFindingsPanel();
  setUpCvssPanel();
  setUpHttpPanel();
});
