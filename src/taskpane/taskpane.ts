import { formulaHandler } from "./formulaHandler";

/* global console, document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("bind").onclick = bind;
    document.getElementById("rebind").onclick = rebind;
    document.getElementById("unbind").onclick = unbind;

    bind();
  }
});

export async function bind() {
  try {
    await formulaHandler.bindExcelEvents();
  } catch (error) {
    console.error(error);
  }
}

export async function rebind() {
  try {
    await formulaHandler.unbindExcelEvents();
    await formulaHandler.bindExcelEvents();
  } catch (error) {
    console.error(error);
  }
}

export async function unbind() {
  try {
    await formulaHandler.unbindExcelEvents();
  } catch (error) {
    console.error(error);
  }
}
