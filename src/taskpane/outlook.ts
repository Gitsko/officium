/* global document, Office */

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  document.getElementById("sideload-msg")!.hidden = true;
  const hostInfo = document.getElementById("host-info")!;
  hostInfo.textContent = `Hallo aus Outlook (${info.platform}).`;
  hostInfo.hidden = false;
});
