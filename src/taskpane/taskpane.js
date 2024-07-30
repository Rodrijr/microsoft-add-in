/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemChanged, loadResourceInformation);
    loadResourceInformation();
  }
});

function getCustomFieldFromLocation(location) {
  // Replace with logic to extract custom field from the location string
  // Assuming the custom field is within square brackets in the location string, e.g., "Room A [customField]"
  const match = location.match(/\[(.*?)\]/);
  return match ? match[1] : null;
}

function loadResourceInformation() {
  const item = Office.context.mailbox.item;
  if (item.location) {
    const customField = getCustomFieldFromLocation(item.location);
    if (customField) {
      updateIframe(customField);
    }
  }
}

function updateIframe(customField) {
  try {
    const iframe = document.createElement("iframe");
    iframe.src = `https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=${customField}`;
    iframe.id = 'miIframe';
    iframe.referrerPolicy = "strict-origin-when-cross-origin";
    const existingIframe = document.getElementById("miIframe");
    if (existingIframe) {
      existingIframe.remove();
    }
    document.getElementById("preview").appendChild(iframe);
  } catch (error) {
    console.error('Error loading iframe:', error);
  }
}
