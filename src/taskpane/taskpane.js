/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('DEBIO LLEGAR 2')
    loadResourceInformation();
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemChanged, loadResourceInformation);
  }
});

function getCustomFieldFromLocation(location) {
  console.log(typeof location)
  // if(typeof location)
  const match = location.match(/\[(.*?)\]/);
  return match ? match[1] : null;
}

function loadResourceInformation() {
  const item = Office.context.mailbox.item;
  let location = item.location;
  console.log('>>>>>>>>>>>>>>>', location)
  console.log('>>>>>>>>>>>>>>> Office.context.mailbox: ', Office.context.mailbox.item.organizer)

  if (location) {
    const customField = getCustomFieldFromLocation(location);
    if (customField) {
      updateIframe(customField);
    }
  }
}

async function updateIframe(customField) {
  try {
    var el = document.createElement("iframe");
    el.src = 'https://iadbdev.service-now.com/x_nuvo_eam_microsoft_add_in.do?location=' + customField;
    el.id = 'miIframe';
    el.referrerpolicy = "strict-origin-when-cross-origin";
    document.getElementById("miIframe")?.remove();
    document.getElementById("preview").appendChild(el);
  } catch (error) {
    console.error('Error loading iframe:', error);
  }
}
