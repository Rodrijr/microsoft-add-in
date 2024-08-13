Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initialize();
  }
});

function initialize() {
  if (Office.context.mailbox.item) {
    monitorLocationChange();
  }
}

function monitorLocationChange() {
  const item = Office.context.mailbox.item;

  // Get the initial location
  const locationValue = item.location.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      document.getElementById("locationValue").textContent = result.value || "None";
    }
  });

  // Monitor for changes in the location field
  item.location.addAsyncHandler((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      document.getElementById("locationValue").textContent = result.value || "None";
    }
  });
}
