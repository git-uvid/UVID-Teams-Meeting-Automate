Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Ensure we are in a message read context
    const item = Office.context.mailbox.item;
    if (item) {
        displayEmailProperties(item);
    }
  }
});

function displayEmailProperties(item) {
  // 1. Get Item ID
  const itemId = item.itemId;
  document.getElementById("item-id").innerText = itemId || "Not available";
  
  // 2. Get Subject
  const subject = item.subject;
  document.getElementById("subject").innerText = subject || "No subject";
  
  // 3. Get Internet Message ID
  const internetMessageId = item.internetMessageId;
  document.getElementById("internet-message-id").innerText = internetMessageId || "Not available";
  
  // 4. Get From details (Display Name and Email Address)
  const from = item.from;
  if (from) {
    document.getElementById("from-name").innerText = from.displayName || "Unknown";
    document.getElementById("from-email").innerText = from.emailAddress || "Unknown";
  } else {
    document.getElementById("from-name").innerText = "Not available";
    document.getElementById("from-email").innerText = "Not available";
  }
}
