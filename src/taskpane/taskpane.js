function logMessageToUI(message) {
  const timestamp = new Date().toISOString();
  const fullMessage = `${timestamp} - ${message}`;

  let logElement = document.getElementById('log');
  if (!logElement) {
    logElement = document.createElement('pre');
    logElement.id = 'log';
    document.body.appendChild(logElement);
  }
  logElement.textContent += fullMessage + '\n';
}

// Verify the add-in is ready and log the host information.
Office.onReady((info) => {
  logMessageToUI('Outlook Add-in is ready!');
  logMessageToUI(`Host: ${info.host}`);
  logMessageToUI('Attempting to fetch email information...');

  // Call getEmailInfo once the add-in is ready.
  getEmailInfo();
});

// Function to fetch the selected email’s information.
async function getEmailInfo() {
  try {
    const item = Office.context.mailbox.item;

    if (!item) {
      logMessageToUI('No email item found. Make sure you open an email and launch the add-in.');
      alert('No email selected. Open an email and try again.');
      return;
    }

    const subject = item.subject || 'No Subject';
    const messageId = item.itemId || 'No Message ID';

    logMessageToUI(`Email Subject: ${subject}`);
    logMessageToUI(`Email Message ID: ${messageId}`);

    // Display the subject in the UI.
    document.getElementById('subject').textContent = `Subject: ${subject}`;

    // Set the click handler for sending the email details to OmniFocus.
    document.getElementById('sendToOmniFocus').onclick = () =>
      sendToOmniFocus(subject, messageId);
  } catch (error) {
    logMessageToUI(`Error fetching email information: ${error.message}`);
    alert('Error fetching email information. Check the logs.');
  }
}

// Function to send task details to OmniFocus.
function sendToOmniFocus(subject, messageId) {
  const emailLink = `outlook://message/${messageId}`;

  logMessageToUI(`Creating OmniFocus task for: ${subject}`);
  logMessageToUI(`Email link: ${emailLink}`);

  alert(`Creating OmniFocus task...\nSubject: ${subject}\nLink: ${emailLink}`);

  fetch('http://localhost:5000/omnifocus', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ task: subject, note: `Link to email: ${emailLink}` }),
  })
    .then((response) => response.json())
    .then((data) => {
      logMessageToUI(`Task sent to OmniFocus: ${JSON.stringify(data)}`);
    })
    .catch((error) => {
      logMessageToUI(`Error sending task to OmniFocus: ${error.message}`);
    });
}
