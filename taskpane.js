/*
 * This is the logic for your Signatory Request Form
 */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Wait for the document to be ready, then attach event listeners
    document.addEventListener("DOMContentLoaded", function() {
        // Get your buttons from the HTML
        const insertButton = document.getElementById("btn-insert");
        const submitButton = document.getElementById("btn-submit");

        // Assign click actions
        if (insertButton) insertButton.onclick = insertDataIntoEmail;
        if (submitButton) submitButton.onclick = submitDataToFlow;
    });
  }
});

/**
 * SCENARIO 1: Write the form data into the email body.
 */
function insertDataIntoEmail() {
  try {
    // 1. Get all values from the form
    const confirm = document.getElementById("confirm-approval").value;
    const context = document.getElementById("background-context").value;
    const cpName = document.getElementById("counterparty-name").value;
    const cpEmail = document.getElementById("counterparty-email").value;
    const xcdName = document.getElementById("xcd-name").value;
    const xcdEmail = document.getElementById("xcd-email").value;
    const xcdRep = document.getElementById("xcd-rep").value;
    const cpRep = document.getElementById("counterparty-rep").value;

    // 2. Format the text as HTML
    // Replace newline characters from the textarea with <br> tags
    const contextHtml = context.replace(/\n/g, '<br>');

    const emailBody = `
      <p>Please find the signatory request details below:</p>
      <ol>
        <li><b>Confirmation of Approval:</b> ${confirm || '<i>Not provided</i>'}</li>
        <li><b>Background / Context:</b><br>${contextHtml || '<i>Not provided</i>'}</li>
        <li><b>Counterparty Signatory Details:</b>
            <ul>
                <li>Name: ${cpName || '<i>Not provided</i>'}</li>
                <li>E-mail: ${cpEmail || '<i>Not provided</i>'}</li>
            </ul>
        </li>
        <li><b>XCD Signatory Details:</b>
            <ul>
                <li>Name: ${xcdName || '<i>Not provided</i>'}</li>
                <li>E-mail: ${xcdEmail || '<i>Not provided</i>'}</li>
            </ul>
        </li>
        <li><b>Points of Contact:</b>
            <ul>
                <li>XCD Representative: ${xcdRep || '<i>Not provided</i>'}</li>
                <li>Counterparty Representative: ${cpRep || '<i>Not provided</i>'}</li>
            </ul>
        </li>
      </ol>
    `;

    // 3. Insert it into the Outlook email body
    Office.context.mailbox.item.body.prependAsync(
      emailBody,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  } catch(e) {
    console.error(e);
  }
}

/**
 * SCENARIO 2: Send the form data to a Power Automate Flow.
 */
async function submitDataToFlow() {
  
  // !!! IMPORTANT !!!
  // 1. Go to Power Automate and create a new Flow.
  // 2. Use the trigger "When an HTTP request is received".
  // 3. Save the Flow. It will generate a "HTTP POST URL".
  // 4. Copy that URL and paste it here:
  const powerAutomateUrl = "PASTE_YOUR_POWER_AUTOMATE_URL_HERE";

  if (powerAutomateUrl === "PASTE_YOUR_POWER_AUTOMATE_URL_HERE") {
    alert("Please update the powerAutomateUrl in taskpane.js first.");
    return;
  }

  // 2. Get the values and put them into a JSON object
  const data = {
    confirmation: document.getElementById("confirm-approval").value,
    background: document.getElementById("background-context").value,
    counterparty_name: document.getElementById("counterparty-name").value,
    counterparty_email: document.getElementById("counterparty-email").value,
    xcd_name: document.getElementById("xcd-name").value,
    xcd_email: document.getElementById("xcd-email").value,
    xcd_rep: document.getElementById("xcd-rep").value,
    counterparty_rep: document.getElementById("counterparty-rep").value
  };

  // 4. Send the data to your Flow
  try {
    const response = await fetch(powerAutomateUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });

    if (response.ok) {
      alert("Request Submitted!"); // Give the user feedback
    } else {
      alert("Error submitting. Please try again.");
    }
  } catch (error) {
    console.error("Failed to send request:", error);
  }
}
