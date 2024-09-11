// Declare Office and Excel APIs
declare const Office: any;
declare const Excel: any;
declare const document: any;

// Define an interface for filters
interface Filter {
  column: string;
  values: string;
}

const EMAIL_DISCLAIMER = "This email was sent using the Excel add-in, 'teacherhelper'. Please do NOT reply to no-reply.teacherhelper@outlook.com";

// Initialize global variables
let filters: Filter[] = [];
let columnTitles: string[] = [];

let instructionsVisible = false;

// Initialize the counter
let emailCounter = 100;
let lastResetDate: string;

// At the top of your file, add these constants for original values
const ORIGINAL_INSTITUTION = "university";
const ORIGINAL_PERSONA = "professor";
const ORIGINAL_AUDIENCE = "students";
const ORIGINAL_TONE = "professional";
const ORIGINAL_MAX_TOKENS = 300;
const ORIGINAL_TEMPERATURE = 0.7;

// Your existing variables
let institution = ORIGINAL_INSTITUTION;
let persona = ORIGINAL_PERSONA;
let audience = ORIGINAL_AUDIENCE;
let tone = ORIGINAL_TONE;
let maxTokens = ORIGINAL_MAX_TOKENS;
let temperature = ORIGINAL_TEMPERATURE;

// Office.onReady function - runs when the Office APIs are ready
Office.onReady((info) => {
  console.log("Office.onReady called");
  document.getElementById("loading")!.style.display = "none";

  if (info.host === Office.HostType.Excel) {
    console.log("Excel host detected");
    document.getElementById("app-body")!.style.display = "flex";
    setupInstructionsToggle();
    initializeAdvancedSettings();
    loadCounterState();
    setupSignInToggle();
    loadSignInOptions();


    // Set up Excel-specific functionality
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // event handler for worksheet change
      sheet.onChanged.add(handleWorksheetChange);
      // event handler for worksheet addition or deletion
      context.workbook.worksheets.onActivated.add(handleWorksheetChange);

      await context.sync();
      console.log("Event handlers added for worksheet changes");

      // Initial population of column titles and menus
      await getColumnTitles();
      populateColumnTitlesMenu();
    }).catch((error) => {
      console.error("Error setting up event handlers:", error);
      showNotification("Failed to set up automatic updates. Please refresh manually if needed.", true);
    });

    // Load the saved signature
    loadSignature();

    // Load the saved sender email
    loadSenderEmail();

    // event listeners for UI elements
    document.getElementById("generate-draft")!.addEventListener("click", generateEmailDraft);
    document.getElementById("send-emails")!.addEventListener("click", sendEmails);
    document.getElementById("options")!.addEventListener("change", handleOptionsChange);
    document.getElementById("add-filter")!.addEventListener("click", addFilter);
    document.getElementById("save-signature")!.addEventListener("click", saveSignature);
    document.getElementById("save-sender-email")!.addEventListener("click", saveSenderEmail);
    document.getElementById("save-signin-options")!.addEventListener("click", saveSignInOptions);

  } else {
    console.log("Unsupported host detected:", info.host);
  }
});

// Function to get column titles from the active Excel worksheet
async function getColumnTitles(): Promise<void> {
  try {
    await Excel.run(async (context: any) => {
      console.log("Fetching column titles...");

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      const headerRow = range.getRow(0);
      headerRow.load("values");

      await context.sync();

      if (!headerRow.values || headerRow.values.length === 0 || headerRow.values[0].length === 0) {
        throw new Error("No data found in the first row of the sheet");
      }

      columnTitles = headerRow.values[0] as string[];
      console.log("Column titles updated:", columnTitles);

      // Update the email column menu
      populateEmailColumnMenu();
    });
  } catch (error) {
    console.error("Error fetching column titles:", error);
    throw error;
  }
}

// Function to populate dropdown menus with column titles
function populateColumnTitlesMenu() {
  const menus = document.getElementsByClassName("columnTitlesMenu");
  Array.from(menus).forEach((menu: HTMLSelectElement) => {
    menu.innerHTML = ""; // Clear existing options

    // Add a default option
    const defaultOption = document.createElement("option");
    defaultOption.text = "Select a column";
    defaultOption.value = "";
    menu.appendChild(defaultOption);

    // Add options for each column title
    columnTitles.forEach((title, index) => {
      const option = document.createElement("option");
      option.value = index.toString();
      option.text = title;
      menu.appendChild(option);
    });
  });
}

// Function to populate the email column dropdown menu
function populateEmailColumnMenu() {
  const emailColumnMenu = document.getElementById("email-column") as HTMLSelectElement;
  if (!emailColumnMenu) {
    console.error("Email column menu not found");
    return;
  }
  emailColumnMenu.innerHTML = ""; // Clear existing options

  // Add a default option
  const defaultOption = document.createElement("option");
  defaultOption.text = "Select a column";
  defaultOption.value = "";
  emailColumnMenu.appendChild(defaultOption);

  // Add options for each column title
  columnTitles.forEach((title, index) => {
    const option = document.createElement("option");
    option.value = index.toString();
    option.text = title;
    emailColumnMenu.appendChild(option);
  });

  // Detect and select email column
  const emailColumnIndex = detectEmailColumn(columnTitles);
  if (emailColumnIndex !== -1) {
    emailColumnMenu.value = emailColumnIndex.toString();
  }

  console.log("Email column menu populated with", columnTitles.length, "options");
  console.log("Automatically selected email column:", emailColumnIndex);
}

// Function to insert a column title into the textarea
function insertColumnTitle(textarea: HTMLTextAreaElement, title: string, position: number) {
  const before = textarea.value.substring(0, position);
  const after = textarea.value.substring(position);
  textarea.value = before + title + after;
  textarea.selectionStart = textarea.selectionEnd = position + title.length;
  textarea.focus();
}

// Function to handle changes in the options dropdown
function handleOptionsChange(event: Event) {
  const select = event.target as HTMLSelectElement;
  const addFilterButton = document.getElementById("add-filter") as HTMLButtonElement;
  addFilterButton.style.display = select.value === "specific" ? "inline-block" : "none";
  if (select.value !== "specific") {
    filters = [];
    updateFilterDisplay();
    console.log("Filters cleared:", JSON.stringify(filters));
  }
}

// Function to add a new filter
function addFilter() {
  const filter: Filter = {
    column: "",
    values: "",
  };
  const filterContainer = document.createElement("div");
  filterContainer.className = "filter-container";
  const columnSelect = createColumnSelect(filter);
  const valuesInput = createValuesInput(filter);
  const deleteButton = createDeleteButton(filter, filterContainer);
  filterContainer.appendChild(columnSelect);
  filterContainer.appendChild(valuesInput);
  filterContainer.appendChild(deleteButton);
  document.getElementById("filter-containers")!.appendChild(filterContainer);
  filters.push(filter);
  console.log("Current filters:", JSON.stringify(filters));

  // Populate the newly created column select
  populateColumnSelect(columnSelect);
}

// Function to create a column select dropdown for filters
function createColumnSelect(filter: Filter): HTMLSelectElement {
  const columnSelect = document.createElement("select");
  columnSelect.className = "columnTitlesMenu";
  columnSelect.addEventListener("change", function (this: HTMLSelectElement) {
    filter.column = columnTitles[parseInt(this.value)];
  });
  return columnSelect;
}

// Function to populate a column select dropdown
function populateColumnSelect(columnSelect: HTMLSelectElement) {
  columnSelect.innerHTML = ""; // Clear existing options

  // Add a default option
  const defaultOption = document.createElement("option");
  defaultOption.text = "Select a column";
  defaultOption.value = "";
  columnSelect.appendChild(defaultOption);

  // Add options for each column title
  columnTitles.forEach((title, index) => {
    const option = document.createElement("option");
    option.value = index.toString();
    option.text = title;
    columnSelect.appendChild(option);
  });
}

// Function to create an input field for filter values
function createValuesInput(filter: Filter): HTMLInputElement {
  const valuesInput = document.createElement("input");
  valuesInput.type = "text";
  valuesInput.placeholder = "Enter values (comma-separated)";
  valuesInput.addEventListener("input", function (this: HTMLInputElement) {
    filter.values = this.value;
  });
  return valuesInput;
}

// Function to create a delete button for filters
function createDeleteButton(filter: Filter, filterContainer: HTMLDivElement): HTMLButtonElement {
  const deleteButton = document.createElement("button");
  deleteButton.textContent = "Delete";
  deleteButton.addEventListener("click", function () {
    filterContainer.remove();
    const index = filters.findIndex((f) => f === filter);
    if (index !== -1) {
      filters.splice(index, 1);
    }
  });
  return deleteButton;
}

// Function to update the filter display
function updateFilterDisplay() {
  const filterContainers = document.getElementById("filter-containers")!;
  filterContainers.innerHTML = "";
}

// Function to generate an email draft using the Claude API
async function generateEmailDraft() {
  const promptElement = document.getElementById("myTextarea") as HTMLTextAreaElement;
  const subjectElement = document.getElementById("email-subject") as HTMLTextAreaElement;
  const bodyElement = document.getElementById("email-body") as HTMLTextAreaElement;
  const prompt = promptElement.value;

  if (!prompt) {
    showNotification("Please enter a prompt before generating an email draft.", true);
    return;
  }

  try {
    console.log("Generating email draft...");
    showNotification("Generating email draft...");
    const generatedText = await callClaudeAPI(prompt, columnTitles);
    const { subject, body } = parseGeneratedEmail(generatedText);
    subjectElement.value = subject;
    bodyElement.value = body;

    // Append the signature and disclaimer to the email body
    const signatureElement = document.getElementById("email-signature") as HTMLTextAreaElement;
    const signature = signatureElement.value;
    bodyElement.value = body + "\n\n" + signature + "\n\n" + `\n\n${EMAIL_DISCLAIMER}`;

    showNotification("Email draft generated successfully!");
  } catch (error) {
    console.error("Error in generateEmailDraft:", error);
    subjectElement.value = "Error generating subject";
    bodyElement.value = "An error occurred. Please try again. Error details: " + error.message;
    showNotification("Failed to generate email draft. Please try again.", true);
  }
}

function parseGeneratedEmail(text: string): { subject: string; body: string } {
  const parts = text.split('\n\n');
  return {
    subject: parts[0].replace('Subject: ', '').trim(),
    body: parts.slice(1).join('\n\n').trim()
  };
}

// Function to connect with Claude API
async function callClaudeAPI(prompt: string, columnTitles: string[]): Promise<string> {
  const apiUrl = "http://localhost:3001/api/generate";

  const ANTHROPIC_API_KEY = (document.getElementById("anthropic-api") as HTMLInputElement).value;

  // Define new variables with default values
  const institution = "university";
  const persona = "professor";
  const audience = "students";
  const tone = "professional";

  if (!ANTHROPIC_API_KEY) {
    throw new Error("Anthropic API key is not set. Please enter it in the Sign In section.");
  }

  try {
    console.log("Calling Claude API...");
    console.log("Prompt:", prompt);
    console.log("Column titles:", columnTitles);
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        prompt: `Human: 
                 Forget any previous instructions.
                 You are an AI assistant helping a ${persona} at a ${institution} write an ${tone} email to to ${audience}. 
                 The ${persona} will provide instructions, and you should write a ${tone} email based on those instructions. 
                 Generate both a subject line and an email body. 
                 Use {{column_title}} as placeholders for personalised information.
                 Only use curly brackets {} This is the only type of brackets you are allowed to use.
                 Available column titles are: ${columnTitles.join(", ")}. 
                 Only stick to the available column titles in the provided menu. 
                 Do not repeat the column titles multiple times if you are creating a list within the generated email body. 
                 This is very important: only provide the email draft ready to be sent instead of writing something along the lines of "Here is an email..." at the beginning. 
                 Do not use a signature for the ${persona}. 
                 Stop generating the email body after you write "kind regards" or "sincerely" or the other similar words that are used to end emails. 
                 Provide the email draft in the following format:

                Subject: [Generated Subject]

                [Generated Email Body]

                Here are the teacher's instructions: "${prompt}"

                Assistant:`,
        max_tokens_to_sample: maxTokens,
        temperature: temperature,
        ANTHROPIC_API_KEY: ANTHROPIC_API_KEY
      }),
    });

    const data = await response.json();
    console.log("API response:", JSON.stringify(data, null, 2));

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}, message: ${JSON.stringify(data, null, 2)}`);
    }

    if (!data.completion) {
      throw new Error("No completion in API response: " + JSON.stringify(data, null, 2));
    }

    return data.completion;
  } catch (error) {
    console.error("Error calling Claude API:", error);
    throw error;
  }
}

// Function to send the generated emails
async function sendEmails() {
  console.log("Sending emails...");

  const SENDGRID_API_KEY = (document.getElementById("sendgrid-api") as HTMLInputElement).value;
  const sendgridEmail = (document.getElementById("sendgrid-email") as HTMLInputElement).value;

  if (!SENDGRID_API_KEY || !sendgridEmail) {
    showNotification("SendGrid API key and email are required. Please enter them in the Sign In section.", true);
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const emailColumnSelect = document.getElementById("email-column") as HTMLSelectElement;
      const emailColumnIndex = parseInt(emailColumnSelect.value);

      console.log(`Selected email column index: ${emailColumnIndex}`);

      if (isNaN(emailColumnIndex)) {
        throw new Error("Please select an email column");
      }

      const usedRange = sheet.getUsedRange();
      usedRange.load("values");
      await context.sync();

      const headerRow = usedRange.values[0] as string[];
      console.log(`Header row: ${JSON.stringify(headerRow)}`);

      const subjectElement = document.getElementById("email-subject") as HTMLTextAreaElement;
      const bodyElement = document.getElementById("email-body") as HTMLTextAreaElement;
      const senderEmailElement = document.getElementById("sender-email") as HTMLTextAreaElement;
      const senderEmail = senderEmailElement.value.trim();
      const signatureElement = document.getElementById("email-signature") as HTMLTextAreaElement;

      if (!subjectElement || !bodyElement) {
        throw new Error("Email subject or body element not found");
      }

      const emailSubjectTemplate = subjectElement.value;
      const emailBodyTemplate = bodyElement.value;

      console.log(`Email subject template: ${emailSubjectTemplate}`);
      console.log(`Email body template: ${emailBodyTemplate}`);

      if (!emailSubjectTemplate || !emailBodyTemplate) {
        throw new Error("Email subject or body template is empty");
      }

      const emails = [];

      let skippedRows = 0;
      let totalRows = usedRange.values.length - 1; // Subtract 1 for header row
      let filteredOutRows = 0;

      console.log(`Processing ${totalRows} rows...`);

      const optionsSelect = document.getElementById("options") as HTMLSelectElement;
      const useFilters = optionsSelect.value === "specific";
      console.log(`Using filters: ${useFilters}`);
      console.log(`Current filters: ${JSON.stringify(filters)}`);

      for (let i = 1; i < usedRange.values.length; i++) {
        const row = usedRange.values[i];
        console.log(`Processing row ${i}: ${JSON.stringify(row)}`);

        const emailAddress = row[emailColumnIndex];
        console.log(`Email address found: ${emailAddress}`);

        if (!emailAddress || !isValidEmail(emailAddress)) {
          console.log(`Skipping row ${i}: Invalid email address`);
          skippedRows++;
          continue;
        }

        if (useFilters) {
          const passesFilter = applyFilters(row, headerRow);
          console.log(`Row ${i} passes filter: ${passesFilter}`);
          if (!passesFilter) {
            filteredOutRows++;
            continue;
          }
        }

        // Check if we have enough email credits (include sender email if provided)
        const totalEmailsToSend = emails.length + (senderEmail ? 1 : 0);
        if (totalEmailsToSend > emailCounter) {
          showWarningModal(`Warning: The number of emails to be sent (${totalEmailsToSend}) exceeds your remaining email sending credit (${emailCounter}). Please reduce the number of recipients or try again tomorrow.`);
          return;
        }

        // If sender email is provided and valid, add it to the emails array
        if (senderEmail && isValidEmail(senderEmail)) {
          emails.push({
            to: senderEmail,
            subject: subjectElement.value,
            html: formatEmailContent(bodyElement.value),
          });
        }

        let personalizedSubject = emailSubjectTemplate;
        let personalizedBody = emailBodyTemplate;

        for (let j = 0; j < row.length; j++) {
          const columnName = headerRow[j];
          const cellValue = row[j];
          const placeholder = `{{${columnName}}}`;
          personalizedSubject = personalizedSubject.replace(new RegExp(placeholder, 'g'), cellValue);
          personalizedBody = personalizedBody.replace(new RegExp(placeholder, 'g'), cellValue);
        }

        personalizedBody = formatEmailContent(personalizedBody);

        emails.push({
          to: emailAddress,
          subject: personalizedSubject,
          html: personalizedBody,
        });
        console.log(`Added email for ${emailAddress} with subject: ${personalizedSubject}`);
      }

      console.log(`Total rows processed: ${totalRows}`);
      console.log(`Skipped rows (invalid email): ${skippedRows}`);
      console.log(`Filtered out rows: ${filteredOutRows}`);
      console.log(`Valid emails found: ${emails.length}`);

      if (emails.length === 0) {
        throw new Error("No valid emails to send. Please check your data and filter criteria.");
      }

      console.log("Prepared emails:", JSON.stringify(emails, null, 2));

      console.log(`Preparing to send ${emails.length} emails`);
      const response = await fetch("http://localhost:3001/api/send-emails", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          emails,
          SENDGRID_API_KEY,
          sendgridEmail
        }),
      });

      const responseData = await response.json();
      console.log("Server response:", JSON.stringify(responseData, null, 2));

      if (!response.ok) {
        throw new Error(
          `HTTP error! status: ${response.status}, message: ${responseData.error}, details: ${JSON.stringify(responseData, null, 2)}`
        );
      }

      if (responseData.sentEmails && Array.isArray(responseData.sentEmails)) {
        console.log(`Successfully sent ${responseData.sentEmails.length} out of ${emails.length} emails`);
        responseData.sentEmails.forEach((email: string) => {
          console.log(`Email sent to: ${email}`);
          decreaseCounter();
        });
      }

      if (responseData.failedEmails && Array.isArray(responseData.failedEmails)) {
        console.log(`Failed to send ${responseData.failedEmails.length} emails`);
        responseData.failedEmails.forEach((failedEmail: { email: string, error: string }) => {
          console.log(`Failed to send email to ${failedEmail.email}. Error: ${failedEmail.error}`);
        });
      }

      showNotification(`Emails sent: ${responseData.sentEmails.length}/${emails.length}. Check console for details.`);

    });
  } catch (error) {
    console.error("Error sending emails:", error);
    showNotification(`Error sending emails: ${error.message}`, true);
  }
}

// Function to format  the email content
function formatEmailContent(content: string): string {
  // Replace single newlines with <br> tags
  content = content.replace(/(?<!\n)\n(?!\n)/g, "<br>");

  // Replace double newlines with paragraph breaks
  content = content.replace(/\n\n/g, "</p><p>");

  // Wrap the entire content in a paragraph tag
  content = `<p>${content}</p>`;

  return content;
}

// Function to check email validity within a column
function isValidEmail(email: string) {
  const re =
    /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(String(email).toLowerCase());
}

// Function to display notifications
function showNotification(message: string, isError: boolean = false) {
  const notificationElement = document.createElement("div");
  notificationElement.textContent = message;
  notificationElement.style.position = "fixed";
  notificationElement.style.top = "10px";
  notificationElement.style.left = "50%";
  notificationElement.style.transform = "translateX(-50%)";
  notificationElement.style.padding = "10px";
  notificationElement.style.borderRadius = "5px";
  notificationElement.style.backgroundColor = isError ? "#ffcccc" : "#ccffcc";
  notificationElement.style.border = `1px solid ${isError ? "#ff0000" : "#00ff00"}`;
  document.body.appendChild(notificationElement);

  setTimeout(() => {
    document.body.removeChild(notificationElement);
  }, 3000);
}

// Function to apply created filters
function applyFilters(row: any[], headerRow: string[]): boolean {
  if (filters.length === 0) return true;

  console.log("Applying filters:", JSON.stringify(filters));

  return filters.every(filter => {
    const columnIndex = headerRow.findIndex(col => col.toLowerCase() === filter.column.toLowerCase());
    if (columnIndex === -1) {
      console.log(`Column not found: ${filter.column}`);
      return false;
    }

    const cellValue = row[columnIndex]?.toString().toLowerCase() ?? "";
    const filterValues = filter.values.toLowerCase().split(',').map(v => v.trim());

    console.log(`Checking column: ${filter.column}, Cell value: ${cellValue}, Filter values: ${filterValues.join(', ')}`);

    const result = filterValues.some(value => cellValue.includes(value));
    console.log(`Filter result for ${filter.column}: ${result}`);
  return result;
  });
}

// Function to automatically detect email columns
function detectEmailColumn(headerRow: string[]): number {
  const emailRegex = /email|e-mail|mail/i;
  return headerRow.findIndex(header => emailRegex.test(header));
}

// Function to automatically register changes in the worksheet
async function handleWorksheetChange(event: Office.EventType) {
  try {
    console.log("Worksheet changed. Updating column titles and email menu.");
    await getColumnTitles();
    populateColumnTitlesMenu();
  } catch (error) {
    console.error("Error handling worksheet change:", error);
    showNotification("Failed to update columns. Please try again.", true);
  }
}

// Function to save the teacher's signature
function saveSignature() {
  const signatureElement = document.getElementById("email-signature") as HTMLTextAreaElement;
  const signature = signatureElement.value;
  Office.context.document.settings.set("teacherSignature", signature);
  Office.context.document.settings.saveAsync(() => {
    console.log("Signature saved successfully");
    showNotification("Signature saved successfully!");
  });
}

// Function to save the loaded signature
function loadSignature() {
  const signature = Office.context.document.settings.get("teacherSignature");
  if (signature) {
    const signatureElement = document.getElementById("email-signature") as HTMLTextAreaElement;
    signatureElement.value = signature;
  }
}

function setupInstructionsToggle() {
  const toggleButton = document.getElementById("toggle-instructions");
  const instructions = document.getElementById("instructions");

  if (toggleButton && instructions) {
    toggleButton.onclick = function() {
      instructionsVisible = !instructionsVisible;
      instructions.style.display = instructionsVisible ? "block" : "none";
      (toggleButton as HTMLButtonElement).textContent = instructionsVisible ? "Hide Instructions" : "Show instructions on how to use teacherhelper";
    }
  }
}

// Function to initialize advanced settings
function initializeAdvancedSettings() {
  console.log("Initializing advanced settings");
  const advancedSettingsBtn = document.getElementById('advanced-settings-btn') as HTMLButtonElement;
  const advancedSettings = document.getElementById('advanced-settings') as HTMLDivElement;
  const resetButton = document.getElementById('advanced-settings-reset') as HTMLButtonElement;
  const saveButton = document.getElementById('advanced-settings-save') as HTMLButtonElement;

  if (!advancedSettingsBtn || !advancedSettings || !resetButton || !saveButton) {
    console.error("One or more advanced settings elements not found");
    return;
  }

  advancedSettingsBtn.addEventListener('click', () => {
    advancedSettings.style.display = advancedSettings.style.display === 'none' ? 'block' : 'none';
  });

  resetButton.addEventListener('click', resetToOriginalValues);
  saveButton.addEventListener('click', saveControlSettings);

  // Load saved settings or use defaults
  loadSavedSettings();

  // Initialize input fields
  updateInputFields();

  // Add event listeners for input changes
  document.getElementById('institution')?.addEventListener('input', (e) => {
    institution = (e.target as HTMLInputElement).value || ORIGINAL_INSTITUTION;
  });

  document.getElementById('persona')?.addEventListener('input', (e) => {
    persona = (e.target as HTMLInputElement).value || ORIGINAL_PERSONA;
  });

  document.getElementById('audience')?.addEventListener('input', (e) => {
    audience = (e.target as HTMLInputElement).value || ORIGINAL_AUDIENCE;
  });

  document.getElementById('tone')?.addEventListener('input', (e) => {
    tone = (e.target as HTMLInputElement).value || ORIGINAL_TONE;
  });

  document.getElementById('max-tokens')?.addEventListener('input', (e) => {
    maxTokens = parseInt((e.target as HTMLInputElement).value);
    (document.getElementById('max-tokens-value') as HTMLSpanElement).textContent = maxTokens.toString();
  });

  document.getElementById('temperature')?.addEventListener('input', (e) => {
    temperature = parseFloat((e.target as HTMLInputElement).value);
    (document.getElementById('temperature-value') as HTMLSpanElement).textContent = temperature.toString();
  });
}

function resetToOriginalValues() {
  institution = ORIGINAL_INSTITUTION;
  persona = ORIGINAL_PERSONA;
  audience = ORIGINAL_AUDIENCE;
  tone = ORIGINAL_TONE;
  maxTokens = ORIGINAL_MAX_TOKENS;
  temperature = ORIGINAL_TEMPERATURE;

  updateInputFields();
  showNotification("Settings reset to original values.");
}

function saveControlSettings() {
  Office.context.document.settings.set("teacherhelper_institution", institution);
  Office.context.document.settings.set("teacherhelper_persona", persona);
  Office.context.document.settings.set("teacherhelper_audience", audience);
  Office.context.document.settings.set("teacherhelper_tone", tone);
  Office.context.document.settings.set("teacherhelper_maxTokens", maxTokens);
  Office.context.document.settings.set("teacherhelper_temperature", temperature);

  Office.context.document.settings.saveAsync(() => {
    console.log("Control settings saved successfully");
    showNotification("Control settings saved successfully!");
  });
}

function loadSavedSettings() {
  institution = Office.context.document.settings.get("teacherhelper_institution") || ORIGINAL_INSTITUTION;
  persona = Office.context.document.settings.get("teacherhelper_persona") || ORIGINAL_PERSONA;
  audience = Office.context.document.settings.get("teacherhelper_audience") || ORIGINAL_AUDIENCE;
  tone = Office.context.document.settings.get("teacherhelper_tone") || ORIGINAL_TONE;
  maxTokens = Office.context.document.settings.get("teacherhelper_maxTokens") || ORIGINAL_MAX_TOKENS;
  temperature = Office.context.document.settings.get("teacherhelper_temperature") || ORIGINAL_TEMPERATURE;
}

function updateInputFields() {
  (document.getElementById('institution') as HTMLInputElement).value = institution;
  (document.getElementById('persona') as HTMLInputElement).value = persona;
  (document.getElementById('audience') as HTMLInputElement).value = audience;
  (document.getElementById('tone') as HTMLInputElement).value = tone;
  (document.getElementById('max-tokens') as HTMLInputElement).value = maxTokens.toString();
  (document.getElementById('temperature') as HTMLInputElement).value = temperature.toString();
  (document.getElementById('max-tokens-value') as HTMLSpanElement).textContent = maxTokens.toString();
  (document.getElementById('temperature-value') as HTMLSpanElement).textContent = temperature.toString();
}

// Function to update the counter display
function updateCounterDisplay() {
  const counterElement = document.getElementById('email-counter');
  if (counterElement) {
    counterElement.textContent = `Emails remaining today: ${emailCounter}`;
  }
}

// Function to decrease the counter
function decreaseCounter() {
  if (emailCounter > 0) {
    emailCounter--;
    updateCounterDisplay();
    saveCounterState();
  }
}

// Function to check and reset the counter if it's a new day
function checkAndResetCounter() {
  const today = new Date().toDateString();

  if (lastResetDate !== today) {
    emailCounter = 100;
    lastResetDate = today;
    saveCounterState();
  }

  updateCounterDisplay();
}

// Function to save the counter state to local storage
function saveCounterState() {
  localStorage.setItem('emailCounter', emailCounter.toString());
  localStorage.setItem('lastResetDate', lastResetDate);
}

// Function to load the counter state from local storage
function loadCounterState() {
  const savedCounter = localStorage.getItem('emailCounter');
  const savedDate = localStorage.getItem('lastResetDate');

  if (savedCounter !== null) {
    emailCounter = parseInt(savedCounter, 10);
  }

  if (savedDate !== null) {
    lastResetDate = savedDate;
  } else {
    lastResetDate = new Date().toDateString();
  }

  checkAndResetCounter();
}

// Function to show a modal warning
function showWarningModal(message: string) {
  // Create modal elements
  const modal = document.createElement('div');
  modal.style.cssText = `
    position: fixed;
    z-index: 1000;
    left: 0;
    top: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0,0,0,0.4);
    display: flex;
    justify-content: center;
    align-items: center;
  `;

  const modalContent = document.createElement('div');
  modalContent.style.cssText = `
    background-color: #fefefe;
    padding: 20px;
    border: 1px solid #888;
    width: 80%;
    max-width: 400px;
    text-align: center;
  `;

  const closeBtn = document.createElement('button');
  closeBtn.textContent = 'Close';
  closeBtn.onclick = () => document.body.removeChild(modal);

  modalContent.innerHTML = `<p>${message}</p>`;
  modalContent.appendChild(closeBtn);
  modal.appendChild(modalContent);

  document.body.appendChild(modal);
}

function saveSenderEmail() {
  const senderEmailElement = document.getElementById("sender-email") as HTMLTextAreaElement;
  const senderEmail = senderEmailElement.value.trim();

  if (senderEmail === "" || isValidEmail(senderEmail)) {
    Office.context.document.settings.set("senderEmail", senderEmail);
    Office.context.document.settings.saveAsync(() => {
      console.log("Sender email saved successfully");
      if (senderEmail === "") {
        showNotification("Your email has been cleared successfully!");
      } else {
        showNotification("Your email has been saved successfully!");
      }
    });
  } else {
    showNotification("Please enter a valid email address or leave it empty to clear.", true);
  }
}

function loadSenderEmail() {
  const senderEmail = Office.context.document.settings.get("senderEmail");
  if (senderEmail) {
    const senderEmailElement = document.getElementById("sender-email") as HTMLTextAreaElement;
    senderEmailElement.value = senderEmail;
  }
}

let signInOptionsVisible = false;

function setupSignInToggle() {
  const toggleButton = document.getElementById("toggle-signin");
  const signInOptions = document.getElementById("signin-options");

  if (toggleButton && signInOptions) {
    toggleButton.onclick = function() {
      signInOptionsVisible = !signInOptionsVisible;
      signInOptions.style.display = signInOptionsVisible ? "block" : "none";
    }
  }
}

function saveSignInOptions() {
  const ANTHROPIC_API_KEY = (document.getElementById("anthropic-api") as HTMLInputElement).value;
  const sendgridEmail = (document.getElementById("sendgrid-email") as HTMLInputElement).value;
  const SENDGRID_API_KEY = (document.getElementById("sendgrid-api") as HTMLInputElement).value;

  Office.context.document.settings.set("ANTHROPIC_API_KEY", ANTHROPIC_API_KEY);
  Office.context.document.settings.set("sendgridEmail", sendgridEmail);
  Office.context.document.settings.set("SENDGRID_API_KEY", SENDGRID_API_KEY);

  Office.context.document.settings.saveAsync(() => {
    console.log("Sign-in options saved successfully");
    showNotification("Sign-in options saved successfully!");
  });
}

function loadSignInOptions() {
  const ANTHROPIC_API_KEY = Office.context.document.settings.get("ANTHROPIC_API_KEY");
  const sendgridEmail = Office.context.document.settings.get("sendgridEmail");
  const SENDGRID_API_KEY = Office.context.document.settings.get("SENDGRID_API_KEY");

  if (ANTHROPIC_API_KEY) {
    (document.getElementById("anthropic-api") as HTMLInputElement).value = ANTHROPIC_API_KEY;
  }
  if (sendgridEmail) {
    (document.getElementById("sendgrid-email") as HTMLInputElement).value = sendgridEmail;
  }
  if (SENDGRID_API_KEY) {
    (document.getElementById("sendgrid-api") as HTMLInputElement).value = SENDGRID_API_KEY;
  }
}