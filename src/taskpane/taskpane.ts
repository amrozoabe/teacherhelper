declare const Office: any;
declare const Excel: any;
declare const document: any;

interface Filter {
  column: string;
  values: string;
}

export class Logging {
  static log(...args: any[]) {
    // eslint-disable-next-line
    console.log(args);
  }

  static error(...args: any[]) {
    // eslint-disable-next-line
    console.log(args);
  }
}

Logging.log("Amro is awesome hello")

let filters: Filter[] = [];
let columnTitles: string[] = [];

Office.onReady((info) => {
  Logging.log("Office.onReady called");
  document.getElementById("loading")!.style.display = "none";

  if (info.host === Office.HostType.Excel) {
    Logging.log("Excel host detected");
    document.getElementById("app-body")!.style.display = "flex";

    getColumnTitles()
      .then(() => {
        Logging.log("Column titles initialized");
        if (columnTitles.length === 0) {
          showNotification("No column titles found. Please ensure your Excel sheet has data in the first row.", true);
          return;
        }
        populateColumnTitlesMenu();
        populateEmailColumnMenu();
      })
      .catch((error) => {
        Logging.error("Failed to initialize column titles:", error);
        showNotification(
          "Failed to initialize column titles. Please ensure your Excel sheet has data and try again.",
          true
        );
      });

    document.getElementById("generate-draft")!.addEventListener("click", generateEmailDraft);
    document.getElementById("send-emails")!.addEventListener("click", sendEmails);
    document.getElementById("options")!.addEventListener("change", handleOptionsChange);
    document.getElementById("add-filter")!.addEventListener("click", addFilter);
    document.getElementById("myTextarea")!.addEventListener("input", handlePromptInput);

    document.getElementById("refresh-columns")!.addEventListener("click", async () => {
      try {
        await getColumnTitles();
        populateColumnTitlesMenu();
        populateEmailColumnMenu();
        showNotification("Columns refreshed successfully!");
      } catch (error) {
        Logging.error("Failed to refresh columns:", error);
        showNotification("Failed to refresh columns. Please try again.", true);
      }
    });
  } else if (info.host === Office.HostType.Outlook) {
    Logging.log("Outlook host detected");
    document.getElementById("app-body")!.style.display = "flex";
    // Add any Outlook-specific initialization here
  } else {
    Logging.log("Unsupported host detected:", info.host);
  }
});

async function getColumnTitles(): Promise<void> {
  const maxRetries = 3;
  for (let i = 0; i < maxRetries; i++) {
    try {
      await Excel.run(async (context: any) => {
        Logging.log("Attempting to fetch column titles...");

        const application = context.workbook.application;
        application.load("suspendScreenUpdatingUntilNextSync");
        await context.sync();

        if (application.suspendScreenUpdatingUntilNextSync) {
          Logging.log(`Excel is in editing mode. Attempting to exit editing mode...`);
          context.runtime.load("enableEvents");
          await context.sync();
          context.runtime.enableEvents = true;
          await context.sync();
        }

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getUsedRange();
        const headerRow = range.getRow(0);
        headerRow.load("values");

        await context.sync();

        if (!headerRow.values || headerRow.values.length === 0 || headerRow.values[0].length === 0) {
          throw new Error("No data found in the first row of the sheet");
        }

        columnTitles = headerRow.values[0] as string[];
        Logging.log("Column titles populated:", columnTitles);
        return; // Successfully fetched titles, exit the retry loop
      });
      break; // If we reach here, we've successfully fetched the titles
    } catch (error) {
      Logging.error(`Error fetching column titles (attempt ${i + 1}):`, error);
      if (i === maxRetries - 1) {
        throw error; // Throw error on last attempt
      }
      await new Promise((resolve) => setTimeout(resolve, 1000)); // Wait for 1 second before retrying
    }
  }
}

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

function populateEmailColumnMenu() {
  const emailColumnMenu = document.getElementById("email-column") as HTMLSelectElement;
  if (!emailColumnMenu) {
    Logging.error("Email column menu not found");
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

  Logging.log("Email column menu populated with", columnTitles.length, "options");
}

function handlePromptInput(event: Event) {
  const textarea = event.target as HTMLTextAreaElement;
  const cursorPosition = textarea.selectionStart;
  const textBeforeCursor = textarea.value.substring(0, cursorPosition);

  if (textBeforeCursor.endsWith("@")) {
    showColumnTitlesMenu(textarea, cursorPosition);
  }
}

function showColumnTitlesMenu(textarea: HTMLTextAreaElement, position: number) {
  // Remove any existing dropdown
  const existingDropdown = document.querySelector(".column-titles-dropdown");
  if (existingDropdown) {
    existingDropdown.remove();
  }

  // Create a dropdown element
  const dropdown = document.createElement("div");
  dropdown.className = "column-titles-dropdown";
  dropdown.style.position = "fixed"; // Change to fixed positioning
  dropdown.style.zIndex = "1000";
  dropdown.style.backgroundColor = "white";
  dropdown.style.border = "1px solid #ccc";
  dropdown.style.maxHeight = "200px";
  dropdown.style.overflowY = "auto";
  dropdown.style.boxShadow = "0 2px 5px rgba(0,0,0,0.2)";

  // Position the dropdown near the cursor
  const rect = textarea.getBoundingClientRect();
  const { top, left } = getCaretCoordinates(textarea, position);

  const dropdownTop = rect.top + top - dropdown.offsetHeight;
  const dropdownLeft = rect.left + left;

  dropdown.style.left = `${dropdownLeft}px`;
  dropdown.style.top = `${dropdownTop}px`;

  // Populate the dropdown with column titles
  columnTitles.forEach((title, index) => {
    const item = document.createElement("div");
    item.textContent = title;
    item.className = "column-title-item";
    item.style.padding = "5px";
    item.style.cursor = "pointer";
    item.onmouseover = () => {
      item.style.backgroundColor = "#f0f0f0";
    };
    item.onmouseout = () => {
      item.style.backgroundColor = "white";
    };
    item.onclick = () => {
      insertColumnTitle(textarea, title, position);
      dropdown.remove();
    };
    dropdown.appendChild(item);
  });

  // Add the dropdown to the body
  document.body.appendChild(dropdown);

  // Adjust the position after adding to the DOM
  const dropdownRect = dropdown.getBoundingClientRect();
  if (dropdownRect.top < 0) {
    dropdown.style.top = `${rect.top + top + 20}px`; // Position below if not enough space above
  }

  // Close the dropdown when clicking outside
  document.addEventListener("click", function closeDropdown(e: MouseEvent) {
    if (!dropdown.contains(e.target as Node) && e.target !== textarea) {
      dropdown.remove();
      document.removeEventListener("click", closeDropdown);
    }
  });
}

function getCaretCoordinates(element: HTMLTextAreaElement, position: number) {
  const div = document.createElement("div");
  const styles = getComputedStyle(element);
  const properties = [
    "fontFamily",
    "fontSize",
    "fontWeight",
    "fontStyle",
    "letterSpacing",
    "textTransform",
    "wordSpacing",
    "textIndent",
    "whiteSpace",
    "lineHeight",
    "padding",
    "border",
    "boxSizing",
  ];

  properties.forEach((prop) => {
    div.style[prop] = styles[prop];
  });

  div.textContent = element.value.substring(0, position);
  div.style.position = "absolute";
  div.style.visibility = "hidden";
  div.style.whiteSpace = "pre-wrap";

  document.body.appendChild(div);
  const coordinates = {
    top: div.offsetHeight - element.scrollTop,
    left: div.offsetWidth - element.scrollLeft,
  };
  document.body.removeChild(div);

  return coordinates;
}

const properties = [
  "direction",
  "boxSizing",
  "width",
  "height",
  "overflowX",
  "overflowY",
  "borderTopWidth",
  "borderRightWidth",
  "borderBottomWidth",
  "borderLeftWidth",
  "paddingTop",
  "paddingRight",
  "paddingBottom",
  "paddingLeft",
  "fontStyle",
  "fontVariant",
  "fontWeight",
  "fontStretch",
  "fontSize",
  "fontSizeAdjust",
  "lineHeight",
  "fontFamily",
  "textAlign",
  "textTransform",
  "textIndent",
  "textDecoration",
  "letterSpacing",
  "wordSpacing",
];

function insertColumnTitle(textarea: HTMLTextAreaElement, title: string, position: number) {
  const before = textarea.value.substring(0, position);
  const after = textarea.value.substring(position);
  textarea.value = before + title + after;
  textarea.selectionStart = textarea.selectionEnd = position + title.length;
  textarea.focus();
}

function handleOptionsChange(event: Event) {
  const select = event.target as HTMLSelectElement;
  const addFilterButton = document.getElementById("add-filter") as HTMLButtonElement;
  addFilterButton.style.display = select.value === "specific" ? "inline-block" : "none";
  if (select.value !== "specific") {
    filters = [];
    updateFilterDisplay();
  }
}

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

  // Populate the newly created column select
  populateColumnSelect(columnSelect);
}

function createColumnSelect(filter: Filter): HTMLSelectElement {
  const columnSelect = document.createElement("select");
  columnSelect.className = "columnTitlesMenu";
  columnSelect.addEventListener("change", function (this: HTMLSelectElement) {
    filter.column = this.value;
  });
  return columnSelect;
}

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

function createValuesInput(filter: Filter): HTMLInputElement {
  const valuesInput = document.createElement("input");
  valuesInput.type = "text";
  valuesInput.placeholder = "Enter values (comma-separated)";
  valuesInput.addEventListener("input", function (this: HTMLInputElement) {
    filter.values = this.value;
  });
  return valuesInput;
}

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

function updateFilterDisplay() {
  const filterContainers = document.getElementById("filter-containers")!;
  filterContainers.innerHTML = "";
}

async function generateEmailDraft() {
  const promptElement = document.getElementById("myTextarea") as HTMLTextAreaElement;
  const resultElement = document.getElementById("result") as HTMLTextAreaElement;
  const prompt = promptElement.value;

  if (!prompt) {
    showNotification("Please enter a prompt before generating an email draft.", true);
    return;
  }

  try {
    Logging.log("Generating email draft...");
    showNotification("Generating email draft...");
    const generatedText = await callClaudeAPI(prompt, columnTitles);
    resultElement.value = generatedText;
    showNotification("Email draft generated successfully!");
  } catch (error) {
    Logging.error("Error in generateEmailDraft:", error);
    resultElement.value = "An error occurred. Please try again. Error details: " + error.message;
    showNotification("Failed to generate email draft. Please try again.", true);
  }
}

async function callClaudeAPI(prompt: string, columnTitles: string[]): Promise<string> {
  const apiUrl = "http://localhost:3001/api/generate";

  try {
    Logging.log("Calling Claude API...");
    Logging.log("Prompt:", prompt);
    Logging.log("Column titles:", columnTitles);
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        prompt: `Human: You are an AI assistant helping a university teacher write an email to students. The teacher will provide instructions, and you should write a professional email based on those instructions. Use {{column_title}} as placeholders for personalized information. Available column titles are: ${columnTitles.join(", ")}. only stick to the available column titles in the provided menu. Only provide the email draft ready to be sent instead of writing something along the lines of "Here is an email..." at the begining. Here are the teacher's instructions: "${prompt}"

              Assistant:`,
        max_tokens_to_sample: 300,
        temperature: 0.7,
      }),
    });

    const data = await response.json();
    Logging.log("API response:", JSON.stringify(data, null, 2));

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}, message: ${JSON.stringify(data, null, 2)}`);
    }

    if (!data.completion) {
      throw new Error("No completion in API response: " + JSON.stringify(data, null, 2));
    }

    return data.completion;
  } catch (error) {
    Logging.error("Error calling Claude API:", error);
    throw error;
  }
}

async function sendEmails() {
  Logging.log("Sending emails...");

  try {
    await Excel.run(async (context) => {
      // Get the active worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Get the email column
      const emailColumnSelect = document.getElementById("email-column") as HTMLSelectElement;
      const emailColumnIndex = parseInt(emailColumnSelect.value);

      if (isNaN(emailColumnIndex)) {
        throw new Error("Please select an email column");
      }

      // Get the used range of the worksheet
      const usedRange = sheet.getUsedRange();
      usedRange.load("values");
      await context.sync();

      // Get the email draft template
      const emailTemplate = (document.getElementById("result") as HTMLTextAreaElement).value;

      // Prepare batch of emails
      const emails = [];

      // Iterate through rows and prepare emails
      for (let i = 1; i < usedRange.values.length; i++) {
        // Start from 1 to skip header row
        const row = usedRange.values[i];
        const emailAddress = row[emailColumnIndex];

        // Skip rows without a valid email
        if (!emailAddress || !isValidEmail(emailAddress)) {
          Logging.log(`Skipping row ${i + 1}: Invalid email address`);
          continue;
        }

        // Personalize the email content
        let personalizedContent = emailTemplate;
        for (let j = 0; j < row.length; j++) {
          const columnName = usedRange.values[0][j];
          const cellValue = row[j];
          personalizedContent = personalizedContent.replace(new RegExp(`{{${columnName}}}`, "g"), cellValue);
        }

        // Improve formatting
        personalizedContent = formatEmailContent(personalizedContent);

        // Add email to batch
        emails.push({
          to: emailAddress,
          subject: "Personalized Bulk Email",
          html: personalizedContent,
        });
      }

      Logging.log("Prepared emails:", JSON.stringify(emails, null, 2));

      const response = await fetch("http://localhost:3001/api/send-emails", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ emails }),
      });

      const responseData = await response.json();
      Logging.log("Server response:", responseData);

      if (!response.ok) {
        throw new Error(
          `HTTP error! status: ${response.status}, message: ${responseData.error}, details: ${JSON.stringify(responseData, null, 2)}`
        );
      }

      Logging.log("Emails sent:", responseData);

      showNotification("All emails sent successfully!");
    });
  } catch (error) {
    Logging.error("Error sending emails:", error);
    showNotification(`Error sending emails: ${error.message}`, true);
  }
}

function formatEmailContent(content: string): string {
  // Replace single newlines with <br> tags
  content = content.replace(/(?<!\n)\n(?!\n)/g, "<br>");

  // Replace double newlines with paragraph breaks
  content = content.replace(/\n\n/g, "</p><p>");

  // Wrap the entire content in a paragraph tag
  content = `<p>${content}</p>`;

  return content;
}

function isValidEmail(email: string) {
  const re =
    /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(String(email).toLowerCase());
}

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
