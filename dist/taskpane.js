/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	// The require scope
/******/ 	var __webpack_require__ = {};
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Logging: function() { return /* binding */ Logging; }
/* harmony export */ });
class Logging {
  static log(...args) {
    // eslint-disable-next-line
    console.log(args);
  }
  static error(...args) {
    // eslint-disable-next-line
    console.error(...rgs);
  }
}
let filters = [];
let columnTitles = [];
Office.onReady(info => {
  console.log("Office.onReady called");
  document.getElementById("loading").style.display = "none";
  if (info.host === Office.HostType.Excel) {
    console.log("Excel host detected");
    document.getElementById("app-body").style.display = "flex";
    Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Add event handler for worksheet change
      sheet.onChanged.add(handleWorksheetChange);

      // Add event handler for worksheet addition or deletion
      context.workbook.worksheets.onActivated.add(handleWorksheetChange);
      await context.sync();
      console.log("Event handlers added for worksheet changes");

      // Initial population of column titles and menus
      await getColumnTitles();
      populateColumnTitlesMenu();
    }).catch(error => {
      console.error("Error setting up event handlers:", error);
      showNotification("Failed to set up automatic updates. Please refresh manually if needed.", true);
    });

    // Remove the event listener for the refresh button as it's no longer needed
    // document.getElementById("refresh-columns")!.addEventListener("click", ...)

    // Keep other event listeners
    document.getElementById("generate-draft").addEventListener("click", generateEmailDraft);
    document.getElementById("send-emails").addEventListener("click", sendEmails);
    document.getElementById("options").addEventListener("change", handleOptionsChange);
    document.getElementById("add-filter").addEventListener("click", addFilter);
    document.getElementById("myTextarea").addEventListener("input", handlePromptInput);
  } else if (info.host === Office.HostType.Outlook) {
    console.log("Outlook host detected");
    document.getElementById("app-body").style.display = "flex";
    // Add any Outlook-specific initialization here
  } else {
    console.log("Unsupported host detected:", info.host);
  }
});
async function getColumnTitles() {
  try {
    await Excel.run(async context => {
      console.log("Fetching column titles...");
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      const headerRow = range.getRow(0);
      headerRow.load("values");
      await context.sync();
      if (!headerRow.values || headerRow.values.length === 0 || headerRow.values[0].length === 0) {
        throw new Error("No data found in the first row of the sheet");
      }
      columnTitles = headerRow.values[0];
      console.log("Column titles updated:", columnTitles);

      // Update the email column menu
      populateEmailColumnMenu();
    });
  } catch (error) {
    console.error("Error fetching column titles:", error);
    throw error;
  }
}
function populateColumnTitlesMenu() {
  const menus = document.getElementsByClassName("columnTitlesMenu");
  Array.from(menus).forEach(menu => {
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
  const emailColumnMenu = document.getElementById("email-column");
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
function handlePromptInput(event) {
  const textarea = event.target;
  const cursorPosition = textarea.selectionStart;
  const textBeforeCursor = textarea.value.substring(0, cursorPosition);
  if (textBeforeCursor.endsWith("@")) {
    showColumnTitlesMenu(textarea, cursorPosition);
  }
}
function showColumnTitlesMenu(textarea, position) {
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
  const {
    top,
    left
  } = getCaretCoordinates(textarea, position);
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
  document.addEventListener("click", function closeDropdown(e) {
    if (!dropdown.contains(e.target) && e.target !== textarea) {
      dropdown.remove();
      document.removeEventListener("click", closeDropdown);
    }
  });
}
function getCaretCoordinates(element, position) {
  const div = document.createElement("div");
  const styles = getComputedStyle(element);
  const properties = ["fontFamily", "fontSize", "fontWeight", "fontStyle", "letterSpacing", "textTransform", "wordSpacing", "textIndent", "whiteSpace", "lineHeight", "padding", "border", "boxSizing"];
  properties.forEach(prop => {
    div.style[prop] = styles[prop];
  });
  div.textContent = element.value.substring(0, position);
  div.style.position = "absolute";
  div.style.visibility = "hidden";
  div.style.whiteSpace = "pre-wrap";
  document.body.appendChild(div);
  const coordinates = {
    top: div.offsetHeight - element.scrollTop,
    left: div.offsetWidth - element.scrollLeft
  };
  document.body.removeChild(div);
  return coordinates;
}
const properties = ["direction", "boxSizing", "width", "height", "overflowX", "overflowY", "borderTopWidth", "borderRightWidth", "borderBottomWidth", "borderLeftWidth", "paddingTop", "paddingRight", "paddingBottom", "paddingLeft", "fontStyle", "fontVariant", "fontWeight", "fontStretch", "fontSize", "fontSizeAdjust", "lineHeight", "fontFamily", "textAlign", "textTransform", "textIndent", "textDecoration", "letterSpacing", "wordSpacing"];
function insertColumnTitle(textarea, title, position) {
  const before = textarea.value.substring(0, position);
  const after = textarea.value.substring(position);
  textarea.value = before + title + after;
  textarea.selectionStart = textarea.selectionEnd = position + title.length;
  textarea.focus();
}
function handleOptionsChange(event) {
  const select = event.target;
  const addFilterButton = document.getElementById("add-filter");
  addFilterButton.style.display = select.value === "specific" ? "inline-block" : "none";
  if (select.value !== "specific") {
    filters = [];
    updateFilterDisplay();
    console.log("Filters cleared:", JSON.stringify(filters));
  }
}
function addFilter() {
  const filter = {
    column: "",
    values: ""
  };
  const filterContainer = document.createElement("div");
  filterContainer.className = "filter-container";
  const columnSelect = createColumnSelect(filter);
  const valuesInput = createValuesInput(filter);
  const deleteButton = createDeleteButton(filter, filterContainer);
  filterContainer.appendChild(columnSelect);
  filterContainer.appendChild(valuesInput);
  filterContainer.appendChild(deleteButton);
  document.getElementById("filter-containers").appendChild(filterContainer);
  filters.push(filter);
  console.log("Current filters:", JSON.stringify(filters));

  // Populate the newly created column select
  populateColumnSelect(columnSelect);
}
function createColumnSelect(filter) {
  const columnSelect = document.createElement("select");
  columnSelect.className = "columnTitlesMenu";
  columnSelect.addEventListener("change", function () {
    filter.column = columnTitles[parseInt(this.value)];
  });
  return columnSelect;
}
function populateColumnSelect(columnSelect) {
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
function createValuesInput(filter) {
  const valuesInput = document.createElement("input");
  valuesInput.type = "text";
  valuesInput.placeholder = "Enter values (comma-separated)";
  valuesInput.addEventListener("input", function () {
    filter.values = this.value;
  });
  return valuesInput;
}
function createDeleteButton(filter, filterContainer) {
  const deleteButton = document.createElement("button");
  deleteButton.textContent = "Delete";
  deleteButton.addEventListener("click", function () {
    filterContainer.remove();
    const index = filters.findIndex(f => f === filter);
    if (index !== -1) {
      filters.splice(index, 1);
    }
  });
  return deleteButton;
}
function updateFilterDisplay() {
  const filterContainers = document.getElementById("filter-containers");
  filterContainers.innerHTML = "";
}
async function generateEmailDraft() {
  const promptElement = document.getElementById("myTextarea");
  const subjectElement = document.getElementById("email-subject");
  const bodyElement = document.getElementById("email-body");
  const prompt = promptElement.value;
  if (!prompt) {
    showNotification("Please enter a prompt before generating an email draft.", true);
    return;
  }
  try {
    console.log("Generating email draft...");
    showNotification("Generating email draft...");
    const generatedText = await callClaudeAPI(prompt, columnTitles);
    const {
      subject,
      body
    } = parseGeneratedEmail(generatedText);
    subjectElement.value = subject;
    bodyElement.value = body;
    showNotification("Email draft generated successfully!");
  } catch (error) {
    console.error("Error in generateEmailDraft:", error);
    subjectElement.value = "Error generating subject";
    bodyElement.value = "An error occurred. Please try again. Error details: " + error.message;
    showNotification("Failed to generate email draft. Please try again.", true);
  }
}
function parseGeneratedEmail(text) {
  const parts = text.split('\n\n');
  return {
    subject: parts[0].replace('Subject: ', '').trim(),
    body: parts.slice(1).join('\n\n').trim()
  };
}
async function callClaudeAPI(prompt, columnTitles) {
  const apiUrl = "http://localhost:3001/api/generate";
  try {
    console.log("Calling Claude API...");
    console.log("Prompt:", prompt);
    console.log("Column titles:", columnTitles);
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        prompt: `Human: You are an AI assistant helping a university teacher write an email to students. The teacher will provide instructions, and you should write a professional email based on those instructions. Generate both a subject line and an email body. Use {{column_title}} as placeholders for personalized information. Available column titles are: ${columnTitles.join(", ")}. only stick to the available column titles in the provided menu. Only provide the email draft ready to be sent instead of writing something along the lines of "Here is an email..." at the begining. Provide the email draft in the following format:

                Subject: [Generated Subject]

                [Generated Email Body]

                Here are the teacher's instructions: "${prompt}"

                Assistant:`,
        max_tokens_to_sample: 300,
        temperature: 0.7
      })
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
async function sendEmails() {
  console.log("Sending emails...");
  try {
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const emailColumnSelect = document.getElementById("email-column");
      const emailColumnIndex = parseInt(emailColumnSelect.value);
      console.log(`Selected email column index: ${emailColumnIndex}`);
      if (isNaN(emailColumnIndex)) {
        throw new Error("Please select an email column");
      }
      const usedRange = sheet.getUsedRange();
      usedRange.load("values");
      await context.sync();
      const headerRow = usedRange.values[0];
      console.log(`Header row: ${JSON.stringify(headerRow)}`);
      const subjectElement = document.getElementById("email-subject");
      const bodyElement = document.getElementById("email-body");
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
      const optionsSelect = document.getElementById("options");
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
          html: personalizedBody
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
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          emails
        })
      });
      const responseData = await response.json();
      console.log("Server response:", JSON.stringify(responseData, null, 2));
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}, message: ${responseData.error}, details: ${JSON.stringify(responseData, null, 2)}`);
      }
      if (responseData.sentEmails && Array.isArray(responseData.sentEmails)) {
        console.log(`Successfully sent ${responseData.sentEmails.length} out of ${emails.length} emails`);
        responseData.sentEmails.forEach(email => {
          console.log(`Email sent to: ${email}`);
        });
      }
      if (responseData.failedEmails && Array.isArray(responseData.failedEmails)) {
        console.log(`Failed to send ${responseData.failedEmails.length} emails`);
        responseData.failedEmails.forEach(failedEmail => {
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
function formatEmailContent(content) {
  // Replace single newlines with <br> tags
  content = content.replace(/(?<!\n)\n(?!\n)/g, "<br>");

  // Replace double newlines with paragraph breaks
  content = content.replace(/\n\n/g, "</p><p>");

  // Wrap the entire content in a paragraph tag
  content = `<p>${content}</p>`;
  return content;
}
function isValidEmail(email) {
  const re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(String(email).toLowerCase());
}
function showNotification(message, isError = false) {
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
function applyFilters(row, headerRow) {
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
function detectEmailColumn(headerRow) {
  const emailRegex = /email|e-mail|mail/i;
  return headerRow.findIndex(header => emailRegex.test(header));
}
async function handleWorksheetChange(event) {
  try {
    console.log("Worksheet changed. Updating column titles and email menu.");
    await getColumnTitles();
    populateColumnTitlesMenu();
    // populateEmailColumnMenu() is now called within getColumnTitles()
  } catch (error) {
    console.error("Error handling worksheet change:", error);
    showNotification("Failed to update columns. Please try again.", true);
  }
}
/******/ })()
;
//# sourceMappingURL=taskpane.js.map