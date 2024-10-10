var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var _this = this;
var EMAIL_DISCLAIMER = "This email was sent using the Excel add-in, 'teacherhelper'. Please do NOT reply to no-reply.teacherhelper@outlook.com";
// Initialize global variables
var filters = [];
var columnTitles = [];
var instructionsVisible = false;
// Initialize the counter
var emailCounter = 100;
var lastResetDate;
// At the top of your file, add these constants for original values
var ORIGINAL_INSTITUTION = "university";
var ORIGINAL_PERSONA = "professor";
var ORIGINAL_AUDIENCE = "students";
var ORIGINAL_TONE = "professional";
var ORIGINAL_MAX_TOKENS = 300;
var ORIGINAL_TEMPERATURE = 0.7;
// Your existing variables
var institution = ORIGINAL_INSTITUTION;
var persona = ORIGINAL_PERSONA;
var audience = ORIGINAL_AUDIENCE;
var tone = ORIGINAL_TONE;
var maxTokens = ORIGINAL_MAX_TOKENS;
var temperature = ORIGINAL_TEMPERATURE;
// Office.onReady function - runs when the Office APIs are ready
Office.onReady(function (info) {
    console.log("Office.onReady called");
    document.getElementById("loading").style.display = "none";
    if (info.host === Office.HostType.Excel) {
        console.log("Excel host detected");
        document.getElementById("app-body").style.display = "flex";
        setupInstructionsToggle();
        initializeAdvancedSettings();
        loadCounterState();
        setupSignInToggle();
        loadSignInOptions();
        // Set up Excel-specific functionality
        Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
            var sheet;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                        // event handler for worksheet change
                        sheet.onChanged.add(handleWorksheetChange);
                        // event handler for worksheet addition or deletion
                        context.workbook.worksheets.onActivated.add(handleWorksheetChange);
                        return [4 /*yield*/, context.sync()];
                    case 1:
                        _a.sent();
                        console.log("Event handlers added for worksheet changes");
                        // Initial population of column titles and menus
                        return [4 /*yield*/, getColumnTitles()];
                    case 2:
                        // Initial population of column titles and menus
                        _a.sent();
                        populateColumnTitlesMenu();
                        return [2 /*return*/];
                }
            });
        }); }).catch(function (error) {
            console.error("Error setting up event handlers:", error);
            showNotification("Failed to set up automatic updates. Please refresh manually if needed.", true);
        });
        // Load the saved signature
        loadSignature();
        // Load the saved sender email
        loadSenderEmail();
        // event listeners for UI elements
        document.getElementById("generate-draft").addEventListener("click", generateEmailDraft);
        document.getElementById("send-emails").addEventListener("click", sendEmails);
        document.getElementById("options").addEventListener("change", handleOptionsChange);
        document.getElementById("add-filter").addEventListener("click", addFilter);
        document.getElementById("save-signature").addEventListener("click", saveSignature);
        document.getElementById("save-sender-email").addEventListener("click", saveSenderEmail);
        document.getElementById("save-signin-options").addEventListener("click", saveSignInOptions);
    }
    else {
        console.log("Unsupported host detected:", info.host);
    }
});
// Function to get column titles from the active Excel worksheet
function getColumnTitles() {
    return __awaiter(this, void 0, void 0, function () {
        var error_1;
        var _this = this;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
                            var sheet, range, headerRow;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        console.log("Fetching column titles...");
                                        sheet = context.workbook.worksheets.getActiveWorksheet();
                                        range = sheet.getUsedRange();
                                        headerRow = range.getRow(0);
                                        headerRow.load("values");
                                        return [4 /*yield*/, context.sync()];
                                    case 1:
                                        _a.sent();
                                        if (!headerRow.values || headerRow.values.length === 0 || headerRow.values[0].length === 0) {
                                            throw new Error("No data found in the first row of the sheet");
                                        }
                                        columnTitles = headerRow.values[0];
                                        console.log("Column titles updated:", columnTitles);
                                        // Update the email column menu
                                        populateEmailColumnMenu();
                                        return [2 /*return*/];
                                }
                            });
                        }); })];
                case 1:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 2:
                    error_1 = _a.sent();
                    console.error("Error fetching column titles:", error_1);
                    throw error_1;
                case 3: return [2 /*return*/];
            }
        });
    });
}
// Function to populate dropdown menus with column titles
function populateColumnTitlesMenu() {
    var menus = document.getElementsByClassName("columnTitlesMenu");
    Array.from(menus).forEach(function (menu) {
        menu.innerHTML = ""; // Clear existing options
        // Add a default option
        var defaultOption = document.createElement("option");
        defaultOption.text = "Select a column";
        defaultOption.value = "";
        menu.appendChild(defaultOption);
        // Add options for each column title
        columnTitles.forEach(function (title, index) {
            var option = document.createElement("option");
            option.value = index.toString();
            option.text = title;
            menu.appendChild(option);
        });
    });
}
// Function to populate the email column dropdown menu
function populateEmailColumnMenu() {
    var emailColumnMenu = document.getElementById("email-column");
    if (!emailColumnMenu) {
        console.error("Email column menu not found");
        return;
    }
    emailColumnMenu.innerHTML = ""; // Clear existing options
    // Add a default option
    var defaultOption = document.createElement("option");
    defaultOption.text = "Select a column";
    defaultOption.value = "";
    emailColumnMenu.appendChild(defaultOption);
    // Add options for each column title
    columnTitles.forEach(function (title, index) {
        var option = document.createElement("option");
        option.value = index.toString();
        option.text = title;
        emailColumnMenu.appendChild(option);
    });
    // Detect and select email column
    var emailColumnIndex = detectEmailColumn(columnTitles);
    if (emailColumnIndex !== -1) {
        emailColumnMenu.value = emailColumnIndex.toString();
    }
    console.log("Email column menu populated with", columnTitles.length, "options");
    console.log("Automatically selected email column:", emailColumnIndex);
}
// Function to insert a column title into the textarea
function insertColumnTitle(textarea, title, position) {
    var before = textarea.value.substring(0, position);
    var after = textarea.value.substring(position);
    textarea.value = before + title + after;
    textarea.selectionStart = textarea.selectionEnd = position + title.length;
    textarea.focus();
}
// Function to handle changes in the options dropdown
function handleOptionsChange(event) {
    var select = event.target;
    var addFilterButton = document.getElementById("add-filter");
    addFilterButton.style.display = select.value === "specific" ? "inline-block" : "none";
    if (select.value !== "specific") {
        filters = [];
        updateFilterDisplay();
        console.log("Filters cleared:", JSON.stringify(filters));
    }
}
// Function to add a new filter
function addFilter() {
    var filter = {
        column: "",
        values: "",
    };
    var filterContainer = document.createElement("div");
    filterContainer.className = "filter-container";
    var columnSelect = createColumnSelect(filter);
    var valuesInput = createValuesInput(filter);
    var deleteButton = createDeleteButton(filter, filterContainer);
    filterContainer.appendChild(columnSelect);
    filterContainer.appendChild(valuesInput);
    filterContainer.appendChild(deleteButton);
    document.getElementById("filter-containers").appendChild(filterContainer);
    filters.push(filter);
    console.log("Current filters:", JSON.stringify(filters));
    // Populate the newly created column select
    populateColumnSelect(columnSelect);
}
// Function to create a column select dropdown for filters
function createColumnSelect(filter) {
    var columnSelect = document.createElement("select");
    columnSelect.className = "columnTitlesMenu";
    columnSelect.addEventListener("change", function () {
        filter.column = columnTitles[parseInt(this.value)];
    });
    return columnSelect;
}
// Function to populate a column select dropdown
function populateColumnSelect(columnSelect) {
    columnSelect.innerHTML = ""; // Clear existing options
    // Add a default option
    var defaultOption = document.createElement("option");
    defaultOption.text = "Select a column";
    defaultOption.value = "";
    columnSelect.appendChild(defaultOption);
    // Add options for each column title
    columnTitles.forEach(function (title, index) {
        var option = document.createElement("option");
        option.value = index.toString();
        option.text = title;
        columnSelect.appendChild(option);
    });
}
// Function to create an input field for filter values
function createValuesInput(filter) {
    var valuesInput = document.createElement("input");
    valuesInput.type = "text";
    valuesInput.placeholder = "Enter values (comma-separated)";
    valuesInput.addEventListener("input", function () {
        filter.values = this.value;
    });
    return valuesInput;
}
// Function to create a delete button for filters
function createDeleteButton(filter, filterContainer) {
    var deleteButton = document.createElement("button");
    deleteButton.textContent = "Delete";
    deleteButton.addEventListener("click", function () {
        filterContainer.remove();
        var index = filters.findIndex(function (f) { return f === filter; });
        if (index !== -1) {
            filters.splice(index, 1);
        }
    });
    return deleteButton;
}
// Function to update the filter display
function updateFilterDisplay() {
    var filterContainers = document.getElementById("filter-containers");
    filterContainers.innerHTML = "";
}
// Function to generate an email draft using the Claude API
function generateEmailDraft() {
    return __awaiter(this, void 0, void 0, function () {
        var promptElement, subjectElement, bodyElement, prompt, generatedText, _a, subject, body, signatureElement, signature, error_2;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    promptElement = document.getElementById("myTextarea");
                    subjectElement = document.getElementById("email-subject");
                    bodyElement = document.getElementById("email-body");
                    prompt = promptElement.value;
                    if (!prompt) {
                        showNotification("Please enter a prompt before generating an email draft.", true);
                        return [2 /*return*/];
                    }
                    _b.label = 1;
                case 1:
                    _b.trys.push([1, 3, , 4]);
                    console.log("Generating email draft...");
                    showNotification("Generating email draft...");
                    return [4 /*yield*/, callClaudeAPI(prompt, columnTitles)];
                case 2:
                    generatedText = _b.sent();
                    _a = parseGeneratedEmail(generatedText), subject = _a.subject, body = _a.body;
                    subjectElement.value = subject;
                    bodyElement.value = body;
                    signatureElement = document.getElementById("email-signature");
                    signature = signatureElement.value;
                    bodyElement.value = body + "\n\n" + signature + "\n\n" + "\n\n".concat(EMAIL_DISCLAIMER);
                    showNotification("Email draft generated successfully!");
                    return [3 /*break*/, 4];
                case 3:
                    error_2 = _b.sent();
                    console.error("Error in generateEmailDraft:", error_2);
                    subjectElement.value = "Error generating subject";
                    bodyElement.value = "An error occurred. Please try again. Error details: " + error_2.message;
                    showNotification("Failed to generate email draft. Please try again.", true);
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    });
}
function parseGeneratedEmail(text) {
    var parts = text.split('\n\n');
    return {
        subject: parts[0].replace('Subject: ', '').trim(),
        body: parts.slice(1).join('\n\n').trim()
    };
}
// Function to connect with Claude API
function callClaudeAPI(prompt, columnTitles) {
    return __awaiter(this, void 0, void 0, function () {
        var apiUrl, ANTHROPIC_API_KEY, institution, persona, audience, tone, response, data, error_3;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    apiUrl = "http://localhost:3001/api/generate";
                    ANTHROPIC_API_KEY = document.getElementById("anthropic-api").value;
                    institution = "university";
                    persona = "professor";
                    audience = "students";
                    tone = "professional";
                    if (!ANTHROPIC_API_KEY) {
                        throw new Error("Anthropic API key is not set. Please enter it in the Sign In section.");
                    }
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 4, , 5]);
                    console.log("Calling Claude API...");
                    console.log("Prompt:", prompt);
                    console.log("Column titles:", columnTitles);
                    return [4 /*yield*/, fetch(apiUrl, {
                            method: "POST",
                            headers: {
                                "Content-Type": "application/json",
                            },
                            body: JSON.stringify({
                                prompt: "Human: \n                 Forget any previous instructions.\n                 You are an AI assistant helping a ".concat(persona, " at a ").concat(institution, " write an ").concat(tone, " email to to ").concat(audience, ". \n                 The ").concat(persona, " will provide instructions, and you should write a ").concat(tone, " email based on those instructions. \n                 Generate both a subject line and an email body. \n                 Use {{column title}} as placeholders for personalised information.\n                 Only use curly brackets {} This is the only type of brackets you are allowed to use.\n                 Available column titles are: ").concat(columnTitles.join(", "), ". \n                 Only stick to the available column titles in the provided menu.\n                 Do not repeat the column titles multiple times if you are creating a list within the generated email body. \n                 This is very important: only provide the email draft ready to be sent instead of writing something along the lines of \"Here is an email...\" at the beginning. \n                 Do not use a signature for the ").concat(persona, ". \n                 Stop generating the email body after you write \"kind regards\" or \"sincerely\" or the other similar words that are used to end emails. \n                 Provide the email draft in the following format:\n\n                [Generated email title]\n\n                [Generated Email Body]\n\n                Here are the teacher's instructions: \"").concat(prompt, "\"\n\n                Assistant:"),
                                max_tokens_to_sample: maxTokens,
                                temperature: temperature,
                                ANTHROPIC_API_KEY: ANTHROPIC_API_KEY
                            }),
                        })];
                case 2:
                    response = _a.sent();
                    return [4 /*yield*/, response.json()];
                case 3:
                    data = _a.sent();
                    console.log("API response:", JSON.stringify(data, null, 2));
                    if (!response.ok) {
                        throw new Error("HTTP error! status: ".concat(response.status, ", message: ").concat(JSON.stringify(data, null, 2)));
                    }
                    if (!data.completion) {
                        throw new Error("No completion in API response: " + JSON.stringify(data, null, 2));
                    }
                    return [2 /*return*/, data.completion];
                case 4:
                    error_3 = _a.sent();
                    console.error("Error calling Claude API:", error_3);
                    throw error_3;
                case 5: return [2 /*return*/];
            }
        });
    });
}
// Function to send the generated emails
function sendEmails() {
    return __awaiter(this, void 0, void 0, function () {
        var SENDGRID_API_KEY, sendgridEmail, error_4;
        var _this = this;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    console.log("Sending emails...");
                    SENDGRID_API_KEY = document.getElementById("sendgrid-api").value;
                    sendgridEmail = document.getElementById("sendgrid-email").value;
                    if (!SENDGRID_API_KEY || !sendgridEmail) {
                        showNotification("SendGrid API key and email are required. Please enter them in the Sign In section.", true);
                        return [2 /*return*/];
                    }
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
                            var sheet, emailColumnSelect, emailColumnIndex, usedRange, headerRow, subjectElement, bodyElement, senderEmailElement, senderEmail, signatureElement, emailSubjectTemplate, emailBodyTemplate, emails, skippedRows, totalRows, filteredOutRows, optionsSelect, useFilters, i, row, emailAddress, passesFilter, totalEmailsToSend, personalizedSubject, personalizedBody, j, columnName, cellValue, placeholder, response, responseData;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        sheet = context.workbook.worksheets.getActiveWorksheet();
                                        emailColumnSelect = document.getElementById("email-column");
                                        emailColumnIndex = parseInt(emailColumnSelect.value);
                                        console.log("Selected email column index: ".concat(emailColumnIndex));
                                        if (isNaN(emailColumnIndex)) {
                                            throw new Error("Please select an email column");
                                        }
                                        usedRange = sheet.getUsedRange();
                                        usedRange.load("values");
                                        return [4 /*yield*/, context.sync()];
                                    case 1:
                                        _a.sent();
                                        headerRow = usedRange.values[0];
                                        console.log("Header row: ".concat(JSON.stringify(headerRow)));
                                        subjectElement = document.getElementById("email-subject");
                                        bodyElement = document.getElementById("email-body");
                                        senderEmailElement = document.getElementById("sender-email");
                                        senderEmail = senderEmailElement.value.trim();
                                        signatureElement = document.getElementById("email-signature");
                                        if (!subjectElement || !bodyElement) {
                                            throw new Error("Email subject or body element not found");
                                        }
                                        emailSubjectTemplate = subjectElement.value;
                                        emailBodyTemplate = bodyElement.value;
                                        console.log("Email subject template: ".concat(emailSubjectTemplate));
                                        console.log("Email body template: ".concat(emailBodyTemplate));
                                        if (!emailSubjectTemplate || !emailBodyTemplate) {
                                            throw new Error("Email subject or body template is empty");
                                        }
                                        emails = [];
                                        skippedRows = 0;
                                        totalRows = usedRange.values.length - 1;
                                        filteredOutRows = 0;
                                        console.log("Processing ".concat(totalRows, " rows..."));
                                        optionsSelect = document.getElementById("options");
                                        useFilters = optionsSelect.value === "specific";
                                        console.log("Using filters: ".concat(useFilters));
                                        console.log("Current filters: ".concat(JSON.stringify(filters)));
                                        for (i = 1; i < usedRange.values.length; i++) {
                                            row = usedRange.values[i];
                                            console.log("Processing row ".concat(i, ": ").concat(JSON.stringify(row)));
                                            emailAddress = row[emailColumnIndex];
                                            console.log("Email address found: ".concat(emailAddress));
                                            if (!emailAddress || !isValidEmail(emailAddress)) {
                                                console.log("Skipping row ".concat(i, ": Invalid email address"));
                                                skippedRows++;
                                                continue;
                                            }
                                            if (useFilters) {
                                                passesFilter = applyFilters(row, headerRow);
                                                console.log("Row ".concat(i, " passes filter: ").concat(passesFilter));
                                                if (!passesFilter) {
                                                    filteredOutRows++;
                                                    continue;
                                                }
                                            }
                                            totalEmailsToSend = emails.length + (senderEmail ? 1 : 0);
                                            if (totalEmailsToSend > emailCounter) {
                                                showWarningModal("Warning: The number of emails to be sent (".concat(totalEmailsToSend, ") exceeds your remaining email sending credit (").concat(emailCounter, "). Please reduce the number of recipients or try again tomorrow."));
                                                return [2 /*return*/];
                                            }
                                            // If sender email is provided and valid, add it to the emails array
                                            if (senderEmail && isValidEmail(senderEmail)) {
                                                emails.push({
                                                    to: senderEmail,
                                                    subject: subjectElement.value,
                                                    html: formatEmailContent(bodyElement.value),
                                                });
                                            }
                                            personalizedSubject = emailSubjectTemplate;
                                            personalizedBody = emailBodyTemplate;
                                            for (j = 0; j < row.length; j++) {
                                                columnName = headerRow[j];
                                                cellValue = row[j];
                                                placeholder = "{{".concat(columnName, "}}");
                                                personalizedSubject = personalizedSubject.replace(new RegExp(placeholder, 'g'), cellValue);
                                                personalizedBody = personalizedBody.replace(new RegExp(placeholder, 'g'), cellValue);
                                            }
                                            personalizedBody = formatEmailContent(personalizedBody);
                                            emails.push({
                                                to: emailAddress,
                                                subject: personalizedSubject,
                                                html: personalizedBody,
                                            });
                                            console.log("Added email for ".concat(emailAddress, " with subject: ").concat(personalizedSubject));
                                        }
                                        console.log("Total rows processed: ".concat(totalRows));
                                        console.log("Skipped rows (invalid email): ".concat(skippedRows));
                                        console.log("Filtered out rows: ".concat(filteredOutRows));
                                        console.log("Valid emails found: ".concat(emails.length));
                                        if (emails.length === 0) {
                                            throw new Error("No valid emails to send. Please check your data and filter criteria.");
                                        }
                                        console.log("Prepared emails:", JSON.stringify(emails, null, 2));
                                        console.log("Preparing to send ".concat(emails.length, " emails"));
                                        return [4 /*yield*/, fetch("http://localhost:3001/api/send-emails", {
                                                method: "POST",
                                                headers: {
                                                    "Content-Type": "application/json",
                                                },
                                                body: JSON.stringify({
                                                    emails: emails,
                                                    SENDGRID_API_KEY: SENDGRID_API_KEY,
                                                    sendgridEmail: sendgridEmail
                                                }),
                                            })];
                                    case 2:
                                        response = _a.sent();
                                        return [4 /*yield*/, response.json()];
                                    case 3:
                                        responseData = _a.sent();
                                        console.log("Server response:", JSON.stringify(responseData, null, 2));
                                        if (!response.ok) {
                                            throw new Error("HTTP error! status: ".concat(response.status, ", message: ").concat(responseData.error, ", details: ").concat(JSON.stringify(responseData, null, 2)));
                                        }
                                        if (responseData.sentEmails && Array.isArray(responseData.sentEmails)) {
                                            console.log("Successfully sent ".concat(responseData.sentEmails.length, " out of ").concat(emails.length, " emails"));
                                            responseData.sentEmails.forEach(function (email) {
                                                console.log("Email sent to: ".concat(email));
                                                decreaseCounter();
                                            });
                                        }
                                        if (responseData.failedEmails && Array.isArray(responseData.failedEmails)) {
                                            console.log("Failed to send ".concat(responseData.failedEmails.length, " emails"));
                                            responseData.failedEmails.forEach(function (failedEmail) {
                                                console.log("Failed to send email to ".concat(failedEmail.email, ". Error: ").concat(failedEmail.error));
                                            });
                                        }
                                        showNotification("Emails sent: ".concat(responseData.sentEmails.length, "/").concat(emails.length, ". Check console for details."));
                                        return [2 /*return*/];
                                }
                            });
                        }); })];
                case 2:
                    _a.sent();
                    return [3 /*break*/, 4];
                case 3:
                    error_4 = _a.sent();
                    console.error("Error sending emails:", error_4);
                    showNotification("Error sending emails: ".concat(error_4.message), true);
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    });
}
// Function to format  the email content
function formatEmailContent(content) {
    // Replace single newlines with <br> tags
    content = content.replace(/(?<!\n)\n(?!\n)/g, "<br>");
    // Replace double newlines with paragraph breaks
    content = content.replace(/\n\n/g, "</p><p>");
    // Wrap the entire content in a paragraph tag
    content = "<p>".concat(content, "</p>");
    return content;
}
// Function to check email validity within a column
function isValidEmail(email) {
    var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(String(email).toLowerCase());
}
// Function to display notifications
function showNotification(message, isError) {
    if (isError === void 0) { isError = false; }
    var notificationElement = document.createElement("div");
    notificationElement.textContent = message;
    notificationElement.style.position = "fixed";
    notificationElement.style.top = "10px";
    notificationElement.style.left = "50%";
    notificationElement.style.transform = "translateX(-50%)";
    notificationElement.style.padding = "10px";
    notificationElement.style.borderRadius = "5px";
    notificationElement.style.backgroundColor = isError ? "#ffcccc" : "#ccffcc";
    notificationElement.style.border = "1px solid ".concat(isError ? "#ff0000" : "#00ff00");
    document.body.appendChild(notificationElement);
    setTimeout(function () {
        document.body.removeChild(notificationElement);
    }, 3000);
}
// Function to apply created filters
function applyFilters(row, headerRow) {
    if (filters.length === 0)
        return true;
    console.log("Applying filters:", JSON.stringify(filters));
    return filters.every(function (filter) {
        var _a, _b;
        var columnIndex = headerRow.findIndex(function (col) { return col.toLowerCase() === filter.column.toLowerCase(); });
        if (columnIndex === -1) {
            console.log("Column not found: ".concat(filter.column));
            return false;
        }
        var cellValue = (_b = (_a = row[columnIndex]) === null || _a === void 0 ? void 0 : _a.toString().toLowerCase()) !== null && _b !== void 0 ? _b : "";
        var filterValues = filter.values.toLowerCase().split(',').map(function (v) { return v.trim(); });
        console.log("Checking column: ".concat(filter.column, ", Cell value: ").concat(cellValue, ", Filter values: ").concat(filterValues.join(', ')));
        var result = filterValues.some(function (value) { return cellValue.includes(value); });
        console.log("Filter result for ".concat(filter.column, ": ").concat(result));
        return result;
    });
}
// Function to automatically detect email columns
function detectEmailColumn(headerRow) {
    var emailRegex = /email|e-mail|mail/i;
    return headerRow.findIndex(function (header) { return emailRegex.test(header); });
}
// Function to automatically register changes in the worksheet
function handleWorksheetChange(event) {
    return __awaiter(this, void 0, void 0, function () {
        var error_5;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    console.log("Worksheet changed. Updating column titles and email menu.");
                    return [4 /*yield*/, getColumnTitles()];
                case 1:
                    _a.sent();
                    populateColumnTitlesMenu();
                    return [3 /*break*/, 3];
                case 2:
                    error_5 = _a.sent();
                    console.error("Error handling worksheet change:", error_5);
                    showNotification("Failed to update columns. Please try again.", true);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    });
}
// Function to save the teacher's signature
function saveSignature() {
    var signatureElement = document.getElementById("email-signature");
    var signature = signatureElement.value;
    Office.context.document.settings.set("teacherSignature", signature);
    Office.context.document.settings.saveAsync(function () {
        console.log("Signature saved successfully");
        showNotification("Signature saved successfully!");
    });
}
// Function to save the loaded signature
function loadSignature() {
    var signature = Office.context.document.settings.get("teacherSignature");
    if (signature) {
        var signatureElement = document.getElementById("email-signature");
        signatureElement.value = signature;
    }
}
function setupInstructionsToggle() {
    var toggleButton = document.getElementById("toggle-instructions");
    var instructions = document.getElementById("instructions");
    if (toggleButton && instructions) {
        toggleButton.onclick = function () {
            instructionsVisible = !instructionsVisible;
            instructions.style.display = instructionsVisible ? "block" : "none";
            toggleButton.textContent = instructionsVisible ? "Hide Instructions" : "Show instructions on how to use teacherhelper";
        };
    }
}
// Function to initialize advanced settings
function initializeAdvancedSettings() {
    var _a, _b, _c, _d, _e, _f;
    console.log("Initializing advanced settings");
    var advancedSettingsBtn = document.getElementById('advanced-settings-btn');
    var advancedSettings = document.getElementById('advanced-settings');
    var resetButton = document.getElementById('advanced-settings-reset');
    var saveButton = document.getElementById('advanced-settings-save');
    if (!advancedSettingsBtn || !advancedSettings || !resetButton || !saveButton) {
        console.error("One or more advanced settings elements not found");
        return;
    }
    advancedSettingsBtn.addEventListener('click', function () {
        advancedSettings.style.display = advancedSettings.style.display === 'none' ? 'block' : 'none';
    });
    resetButton.addEventListener('click', resetToOriginalValues);
    saveButton.addEventListener('click', saveControlSettings);
    // Load saved settings or use defaults
    loadSavedSettings();
    // Initialize input fields
    updateInputFields();
    // Add event listeners for input changes
    (_a = document.getElementById('institution')) === null || _a === void 0 ? void 0 : _a.addEventListener('input', function (e) {
        institution = e.target.value || ORIGINAL_INSTITUTION;
    });
    (_b = document.getElementById('persona')) === null || _b === void 0 ? void 0 : _b.addEventListener('input', function (e) {
        persona = e.target.value || ORIGINAL_PERSONA;
    });
    (_c = document.getElementById('audience')) === null || _c === void 0 ? void 0 : _c.addEventListener('input', function (e) {
        audience = e.target.value || ORIGINAL_AUDIENCE;
    });
    (_d = document.getElementById('tone')) === null || _d === void 0 ? void 0 : _d.addEventListener('input', function (e) {
        tone = e.target.value || ORIGINAL_TONE;
    });
    (_e = document.getElementById('max-tokens')) === null || _e === void 0 ? void 0 : _e.addEventListener('input', function (e) {
        maxTokens = parseInt(e.target.value);
        document.getElementById('max-tokens-value').textContent = maxTokens.toString();
    });
    (_f = document.getElementById('temperature')) === null || _f === void 0 ? void 0 : _f.addEventListener('input', function (e) {
        temperature = parseFloat(e.target.value);
        document.getElementById('temperature-value').textContent = temperature.toString();
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
    Office.context.document.settings.saveAsync(function () {
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
    document.getElementById('institution').value = institution;
    document.getElementById('persona').value = persona;
    document.getElementById('audience').value = audience;
    document.getElementById('tone').value = tone;
    document.getElementById('max-tokens').value = maxTokens.toString();
    document.getElementById('temperature').value = temperature.toString();
    document.getElementById('max-tokens-value').textContent = maxTokens.toString();
    document.getElementById('temperature-value').textContent = temperature.toString();
}
// Function to update the counter display
function updateCounterDisplay() {
    var counterElement = document.getElementById('email-counter');
    if (counterElement) {
        counterElement.textContent = "Emails remaining today: ".concat(emailCounter);
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
    var today = new Date().toDateString();
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
    var savedCounter = localStorage.getItem('emailCounter');
    var savedDate = localStorage.getItem('lastResetDate');
    if (savedCounter !== null) {
        emailCounter = parseInt(savedCounter, 10);
    }
    if (savedDate !== null) {
        lastResetDate = savedDate;
    }
    else {
        lastResetDate = new Date().toDateString();
    }
    checkAndResetCounter();
}
// Function to show a modal warning
function showWarningModal(message) {
    // Create modal elements
    var modal = document.createElement('div');
    modal.style.cssText = "\n    position: fixed;\n    z-index: 1000;\n    left: 0;\n    top: 0;\n    width: 100%;\n    height: 100%;\n    background-color: rgba(0,0,0,0.4);\n    display: flex;\n    justify-content: center;\n    align-items: center;\n  ";
    var modalContent = document.createElement('div');
    modalContent.style.cssText = "\n    background-color: #fefefe;\n    padding: 20px;\n    border: 1px solid #888;\n    width: 80%;\n    max-width: 400px;\n    text-align: center;\n  ";
    var closeBtn = document.createElement('button');
    closeBtn.textContent = 'Close';
    closeBtn.onclick = function () { return document.body.removeChild(modal); };
    modalContent.innerHTML = "<p>".concat(message, "</p>");
    modalContent.appendChild(closeBtn);
    modal.appendChild(modalContent);
    document.body.appendChild(modal);
}
function saveSenderEmail() {
    var senderEmailElement = document.getElementById("sender-email");
    var senderEmail = senderEmailElement.value.trim();
    if (senderEmail === "" || isValidEmail(senderEmail)) {
        Office.context.document.settings.set("senderEmail", senderEmail);
        Office.context.document.settings.saveAsync(function () {
            console.log("Sender email saved successfully");
            if (senderEmail === "") {
                showNotification("Your email has been cleared successfully!");
            }
            else {
                showNotification("Your email has been saved successfully!");
            }
        });
    }
    else {
        showNotification("Please enter a valid email address or leave it empty to clear.", true);
    }
}
function loadSenderEmail() {
    var senderEmail = Office.context.document.settings.get("senderEmail");
    if (senderEmail) {
        var senderEmailElement = document.getElementById("sender-email");
        senderEmailElement.value = senderEmail;
    }
}
var signInOptionsVisible = false;
function setupSignInToggle() {
    var toggleButton = document.getElementById("toggle-signin");
    var signInOptions = document.getElementById("signin-options");
    if (toggleButton && signInOptions) {
        toggleButton.onclick = function () {
            signInOptionsVisible = !signInOptionsVisible;
            signInOptions.style.display = signInOptionsVisible ? "block" : "none";
        };
    }
}
function saveSignInOptions() {
    var ANTHROPIC_API_KEY = document.getElementById("anthropic-api").value;
    var sendgridEmail = document.getElementById("sendgrid-email").value;
    var SENDGRID_API_KEY = document.getElementById("sendgrid-api").value;
    Office.context.document.settings.set("ANTHROPIC_API_KEY", ANTHROPIC_API_KEY);
    Office.context.document.settings.set("sendgridEmail", sendgridEmail);
    Office.context.document.settings.set("SENDGRID_API_KEY", SENDGRID_API_KEY);
    Office.context.document.settings.saveAsync(function () {
        console.log("Sign-in options saved successfully");
        showNotification("Sign-in options saved successfully!");
    });
}
function loadSignInOptions() {
    var ANTHROPIC_API_KEY = Office.context.document.settings.get("ANTHROPIC_API_KEY");
    var sendgridEmail = Office.context.document.settings.get("sendgridEmail");
    var SENDGRID_API_KEY = Office.context.document.settings.get("SENDGRID_API_KEY");
    if (ANTHROPIC_API_KEY) {
        document.getElementById("anthropic-api").value = ANTHROPIC_API_KEY;
    }
    if (sendgridEmail) {
        document.getElementById("sendgrid-email").value = sendgridEmail;
    }
    if (SENDGRID_API_KEY) {
        document.getElementById("sendgrid-api").value = SENDGRID_API_KEY;
    }
}
