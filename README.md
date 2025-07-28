MSRA Tool Code Documentation
This document provides an overview and detailed explanation of the Method Statement and Risk Assessment (MSRA) Tool's codebase, split into index.html, style.css, and script.js for better organization and maintainability.
1. index.html
This file forms the structure of the MSRA Tool web application. It includes all the visual elements users interact with, along with links to external libraries and local CSS and JavaScript files.
Key Sections:
 * Document Metadata (<head>):
   * charset="UTF-8": Specifies character encoding.
   * viewport: Configures the viewport for responsive behavior across devices.
   * <title>: Sets the title displayed in the browser tab.
   * External Libraries: Links to html2canvas.min.js (for image export) and jspdf.umd.min.js (for PDF export) via CDN.
   * Stylesheet Link: Links to style.css for styling the application.
 * Main Container (<body>):
   * The div with class container wraps the entire application content, providing central alignment and basic styling.
   * Top Buttons: Includes buttons for Import MSRA from CSV, Export MSRA to CSV (top), Convert Excel CSV, Help, and Download CSV Template. An input type="file" (hidden) is also present for CSV import.
   * Unsaved Changes Warning: A div with id="unsaved-changes-warning" is initially hidden and appears when data is modified.
   * Two-Column Layout (two-column-layout):
     * MSRA Information Section (msra-info-section): Collects general MSRA details like Job Name, Prepared By, Contractor Name, Contractor Person Name, and Date of Preparation.
     * Type of Work Checkboxes: A checkbox-group allowing users to select various types of work (e.g., Height, Hot, Electrical).
     * PPEs to be Used Checkboxes: Another checkbox-group for selecting required Personal Protective Equipment (PPEs).
     * Tasks & Risks Input Section (task-input-section): Contains input fields for Task Description and buttons to Add New Task and Add Risk to Current Task, along with Risk Description, Initial Likelihood, and Initial Severity selectors.
 * MSRA Overview Section (msra-overview-section):
   * Header: Displays a placeholder logo (Schneider Electric) and the main title "Method Statement and Risk Assessment (MSRA)".
   * MSRA Summary Info (msra-summary-info): A div where the entered MSRA information (Job Name, Prepared By, etc.) is dynamically displayed for review.
   * Table Container (table-container): Wraps the main MSRA table, enabling horizontal scrolling for wider content.
   * MSRA Table (msraTable): The core of the application, displaying tasks and their associated risks in a tabular format. It includes columns for:
     * Task ID, Task Description
     * Risk ID, Risk Description
     * Initial Likelihood, Initial Severity, Initial Risk Level
     * Countermeasure
     * New Likelihood, New Severity, Residual Risk Level
     * Risk Actions (Edit Risk, Add/Edit CM, Delete Risk)
     * Task Actions (Add Risk, Insert Task Below)
   * Bottom Export Buttons: Provides buttons to Export MSRA to CSV, Export MSRA to Image, and Export MSRA to PDF.
 * Modals:
   * Countermeasure Modal (countermeasureModal): A hidden overlay that appears when adding or editing countermeasures. It allows input for Countermeasure / Control Measures, New Likelihood, and New Severity.
   * Add Risk to Existing Task Modal (addRiskToExistingTaskModal): A hidden overlay for adding new risks to an already existing task, similar fields to the main risk input.
   * Edit Risk Modal (editRiskModal): A hidden overlay for modifying an existing risk's description, initial likelihood, and initial severity.
   * Insert Task Modal (insertTaskModal): A hidden overlay for inserting a new task at a specific position in the task list.
   * Help Modal (helpModal): A hidden overlay providing information about what an MSRA is and the "THRIVE" methodology, including a placeholder YouTube video.
 * JavaScript Link: Links to script.js at the end of the <body> for optimal loading performance.
2. style.css
This file contains all the Cascading Style Sheets (CSS) rules that define the visual presentation of the MSRA Tool. It controls the layout, colors, fonts, spacing, and responsiveness of the web page elements.
Key Styling Areas:
 * General Body and Container: Sets basic font, background color, margins, and centers the main container. The container has a maximum width, padding, rounded corners, and a box shadow.
 * Headings (h1, h2): Defines text color, alignment, and bottom margin for titles.
 * Sections: Applies common padding, border, and border-radius to all major content sections. Specific id selectors (#msra-info-section, #task-input-section, #msra-overview-section) apply unique background colors to differentiate sections visually.
 * Form Elements (.form-group, input, select, textarea): Styles labels, input fields, select boxes, and text areas for consistent appearance, including width, padding, borders, border-radius, and font size.
 * Buttons: Defines a consistent style for all buttons, including background color, text color, padding, borders, border-radius, cursor, font size, and hover/active effects (transition and transform for subtle animations).
   * Specific id selectors provide different color schemes for various button actions (e.g., addTaskBtn for green, delete-btn for red, exportImageBtn for orange).
 * Table Styling (table, th, td):
   * table-container: Enables horizontal scrolling (overflow-x: auto) for the table, ensuring it remains usable on smaller screens.
   * table: Sets full width, collapses borders, and defines a minimum width to prevent columns from becoming too narrow.
   * th, td: Sets padding, text alignment, font size, and default white-space: nowrap to prevent text wrapping.
   * td.wrap-text: An optional class to allow text wrapping in specific table cells.
   * th: Styles table headers with a light background, bold text, and specific text color.
   * tr:nth-child(even) and tr:hover: Provide alternating row backgrounds and hover effects for better readability.
 * Risk Level Classes: Defines distinct background and text colors for different risk levels (low, medium, high, critical, unknown) to visually highlight risk severity in the table.
 * Table Action Buttons: Styles the "Delete", "Countermeasure", "Add Risk", "Insert Task", and "Edit Risk" buttons within the table cells, often with smaller padding and font size for compactness.
 * Task Header Row (.task-header-row): Special styling for the task description rows, including background, bold text, and flexible layout (display: flex) to accommodate the task title and action buttons.
 * Modal Styles (.modal-overlay, .modal-content): Styles the pop-up modals, including the dark overlay background, modal content box (white background, padding, rounded corners, shadow), and a close button (modal-close-btn).
 * MSRA Info Summary: Styles the dynamic information display section.
 * Unsaved Changes Warning: Styles the warning message for unsaved data.
 * Logo and Title (#msra-overview-header): Uses flexbox for alignment of the logo and title in the overview section.
 * Two-Column Layout: Employs flexbox (display: flex) to create a two-column layout for the MSRA Information and Task Input sections on larger screens.
 * Checkbox Group Styling: Provides visual grouping and layout for the "Type of Work" and "PPEs to be Used" checkboxes.
 * Help Modal Specific Styles: Adds more specific styling for the content within the help modal, including video embedding.
 * Signature Block Styles: Defines the layout and appearance of the signature lines at the bottom of the exported PDF/image.
 * Responsive Adjustments (@media (max-width: 768px)): Includes media queries to adapt the layout for smaller screens (e.g., mobile devices). This involves adjusting padding, making buttons full-width, forcing horizontal scrolling for tables, and stacking two-column layouts into a single column.
3. script.js
This file contains all the JavaScript logic that makes the MSRA Tool interactive and functional. It handles data management, user input, dynamic table rendering, modal interactions, and export functionalities.
Key Functionalities and Variables:
 * Global Variables & Constants:
   * jsPDF: Imported from the jsPDF library.
   * DOM Element References: Stores references to various HTML elements using document.getElementById() and document.querySelector() for efficient manipulation.
   * msraData: A central JavaScript object that holds all the application's data, including jobInfo (general MSRA details and checkboxes) and an array of tasks. Each task object contains its id, description, and an array of risks.
   * currentTaskId: Tracks the ID of the currently active task for adding risks.
   * taskIdCounter, riskIdCounter: Used to generate unique IDs for new tasks and risks.
   * dataModified: A boolean flag (true/false) that indicates if any changes have been made to the msraData since the last save/import, used for the "unsaved changes" warning.
   * editingCountermeasureRisk: Temporarily stores the risk object being edited in the Countermeasure modal.
   * editingRiskObject: Temporarily stores the risk object being edited in the Edit Risk modal.
   * targetTaskIdForNewRisk: Stores the ID of the task to which a new risk should be added via the "Add Risk to Task" modal.
   * insertIndexForNewTask: Stores the array index where a new task should be inserted via the "Insert Task Below" modal.
 * Risk Matrix Logic:
   * riskMatrixLevels: A 2D array defining the risk levels (LOW, MEDIUM, HIGH, CRITICAL) based on Likelihood (rows 1-5) and Severity (columns 1-5).
   * getRiskLevel(likelihoodValue, severityValue): A crucial function that calculates the risk level (e.g., "HIGH", "MEDIUM") based on numerical likelihood and severity inputs using the riskMatrixLevels table. It returns "Unknown" for invalid inputs.
   * getRiskLevelClass(level): Returns the corresponding CSS class name (e.g., "risk-level-high") for a given risk level, used to apply styling to table cells.
 * Helper Functions for Data Conversion:
   * getLikelihoodLabelFromValue(value): Converts a numerical likelihood value (e.g., 5) to its letter label (e.g., A).
   * getImpactLabelFromValue(value): Converts a numerical impact value (e.g., 1) to its numerical label (e.g., 1).
   * getLikelihoodValueFromLabel(label): Converts a likelihood letter label (e.g., A) back to its numerical value (e.g., 5).
   * getImpactValueFromLabel(label): Converts an impact numerical label (e.g., 1) back to its numerical value (e.g., 1).
 * Rendering Functions:
   * renderMSRAInfoSummary(): Updates the input fields in the "MSRA Information" section and dynamically populates the msra-summary-info div with the current msraData.jobInfo. It also syncs the state of the "Type of Work" and "PPEs to be Used" checkboxes.
   * renderMSRATable(): Clears and re-renders the entire #msraTable based on the msraData.tasks array. It iterates through each task and its associated risks, creating table rows, cells, and dynamically adding action buttons (Delete, Add/Edit Countermeasure, Edit Risk, Add Risk to Task, Insert Task Below). It also applies the correct CSS classes for risk level highlighting.
 * Modal Management Functions:
   * openCountermeasureModal(taskId, riskId): Populates and displays the Countermeasure modal with data from the selected risk.
   * closeCountermeasureModal(): Hides the Countermeasure modal and clears its input fields.
   * openAddRiskToExistingTaskModal(taskId): Populates and displays the Add Risk to Existing Task modal for a specific task.
   * closeAddRiskToExistingTaskModal(): Hides the Add Risk to Existing Task modal and clears its input fields.
   * openEditRiskModal(taskId, riskId): Populates and displays the Edit Risk modal with data from the selected risk.
   * closeEditRiskModal(): Hides the Edit Risk modal and clears its input fields.
   * openInsertTaskModal(index): Populates and displays the Insert Task modal, indicating where the new task will be inserted.
   * closeInsertTaskModal(): Hides the Insert Task modal and clears its input fields.
 * Event Listeners:
   * Input Fields (MSRA Info): input event listeners on all jobInfo fields (jobName, preparedBy, etc.) and change listeners on checkboxes update msraData.jobInfo, re-render the summary, and set dataModified to true.
   * addTaskBtn: On click, validates task description, creates a new task object, adds it to msraData.tasks, sets currentTaskId, and re-renders the table.
   * addRiskToTaskBtn: On click, validates risk inputs, creates a new risk object, adds it to the risks array of the currentTaskId, and re-renders the table.
   * saveCountermeasureBtn: On click, retrieves values from the Countermeasure modal, updates the editingCountermeasureRisk object, calculates the residualLevel, and re-renders the table.
   * saveAddRiskModalBtn: On click, creates a new risk and adds it to the task specified by targetTaskIdForNewRisk.
   * saveEditRiskModalBtn: On click, updates the properties of editingRiskObject with new values from the Edit Risk modal.
   * saveInsertTaskModalBtn: On click, creates a new task and inserts it into the msraData.tasks array at the specified insertIndexForNewTask.
   * Modal Close Buttons & Overlay Clicks: Event listeners for click events on modal close buttons (x) and the modal overlay itself (to close when clicking outside the content).
   * helpBtn: Displays the help modal.
   * exportCsvBtn (top & bottom): Calls performCsvExport() to generate and download a CSV file of the current MSRA data.
   * downloadTemplateBtn: Calls downloadCsvTemplate() to provide a blank CSV template.
   * exportImageBtn: Uses html2canvas to capture the #msra-overview-section as an image (PNG) and triggers a download. It temporarily hides interactive elements before capture and restores them afterward.
   * exportPdfBtn: Uses html2canvas and jsPDF to capture the #msra-overview-section as a multi-page PDF. It also temporarily adds a signature block for the export, hides interactive elements, and restores them after generation.
   * importCsvBtn: Triggers a click on the hidden csvFileInput to allow user to select a CSV file.
   * csvFileInput change event: Reads the selected CSV file, calls parseAndImportCsv() to process the content, and updates the msraData. Includes error handling.
   * convertExcelCsvBtn: Provides a utility to convert CSV files exported from Excel (which might have formatting issues) back to the format expected by the tool's import function. This involves robust CSV parsing, date format conversion, and re-quoting fields.
   * beforeunload event: Displays a browser-level warning if the user attempts to leave the page while dataModified is true or if there's any data in msraData.
 * parseAndImportCsv(csvContent):
   * This function is responsible for taking the raw CSV content and parsing it into the msraData structure.
   * It handles both the fixed MSRA header information (Job Name, Prepared By, etc.) and the dynamic task and risk table data.
   * It uses a custom cleanCsvPart helper to properly unquote and unescape CSV fields.
   * It dynamically finds the main table header row to start parsing task and risk data.
   * It reconstructs task and risk objects, ensuring likelihood and impact values are correctly parsed (handling cases where they might be missing or non-numeric).
   * It recalculates initial and residual risk levels.
   * It sorts tasks and risks by their IDs and re-assigns sequential task IDs for consistency after import.
   * Updates taskIdCounter and riskIdCounter to ensure new additions don't conflict with imported IDs.
 * convertCsvFormat(csvContent):
   * This function is specifically designed to take a CSV file that might have been opened and saved in Excel (which can alter formatting, especially quoting and date formats) and convert it back to a format compatible with the tool's parseAndImportCsv function.
   * It parses each line, removes trailing commas often added by Excel, converts date formats (DD-MM-YYYY to YYYY-MM-DD), and ensures that fields that should be quoted (like descriptions and countermeasures) are properly re-quoted, and internal double quotes are correctly escaped ("").
   * It uses a more robust parseCsvLine helper that correctly handles commas and escaped quotes within fields.
 * updateUnsavedChangesWarning(): Controls the visibility of the yellow warning message at the top of the page based on the dataModified flag and whether any data exists in msraData.
 * Initialization (DOMContentLoaded):
   * When the HTML document is fully loaded, the DOMContentLoaded event listener triggers the initial calls to renderMSRAInfoSummary() and renderMSRATable() to display any default or pre-loaded data, and updateUnsavedChangesWarning() to set the initial warning state.
Getting Started
To use this MSRA Tool locally:
 * Save the index.html content as index.html.
 * Save the style.css content as style.css in the same directory.
 * Save the script.js content as script.js in the same directory.
 * Open index.html in your web browser.
This organized structure makes it easier to navigate, understand, and modify specific parts of the application.
