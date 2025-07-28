// Import jsPDF from the UMD bundle
const { jsPDF } = window.jspdf;

document.addEventListener('DOMContentLoaded', () => {
    // MSRA Information fields
    const jobNameInput = document.getElementById('jobName');
    const preparedByInput = document.getElementById('preparedBy');
    const contractorNameInput = document.getElementById('contractorName');
    const contractorPersonNameInput = document.getElementById('contractorPersonName');
    const datePreparedInput = document.getElementById('datePrepared');
    const msraSummaryInfoDiv = document.getElementById('msra-summary-info');

    // New: Checkbox elements - Corrected selectors to use new IDs
    const workTypeCheckboxesContainer = document.getElementById('type-of-work-checkboxes');
    const workTypeCheckboxes = workTypeCheckboxesContainer.querySelectorAll('input[type="checkbox"]');

    const ppeCheckboxesContainer = document.getElementById('ppes-to-be-used-checkboxes');
    const ppeCheckboxes = ppeCheckboxesContainer.querySelectorAll('input[type="checkbox"]');


    // Task input fields
    const taskDescriptionInput = document.getElementById('taskDescription');
    const addTaskBtn = document.getElementById('addTaskBtn');

    // Risk input fields (Initial Risk)
    const riskDescriptionInput = document.getElementById('riskDescription');
    const likelihoodSelect = document.getElementById('likelihood');
    const impactSelect = document.getElementById('impact');
    const addRiskToTaskBtn = document.getElementById('addRiskToTaskBtn');

    // Modal elements (Countermeasure)
    const countermeasureModal = document.getElementById('countermeasureModal');
    const closeModalBtn = document.getElementById('closeModalBtn');
    const modalRiskIdSpan = document.getElementById('modalRiskId');
    const modalCountermeasureInput = document.getElementById('modalCountermeasure');
    const modalNewLikelihoodSelect = document.getElementById('modalNewLikelihood');
    const modalNewImpactSelect = document.getElementById('modalNewImpact');
    const saveCountermeasureBtn = document.getElementById('saveCountermeasureBtn');

    // Modal elements (Add Risk to Existing Task)
    const addRiskToExistingTaskModal = document.getElementById('addRiskToExistingTaskModal');
    const closeAddRiskModalBtn = document.getElementById('closeAddRiskModalBtn');
    const addRiskModalTaskIdSpan = document.getElementById('addRiskModalTaskId');
    const addRiskModalDescriptionInput = document.getElementById('addRiskModalDescription');
    const addRiskModalLikelihoodSelect = document.getElementById('addRiskModalLikelihood');
    const addRiskModalImpactSelect = document.getElementById('addRiskModalImpact');
    const saveAddRiskModalBtn = document.getElementById('saveAddRiskModalBtn');
    let targetTaskIdForNewRisk = null; // To store which task to add risk to

    // Modal elements (Edit Risk)
    const editRiskModal = document.getElementById('editRiskModal');
    const closeEditRiskModalBtn = document.getElementById('closeEditRiskModalBtn');
    const editRiskModalRiskIdSpan = document.getElementById('editRiskModalRiskId');
    const editRiskModalDescriptionInput = document.getElementById('editRiskModalDescription');
    const editRiskModalLikelihoodSelect = document.getElementById('editRiskModalLikelihood');
    const editRiskModalImpactSelect = document.getElementById('editRiskModalImpact');
    const saveEditRiskModalBtn = document.getElementById('saveEditRiskModalBtn');
    let editingRiskObject = null; // Stores the risk object being edited in the edit risk modal

    // Modal elements (Insert Task)
    const insertTaskModal = document.getElementById('insertTaskModal');
    const closeInsertTaskModalBtn = document.getElementById('closeInsertTaskModalBtn');
    const insertTaskModalContextP = document.getElementById('insertTaskModalContext');
    const insertTaskModalDescriptionInput = document.getElementById('insertTaskModalDescription');
    const saveInsertTaskModalBtn = document.getElementById('saveInsertTaskModalBtn');
    let insertIndexForNewTask = null; // To store the index where the new task should be inserted

    // Table body
    const msraTableBody = document.querySelector('#msraTable tbody');
    const exportCsvBtn = document.getElementById('exportCsvBtn'); // Bottom CSV export
    const exportImageBtn = document.getElementById('exportImageBtn');
    const exportPdfBtn = document.getElementById('exportPdfBtn');
    const importCsvBtn = document.getElementById('importCsvBtn');
    const exportCsvBtnTop = document.getElementById('exportCsvBtnTop'); // Top CSV export
    const csvFileInput = document.getElementById('csvFileInput');
    const msraOverviewSection = document.getElementById('msra-overview-section');
    const tableContainer = document.querySelector('.table-container'); // Get the table container
    const helpBtn = document.getElementById('helpBtn'); // Help button
    const helpModal = document.getElementById('helpModal'); // Help modal
    const closeHelpModalBtn = document.getElementById('closeHelpModalBtn'); // Close help modal button
    const unsavedChangesWarning = document.getElementById('unsaved-changes-warning'); // Unsaved changes warning
    const convertExcelCsvBtn = document.getElementById('convertExcelCsvBtn'); // New converter button

    // Data structures
    let msraData = {
        jobInfo: {
            jobName: '',
            preparedBy: '',
            contractorName: '',
            contractorPersonName: '',
            datePrepared: '',
            typeOfWork: [], // New: Array to store selected work types
            ppesUsed: []    // New: Array to store selected PPEs
        },
        tasks: []
    };
    let currentTaskId = null;
    let taskIdCounter = 1;
    let riskIdCounter = 1;
    let dataModified = false; // Flag to track if data has been modified

    // --- START CUSTOMIZATION AREA: RISK MATRIX LOGIC ---
    const riskMatrixLevels = [
        ['LOW',   'LOW',    'LOW',    'MEDIUM',  'MEDIUM'], // Likelihood 1 (E)
        ['LOW',   'LOW',    'MEDIUM', 'MEDIUM',  'HIGH'],   // Likelihood 2 (D)
        ['LOW',   'MEDIUM', 'MEDIUM', 'HIGH',    'HIGH'],   // Likelihood 3 (C)
        ['MEDIUM', 'MEDIUM', 'HIGH',   'HIGH',    'CRITICAL'], // Likelihood 4 (B)
        ['MEDIUM', 'HIGH',   'HIGH',   'CRITICAL','CRITICAL']  // Likelihood 5 (A)
    ];

    const getRiskLevel = (likelihoodValue, severityValue) => {
        // Handle null or non-numeric inputs gracefully
        if (likelihoodValue === null || isNaN(likelihoodValue) ||
            severityValue === null || isNaN(severityValue)) {
            return 'Unknown';
        }

        const likelihoodIndex = likelihoodValue - 1;
        const severityIndex = severityValue - 1;

        if (likelihoodIndex >= 0 && likelihoodIndex < riskMatrixLevels.length &&
            severityIndex >= 0 && severityIndex < riskMatrixLevels[0].length) {
            return riskMatrixLevels[likelihoodIndex][severityIndex];
        }
        return 'Unknown';
    };

    const getRiskLevelClass = (level) => {
        switch (level.toUpperCase()) {
            case 'LOW': return 'risk-level-low';
            case 'MEDIUM': return 'risk-level-medium';
            case 'HIGH': return 'risk-level-high';
            case 'CRITICAL': return 'risk-level-critical';
            default: return 'risk-level-unknown'; // Always return a non-empty string
        }
    };
    // --- END CUSTOMIZATION AREA ---

    // Helper to get option label from value
    const getLikelihoodLabelFromValue = (value) => {
        // Using a map for efficiency and robustness
        const likelihoodMap = {
            '5': 'A', '4': 'B', '3': 'C', '2': 'D', '1': 'E'
        };
        return likelihoodMap[value] || '';
    };

    const getImpactLabelFromValue = (value) => {
        // Using a map for efficiency and robustness
        const impactMap = {
            '1': '1', '2': '2', '3': '3', '4': '4', '5': '5'
        };
        return impactMap[value] || '';
    };


    // Function to render MSRA Information Summary
    const renderMSRAInfoSummary = () => {
        const info = msraData.jobInfo;
        // Update input fields with current data
        jobNameInput.value = info.jobName;
        preparedByInput.value = info.preparedBy;
        contractorNameInput.value = info.contractorName;
        contractorPersonNameInput.value = info.contractorPersonName;
        datePreparedInput.value = info.datePrepared;

        // Update checkbox states
        workTypeCheckboxes.forEach(checkbox => {
            checkbox.checked = info.typeOfWork.includes(checkbox.value);
        });
        ppeCheckboxes.forEach(checkbox => {
            checkbox.checked = info.ppesUsed.includes(checkbox.value);
        });


        msraSummaryInfoDiv.innerHTML = `
            <p><strong>Job Name:</strong> ${info.jobName || 'N/A'}</p>
            <p><strong>Prepared By:</strong> ${info.preparedBy || 'N/A'}</p>
            <p><strong>Contractor Name:</strong> ${info.contractorName || 'N/A'}</p>
            <p><strong>Contractor Person Name:</strong> ${info.contractorPersonName || 'N/A'}</p>
            <p><strong>Date of Preparation:</strong> ${info.datePrepared || 'N/A'}</p>
            <p><strong>Type of Work:</strong> ${info.typeOfWork.length > 0 ? info.typeOfWork.join(', ') : 'N/A'}</p>
            <p><strong>PPEs to be Used:</strong> ${info.ppesUsed.length > 0 ? info.ppesUsed.join(', ') : 'N/A'}</p>
        `;
    };


    // Function to render the entire MSRA table (tasks and their risks)
    const renderMSRATable = () => {
        msraTableBody.innerHTML = ''; // Clear existing rows

        msraData.tasks.forEach((task, taskIndex) => { // Added taskIndex
            // Add a header row for the task
            const taskHeaderRow = msraTableBody.insertRow();
            taskHeaderRow.classList.add('task-header-row');
            // Create a single cell that spans all columns
            const taskHeaderCell = taskHeaderRow.insertCell();
            taskHeaderCell.colSpan = document.querySelector('#msraTable thead tr').children.length;
            taskHeaderCell.innerHTML = `
                <span class="task-title">Task ${task.id}: ${task.description}</span>
                <div class="task-actions">
                    <button class="add-risk-btn" data-task-id="${task.id}">Add Risk</button>
                    <button class="insert-task-btn" data-insert-index="${taskIndex}">Insert Task Below</button>
                </div>
            `;

            if (task.risks.length === 0) {
                const noRiskRow = msraTableBody.insertRow();
                const noRiskCell = noRiskRow.insertCell();
                noRiskCell.colSpan = document.querySelector('#msraTable thead tr').children.length;
                noRiskCell.textContent = 'No risks added for this task yet.';
                noRiskCell.style.fontStyle = 'italic';
                noRiskCell.style.textAlign = 'center';
                noRiskCell.style.padding = '10px';
            } else {
                task.risks.forEach(risk => {
                    const row = msraTableBody.insertRow();
                    row.setAttribute('data-task-id', task.id);
                    row.setAttribute('data-risk-id', risk.id);

                    // Task ID
                    const taskIdCell = row.insertCell();
                    taskIdCell.setAttribute('data-label', 'Task ID');
                    taskIdCell.textContent = task.id;

                    // Task Description
                    const taskDescriptionCell = row.insertCell();
                    taskDescriptionCell.setAttribute('data-label', 'Task Description');
                    taskDescriptionCell.textContent = task.description;

                    // Risk ID
                    const riskIdCell = row.insertCell();
                    riskIdCell.setAttribute('data-label', 'Risk ID');
                    riskIdCell.textContent = risk.id;

                    // Risk Description
                    const riskDescriptionCell = row.insertCell();
                    riskDescriptionCell.setAttribute('data-label', 'Risk Description');
                    riskDescriptionCell.textContent = risk.description;

                    // Initial Likelihood
                    const initialLikelihoodCell = row.insertCell();
                    initialLikelihoodCell.setAttribute('data-label', 'Initial Likelihood');
                    initialLikelihoodCell.textContent = risk.likelihoodLabel;

                    // Initial Severity
                    const initialSeverityCell = row.insertCell();
                    initialSeverityCell.setAttribute('data-label', 'Initial Severity');
                    initialSeverityCell.textContent = risk.impactLabel;

                    // Initial Risk Level
                    const initialLevelClass = getRiskLevelClass(risk.initialLevel);
                    const initialLevelCell = row.insertCell();
                    initialLevelCell.setAttribute('data-label', 'Initial Risk Level');
                    initialLevelCell.textContent = risk.initialLevel;
                    if (initialLevelClass) { // Only add if not empty
                        initialLevelCell.classList.add(initialLevelClass);
                    }

                    // Countermeasure
                    const countermeasureCell = row.insertCell();
                    countermeasureCell.setAttribute('data-label', 'Countermeasure');
                    countermeasureCell.textContent = risk.countermeasure || 'N/A';

                    // New Likelihood
                    const newLikelihoodCell = row.insertCell();
                    newLikelihoodCell.setAttribute('data-label', 'New Likelihood');
                    newLikelihoodCell.textContent = risk.newLikelihoodLabel || 'N/A';

                    // New Severity
                    const newSeverityCell = row.insertCell();
                    newSeverityCell.setAttribute('data-label', 'New Severity');
                    newSeverityCell.textContent = risk.newImpactLabel || 'N/A';

                    // Residual Risk Level
                    const residualLevelClass = getRiskLevelClass(risk.residualLevel);
                    const residualLevelCell = row.insertCell();
                    residualLevelCell.setAttribute('data-label', 'Residual Risk Level');
                    residualLevelCell.textContent = risk.residualLevel || 'N/A';
                    if (residualLevelClass) { // Only add if not empty
                        residualLevelCell.classList.add(residualLevelClass);
                    }

                    // Risk Actions (for risk-level actions)
                    const riskActionsCell = row.insertCell();
                    riskActionsCell.setAttribute('data-label', 'Risk Actions');
                    const editRiskBtn = document.createElement('button'); // New Edit Risk button
                    editRiskBtn.textContent = 'Edit Risk';
                    editRiskBtn.classList.add('edit-risk-btn');
                    editRiskBtn.onclick = () => openEditRiskModal(task.id, risk.id);
                    riskActionsCell.appendChild(editRiskBtn);

                    const addCountermeasureBtn = document.createElement('button');
                    addCountermeasureBtn.textContent = risk.countermeasure ? 'Edit CM' : 'Add CM';
                    addCountermeasureBtn.classList.add('countermeasure-btn');
                    addCountermeasureBtn.onclick = () => openCountermeasureModal(task.id, risk.id);
                    riskActionsCell.appendChild(addCountermeasureBtn);

                    const deleteButton = document.createElement('button');
                    deleteButton.textContent = 'Delete Risk';
                    deleteButton.classList.add('delete-btn');
                    deleteButton.onclick = () => {
                        const taskIndex = msraData.tasks.findIndex(t => t.id === task.id);
                        if (taskIndex !== -1) {
                            msraData.tasks[taskIndex].risks = msraData.tasks[taskIndex].risks.filter(r => r.id !== risk.id);
                            renderMSRATable();
                            dataModified = true; // Data modified
                            updateUnsavedChangesWarning(); // Update warning message
                        }
                    };
                    riskActionsCell.appendChild(deleteButton);

                    // Task Actions (newly added column for task-level actions, empty for risk rows)
                    row.insertCell().setAttribute('data-label', 'Task Actions');
                });
            }
        });

        // Attach event listeners to newly created buttons
        document.querySelectorAll('.add-risk-btn').forEach(button => {
            button.addEventListener('click', (e) => {
                const taskId = parseInt(e.target.dataset.taskId);
                openAddRiskToExistingTaskModal(taskId);
            });
        });

        document.querySelectorAll('.insert-task-btn').forEach(button => {
            button.addEventListener('click', (e) => {
                const insertIndex = parseInt(e.target.dataset.insertIndex);
                openInsertTaskModal(insertIndex);
            });
        });
    };

    // Function to open the Countermeasure Modal
    const openCountermeasureModal = (taskId, riskId) => {
        const task = msraData.tasks.find(t => t.id === taskId);
        if (!task) return;
        const risk = task.risks.find(r => r.id === riskId);
        if (!risk) return;

        editingCountermeasureRisk = risk; // Store the risk object being edited for countermeasure

        modalRiskIdSpan.textContent = risk.id;
        modalCountermeasureInput.value = risk.countermeasure || '';
        modalNewLikelihoodSelect.value = risk.newLikelihood || '';
        modalNewImpactSelect.value = risk.newImpact || '';

        countermeasureModal.style.display = 'flex'; // Show the modal
    };

    // Function to close the Countermeasure Modal
    const closeCountermeasureModal = () => {
        countermeasureModal.style.display = 'none';
        modalCountermeasureInput.value = '';
        modalNewLikelihoodSelect.value = '';
        modalNewImpactSelect.value = '';
        editingCountermeasureRisk = null;
    };

    // Function to open Add Risk to Existing Task Modal
    const openAddRiskToExistingTaskModal = (taskId) => {
        targetTaskIdForNewRisk = taskId;
        addRiskModalTaskIdSpan.textContent = taskId;
        addRiskModalDescriptionInput.value = '';
        addRiskModalLikelihoodSelect.value = '';
        addRiskModalImpactSelect.value = '';
        addRiskToExistingTaskModal.style.display = 'flex';
    };

    // Function to close Add Risk to Existing Task Modal
    const closeAddRiskToExistingTaskModal = () => {
        addRiskToExistingTaskModal.style.display = 'none';
        targetTaskIdForNewRisk = null;
        addRiskModalDescriptionInput.value = '';
        addRiskModalLikelihoodSelect.value = '';
        addRiskModalImpactSelect.value = '';
    };

    // Function to open Edit Risk Modal
    const openEditRiskModal = (taskId, riskId) => {
        const task = msraData.tasks.find(t => t.id === taskId);
        if (!task) return;
        const risk = task.risks.find(r => r.id === riskId);
        if (!risk) return;

        editingRiskObject = risk; // Store the risk object being edited
        
        editRiskModalRiskIdSpan.textContent = risk.id;
        editRiskModalDescriptionInput.value = risk.description;
        editRiskModalLikelihoodSelect.value = risk.likelihood;
        editRiskModalImpactSelect.value = risk.impact;

        editRiskModal.style.display = 'flex';
    };

    // Function to close Edit Risk Modal
    const closeEditRiskModal = () => {
        editRiskModal.style.display = 'none';
        editRiskModalDescriptionInput.value = '';
        editRiskModalLikelihoodSelect.value = '';
        editRiskModalImpactSelect.value = '';
        editingRiskObject = null;
    };

    // Function to open Insert Task Modal
    const openInsertTaskModal = (index) => {
        insertIndexForNewTask = index;
        if (index === msraData.tasks.length) {
            insertTaskModalContextP.textContent = "Inserting a new task at the end.";
        } else {
            insertTaskModalContextP.textContent = `Inserting new task below Task ${msraData.tasks[index].id}: "${msraData.tasks[index].description}"`;
        }
        insertTaskModalDescriptionInput.value = '';
        insertTaskModal.style.display = 'flex';
    };

    // Function to close Insert Task Modal
    const closeInsertTaskModal = () => {
        insertTaskModal.style.display = 'none';
        insertIndexForNewTask = null;
        insertTaskModalDescriptionInput.value = '';
    };


    // Event listener for "Add New Task" button (main input section)
    addTaskBtn.addEventListener('click', () => {
        const description = taskDescriptionInput.value.trim();
        if (!description) {
            // Replaced alert with custom message box or modal if available
            const message = 'Please enter a Task Description.';
            alert(message);
            return;
        }

        const newTask = {
            id: taskIdCounter++,
            description: description,
            risks: []
        };
        msraData.tasks.push(newTask);
        currentTaskId = newTask.id; // Set this as the current task for adding risks

        renderMSRATable(); // Re-render to show the new task
        dataModified = true; // Data modified
        updateUnsavedChangesWarning(); // Update warning message

        // Clear task input and focus on risk input
        taskDescriptionInput.value = '';
        riskDescriptionInput.focus();

        // Replaced alert with custom message box or modal if available
        const message = `New Task "${description}" added. You can now add risks to this task.`;
        alert(message);
    });

    // Event listener for "Add Risk to Current Task" button (main input section)
    addRiskToTaskBtn.addEventListener('click', () => {
        if (currentTaskId === null) {
            // Replaced alert with custom message box or modal if available
            const message = 'Please add a new task first before adding risks.';
            alert(message);
            return;
        }

        const description = riskDescriptionInput.value.trim();
        const likelihoodValue = parseInt(likelihoodSelect.value);
        const impactValue = parseInt(impactSelect.value);

        if (!description || isNaN(likelihoodValue) || isNaN(impactValue)) {
            // Replaced alert with custom message box or modal if available
            const message = 'Please fill in Risk Description, Initial Likelihood, and Initial Severity.';
            alert(message);
            return;
        }

        const initialLevel = getRiskLevel(likelihoodValue, impactValue);

        const newRisk = {
            id: riskIdCounter++,
            description: description,
            likelihood: likelihoodValue,
            likelihoodLabel: likelihoodSelect.options[likelihoodSelect.selectedIndex].text.split(' - ')[0],
            impact: impactValue,
            impactLabel: impactSelect.options[impactSelect.selectedIndex].text.split(' - ')[0],
            initialLevel: initialLevel,
            countermeasure: '',
            newLikelihood: null,
            newLikelihoodLabel: '',
            newImpact: null,
            newImpactLabel: '',
            residualLevel: ''
        };

        const currentTask = msraData.tasks.find(task => task.id === currentTaskId);
        if (currentTask) {
            currentTask.risks.push(newRisk);
            renderMSRATable();
            dataModified = true; // Data modified
            updateUnsavedChangesWarning(); // Update warning message
        } else {
            // Replaced alert with custom message box or modal if available
            const message = 'Error: Current task not found. Please add a new task.';
            alert(message);
        }

        riskDescriptionInput.value = '';
        likelihoodSelect.value = '';
        impactSelect.value = '';
    });

    // Event listener for saving countermeasure from modal
    saveCountermeasureBtn.addEventListener('click', () => {
        if (!editingCountermeasureRisk) {
            // Replaced alert with custom message box or modal if available
            const message = 'No risk selected for editing.';
            alert(message);
            return;
        }

        const countermeasure = modalCountermeasureInput.value.trim();
        const newLikelihoodValue = parseInt(modalNewLikelihoodSelect.value);
        const newImpactValue = parseInt(modalNewImpactSelect.value);

        if (!countermeasure || isNaN(newLikelihoodValue) || isNaN(newImpactValue)) {
            // Replaced alert with custom message box or modal if available
            const message = 'Please fill in Countermeasure, New Likelihood, and New Severity in the modal.';
            alert(message);
            return;
        }

        editingCountermeasureRisk.countermeasure = countermeasure;
        editingCountermeasureRisk.newLikelihood = newLikelihoodValue;
        editingCountermeasureRisk.newLikelihoodLabel = modalNewLikelihoodSelect.options[modalNewLikelihoodSelect.selectedIndex].text.split(' - ')[0];
        editingCountermeasureRisk.newImpact = newImpactValue;
        editingCountermeasureRisk.newImpactLabel = modalNewImpactSelect.options[modalNewImpactSelect.selectedIndex].text.split(' - ')[0];
        editingCountermeasureRisk.residualLevel = getRiskLevel(newLikelihoodValue, newImpactValue);

        renderMSRATable();
        closeCountermeasureModal();
        dataModified = true; // Data modified
        updateUnsavedChangesWarning(); // Update warning message
    });

    // Event listener for saving new risk from "Add Risk to Existing Task" modal
    saveAddRiskModalBtn.addEventListener('click', () => {
        if (targetTaskIdForNewRisk === null) {
            // Replaced alert with custom message box or modal if available
            const message = 'Error: No target task selected for adding risk.';
            alert(message);
            return;
        }

        const description = addRiskModalDescriptionInput.value.trim();
        const likelihoodValue = parseInt(addRiskModalLikelihoodSelect.value);
        const impactValue = parseInt(addRiskModalImpactSelect.value);

        if (!description || isNaN(likelihoodValue) || isNaN(impactValue)) {
            // Replaced alert with custom message box or modal if available
            const message = 'Please fill in Risk Description, Initial Likelihood, and Initial Severity in the modal.';
            alert(message);
            return;
        }

        const initialLevel = getRiskLevel(likelihoodValue, impactValue);

        const newRisk = {
            id: riskIdCounter++,
            description: description,
            likelihood: likelihoodValue,
            likelihoodLabel: addRiskModalLikelihoodSelect.options[addRiskModalLikelihoodSelect.selectedIndex].text.split(' - ')[0],
            impact: impactValue,
            impactLabel: addRiskModalImpactSelect.options[addRiskModalImpactSelect.selectedIndex].text.split(' - ')[0],
            initialLevel: initialLevel,
            countermeasure: '',
            newLikelihood: null,
            newLikelihoodLabel: '',
            newImpact: null,
            newImpactLabel: '',
            residualLevel: ''
        };

        const targetTask = msraData.tasks.find(task => task.id === targetTaskIdForNewRisk);
        if (targetTask) {
            targetTask.risks.push(newRisk);
            renderMSRATable();
            closeAddRiskToExistingTaskModal();
            dataModified = true; // Data modified
            updateUnsavedChangesWarning(); // Update warning message
        } else {
            // Replaced alert with custom message box or modal if available
            const message = 'Error: Target task not found for adding risk.';
            alert(message);
        }
    });

    // Event listener for saving changes from "Edit Risk" modal
    saveEditRiskModalBtn.addEventListener('click', () => {
        if (!editingRiskObject) {
            // Replaced alert with custom message box or modal if available
            const message = 'No risk selected for editing.';
            alert(message);
            return;
        }

        const description = editRiskModalDescriptionInput.value.trim();
        const likelihoodValue = parseInt(editRiskModalLikelihoodSelect.value);
        const impactValue = parseInt(editRiskModalImpactSelect.value);

        if (!description || isNaN(likelihoodValue) || isNaN(impactValue)) {
            // Replaced alert with custom message box or modal if available
            const message = 'Please fill in Risk Description, Initial Likelihood, and Initial Severity in the modal.';
            alert(message);
            return;
        }

        editingRiskObject.description = description;
        editingRiskObject.likelihood = likelihoodValue;
        editingRiskObject.likelihoodLabel = editRiskModalLikelihoodSelect.options[editRiskModalLikelihoodSelect.selectedIndex].text.split(' - ')[0];
        editingRiskObject.impact = impactValue;
        editingRiskObject.impactLabel = editRiskModalImpactSelect.options[editRiskModalImpactSelect.selectedIndex].text.split(' - ')[0];
        editingRiskObject.initialLevel = getRiskLevel(likelihoodValue, impactValue);

        renderMSRATable();
        closeEditRiskModal();
        dataModified = true; // Data modified
        updateUnsavedChangesWarning(); // Update warning message
    });


    // Event listener for saving new task from "Insert Task" modal
    saveInsertTaskModalBtn.addEventListener('click', () => {
        const description = insertTaskModalDescriptionInput.value.trim();
        if (!description) {
            // Replaced alert with custom message box or modal if available
            const message = 'Please enter a Task Description for the new task.';
            alert(message);
            return;
        }

        const newTask = {
            id: taskIdCounter++,
            description: description,
            risks: []
        };

        if (insertIndexForNewTask !== null && insertIndexForNewTask < msraData.tasks.length) {
            msraData.tasks.splice(insertIndexForNewTask + 1, 0, newTask); // Insert after the current task
        } else {
            msraData.tasks.push(newTask); // Add to end if no specific index or index is out of bounds
        }

        renderMSRATable();
        closeInsertTaskModal();
        dataModified = true; // Data modified
        updateUnsavedChangesWarning(); // Update warning message
    });


    // Event listeners for closing modals
    if (closeModalBtn) {
        closeModalBtn.addEventListener('click', closeCountermeasureModal);
    }
    if (countermeasureModal) {
        countermeasureModal.addEventListener('click', (e) => {
            if (e.target === countermeasureModal) {
                closeCountermeasureModal();
            }
        });
    }

    if (closeAddRiskModalBtn) {
        closeAddRiskModalBtn.addEventListener('click', closeAddRiskToExistingTaskModal);
    }
    if (addRiskToExistingTaskModal) {
        addRiskToExistingTaskModal.addEventListener('click', (e) => {
            if (e.target === addRiskToExistingTaskModal) {
                closeAddRiskToExistingTaskModal();
            }
        });
    }

    if (closeEditRiskModalBtn) {
        closeEditRiskModalBtn.addEventListener('click', closeEditRiskModal);
    }
    if (editRiskModal) {
        editRiskModal.addEventListener('click', (e) => {
            if (e.target === editRiskModal) {
                closeEditRiskModal();
            }
        });
    }

    if (closeInsertTaskModalBtn) {
        closeInsertTaskModalBtn.addEventListener('click', closeInsertTaskModal);
    }
    if (insertTaskModal) {
        insertTaskModal.addEventListener('click', (e) => {
            if (e.target === insertTaskModal) {
                closeInsertTaskModal();
            }
        });
    }

    // Event listener for Help button
    if (helpBtn) {
        helpBtn.addEventListener('click', () => {
            helpModal.style.display = 'flex';
        });
    }

    // Event listener for closing Help modal
    if (closeHelpModalBtn) {
        closeHelpModalBtn.addEventListener('click', () => {
            helpModal.style.display = 'none';
        });
    }
    if (helpModal) {
        helpModal.addEventListener('click', (e) => {
            if (e.target === helpModal) {
                helpModal.style.display = 'none';
            }
        });
    }


    // Event listeners for MSRA Information fields to update msraData and re-render summary
    jobNameInput.addEventListener('input', (e) => { msraData.jobInfo.jobName = e.target.value; renderMSRAInfoSummary(); dataModified = true; updateUnsavedChangesWarning(); });
    preparedByInput.addEventListener('input', (e) => { msraData.jobInfo.preparedBy = e.target.value; renderMSRAInfoSummary(); dataModified = true; updateUnsavedChangesWarning(); });
    contractorNameInput.addEventListener('input', (e) => { msraData.jobInfo.contractorName = e.target.value; renderMSRAInfoSummary(); dataModified = true; updateUnsavedChangesWarning(); });
    contractorPersonNameInput.addEventListener('input', (e) => { msraData.jobInfo.contractorPersonName = e.target.value; renderMSRAInfoSummary(); dataModified = true; updateUnsavedChangesWarning(); });
    datePreparedInput.addEventListener('input', (e) => { msraData.jobInfo.datePrepared = e.target.value; renderMSRAInfoSummary(); dataModified = true; updateUnsavedChangesWarning(); });

    // New: Event listeners for Type of Work checkboxes
    workTypeCheckboxes.forEach(checkbox => {
        checkbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                msraData.jobInfo.typeOfWork.push(e.target.value);
            } else {
                msraData.jobInfo.typeOfWork = msraData.jobInfo.typeOfWork.filter(item => item !== e.target.value);
            }
            renderMSRAInfoSummary();
            dataModified = true;
            updateUnsavedChangesWarning();
        });
    });

    // New: Event listeners for PPEs to be Used checkboxes
    ppeCheckboxes.forEach(checkbox => {
        checkbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                msraData.jobInfo.ppesUsed.push(e.target.value);
            } else {
                msraData.jobInfo.ppesUsed = msraData.jobInfo.ppesUsed.filter(item => item !== e.target.value);
            }
            renderMSRAInfoSummary();
            dataModified = true;
            updateUnsavedChangesWarning();
        });
    });


    // Event listener for "Export MSRA to CSV" button (bottom)
    exportCsvBtn.addEventListener('click', () => {
        performCsvExport();
    });

    // Event listener for "Export MSRA to CSV" button (top)
    exportCsvBtnTop.addEventListener('click', () => {
        performCsvExport();
    });

    // New: Event listener for "Download CSV Template" button
    const downloadTemplateBtn = document.getElementById('downloadTemplateBtn');
    if (downloadTemplateBtn) {
        downloadTemplateBtn.addEventListener('click', () => {
            downloadCsvTemplate();
        });
    }

    // Reusable CSV export function
    const performCsvExport = () => {
        if (msraData.tasks.length === 0 && !Object.values(msraData.jobInfo).some(value => value !== '' && value !== null && (Array.isArray(value) ? value.length > 0 : true))) {
            // Replaced alert with custom message box or modal if available
            const message = 'No MSRA data to export. Please fill in MSRA information or add tasks/risks.';
            alert(message);
            return;
        }

        let csvRows = [];

        // Add MSRA Header Information - NO QUOTES for these simple fields
        csvRows.push('Method Statement and Risk Assessment');
        csvRows.push(`Job Name:,${msraData.jobInfo.jobName}`);
        csvRows.push(`Prepared By:,${msraData.jobInfo.preparedBy}`);
        csvRows.push(`Contractor Name:,${msraData.jobInfo.contractorName}`);
        csvRows.push(`Contractor Person Name:,${msraData.jobInfo.contractorPersonName}`);
        csvRows.push(`Date of Preparation:,${msraData.jobInfo.datePrepared}`);

        // New representation for Type of Work - NO QUOTES for headers
        const allWorkTypes = Array.from(workTypeCheckboxes).map(checkbox => checkbox.value);
        const workTypeHeader = ["Type of Work:"].concat(allWorkTypes).join(',');
        csvRows.push(workTypeHeader);
        const workTypeData = ["", ...allWorkTypes.map(type => msraData.jobInfo.typeOfWork.includes(type) ? "Yes" : "No")].join(',');
        csvRows.push(workTypeData);

        // New representation for PPEs to be Used - NO QUOTES for headers
        const allPPEs = Array.from(ppeCheckboxes).map(checkbox => checkbox.value);
        const ppeHeader = ["PPEs to be Used:"].concat(allPPEs).join(',');
        csvRows.push(ppeHeader);
        const ppeData = ["", ...allPPEs.map(ppe => msraData.jobInfo.ppesUsed.includes(ppe) ? "Yes" : "No")].join(',');
        csvRows.push(ppeData);

        csvRows.push(''); // Empty line for separation

        // CSV Header for Tasks and Risks - NO QUOTES for headers
        const header = [
            'Task ID', 'Task Description', 'Risk ID', 'Risk Description',
            'Initial Likelihood', 'Initial Severity', 'Initial Risk Level',
            'Countermeasure',
            'New Likelihood', 'New Severity', 'Residual Risk Level'
        ];
        csvRows.push(header.join(',')); // No quotes for header

        // Add Task and Risk Data - retain quoting for descriptions/countermeasures if they contain commas
        msraData.tasks.forEach(task => {
            const escapedTaskDescription = `"${task.description.replace(/"/g, '""')}"`;
            if (task.risks.length === 0) {
                const row = [
                    task.id,
                    escapedTaskDescription,
                    '', '', '', '', '', '', '', '', ''
                ];
                csvRows.push(row.join(','));
            } else {
                task.risks.forEach(risk => {
                    const escapedRiskDescription = `"${risk.description.replace(/"/g, '""')}"`;
                    const escapedCountermeasure = `"${(risk.countermeasure || '').replace(/"/g, '""')}"`;
                    const row = [
                        task.id,
                        escapedTaskDescription,
                        risk.id,
                        escapedRiskDescription,
                        risk.likelihoodLabel,
                        risk.impactLabel,
                        risk.initialLevel,
                        escapedCountermeasure,
                        risk.newLikelihoodLabel,
                        risk.newImpactLabel,
                        risk.residualLevel
                    ];
                    csvRows.push(row.join(','));
                });
            }
        });

        const csvString = csvRows.join('\n');
        const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.setAttribute('download', 'method_statement_risk_assessment.csv');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        dataModified = false; // Data saved, reset flag
        updateUnsavedChangesWarning(); // Update warning message
    };

    // New: Function to download CSV template
    const downloadCsvTemplate = () => {
        let csvRows = [];

        // Add MSRA Header Information (empty values) - NO QUOTES
        csvRows.push('Method Statement and Risk Assessment');
        csvRows.push(`Job Name:,`);
        csvRows.push(`Prepared By:,`);
        csvRows.push(`Contractor Name:,`);
        csvRows.push(`Contractor Person Name:,`);
        csvRows.push(`Date of Preparation:,`);

        // New representation for Type of Work in template - NO QUOTES for headers
        const allWorkTypes = Array.from(workTypeCheckboxes).map(checkbox => checkbox.value);
        const workTypeHeader = ["Type of Work:"].concat(allWorkTypes).join(',');
        csvRows.push(workTypeHeader);
        const workTypeTemplateData = ["", ...allWorkTypes.map(type => "No")].join(','); // Default to No
        csvRows.push(workTypeTemplateData);

        // New representation for PPEs to be Used in template - NO QUOTES for headers
        const allPPEs = Array.from(ppeCheckboxes).map(checkbox => checkbox.value);
        const ppeHeader = ["PPEs to be Used:"].concat(allPPEs).join(',');
        csvRows.push(ppeHeader);
        const ppeTemplateData = ["", ...allPPEs.map(ppe => "No")].join(','); // Default to No
        csvRows.push(ppeTemplateData);

        csvRows.push(''); // Empty line for separation

        // CSV Header for Tasks and Risks - NO QUOTES for headers
        const header = [
            'Task ID', 'Task Description', 'Risk ID', 'Risk Description',
            'Initial Likelihood', 'Initial Severity', 'Initial Risk Level',
            'Countermeasure',
            'New Likelihood', 'New Severity', 'Residual Risk Level'
        ];
        csvRows.push(header.join(',')); // No quotes for header
        csvRows.push(''); // Add an empty data row for clarity in template

        const csvString = csvRows.join('\n');
        const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.setAttribute('download', 'msra_template.csv');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };


    // Event listener for "Export MSRA to Image" button
    exportImageBtn.addEventListener('click', () => {
        if (msraData.tasks.length === 0 && !Object.values(msraData.jobInfo).some(value => value !== '' && value !== null && (Array.isArray(value) ? value.length > 0 : true))) {
            // Replaced alert with custom message box or modal if available
            const message = 'No MSRA data to export. Please fill in MSRA information or add tasks/risks.';
            alert(message);
            return;
        }

        // Store original display styles of all buttons to be hidden
        const elementsToHide = [
            ...document.querySelectorAll('.task-actions, .delete-btn, .countermeasure-btn, .edit-risk-btn'),
            exportCsvBtn, exportImageBtn, exportPdfBtn, importCsvBtn, exportCsvBtnTop, downloadTemplateBtn,
            unsavedChangesWarning
        ];
        const originalDisplays = elementsToHide.map(el => ({ element: el, display: el.style.display }));

        // Temporarily hide all specified buttons
        elementsToHide.forEach(el => el.style.display = 'none');

        // Temporarily remove overflow-x from table-container
        const originalOverflowX = tableContainer.style.overflowX;
        tableContainer.style.overflowX = 'visible';

        html2canvas(msraOverviewSection, {
            scale: 2,
            logging: false,
            useCORS: true
        }).then(canvas => {
            const imageDataURL = canvas.toDataURL('image/png');
            const link = document.createElement('a');
            link.href = imageDataURL;
            link.download = 'msra_overview.png';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }).catch(error => {
            console.error('Error exporting to image:', error);
            // Replaced alert with custom message box or modal if available
            const message = 'Failed to export to image. Please try again.';
            alert(message);
        }).finally(() => {
            // Restore original button display
            originalDisplays.forEach(item => {
                item.element.style.display = item.display;
            });
            // Restore overflow-x
            tableContainer.style.overflowX = originalOverflowX;
        });
    });

    // Event listener for "Export MSRA to PDF" button
    exportPdfBtn.addEventListener('click', () => {
        if (msraData.tasks.length === 0 && !Object.values(msraData.jobInfo).some(value => value !== '' && value !== null && (Array.isArray(value) ? value.length > 0 : true))) {
            // Replaced alert with custom message box or modal if available
            const message = 'No MSRA data to export. Please fill in MSRA information or add tasks/risks.';
            alert(message);
            return;
        }

        const originalButtonText = exportPdfBtn.textContent;
        exportPdfBtn.textContent = 'Generating PDF...';
        exportPdfBtn.disabled = true;

        // Create the signature block element
        const signatureBlock = document.createElement('div');
        signatureBlock.className = 'signature-block';
        signatureBlock.innerHTML = `
            <div class="signature-item">
                <div class="signature-line"></div>
                <div class="signature-label">Prepared by:</div>
            </div>
            <div class="signature-item">
                <div class="signature-line"></div>
                <div class="signature-label">Checked by:</div>
            </div>
            <div class="signature-item">
                <div class="signature-line"></div>
                <div class="signature-label">Approved by:</div>
            </div>
        `;

        // Store original display styles of all buttons to be hidden
        const elementsToHide = [
            ...document.querySelectorAll('.task-actions, .delete-btn, .countermeasure-btn, .edit-risk-btn'),
            exportCsvBtn, exportImageBtn, exportPdfBtn, importCsvBtn, exportCsvBtnTop, downloadTemplateBtn,
            unsavedChangesWarning
        ];
        const originalDisplays = elementsToHide.map(el => ({ element: el, display: el.style.display }));

        // Temporarily hide all specified buttons
        elementsToHide.forEach(el => el.style.display = 'none');

        // Temporarily remove overflow-x from table-container
        const originalOverflowX = tableContainer.style.overflowX;
        tableContainer.style.overflowX = 'visible';

        // Append the signature block to the section to be captured by html2canvas
        msraOverviewSection.appendChild(signatureBlock);

        html2canvas(msraOverviewSection, {
            scale: 2,
            logging: false,
            useCORS: true
        }).then(canvas => {
            const imgData = canvas.toDataURL('image/png');
            const imgWidth = 210;
            const pageHeight = 297;
            const imgHeight = canvas.height * imgWidth / canvas.width;
            let heightLeft = imgHeight;

            const doc = new jsPDF('p', 'mm', 'a4');
            let position = 0;

            doc.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;

            while (heightLeft >= 0) {
                position = heightLeft - imgHeight;
                doc.addPage();
                doc.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
                heightLeft -= pageHeight;
            }

            doc.save('msra_overview.pdf');

        }).catch(error => {
            console.error('Error exporting to PDF:', error);
            // Replaced alert with custom message box or modal if available
            const message = 'Failed to export to PDF. Please try again.';
            alert(message);
        }).finally(() => {
            exportPdfBtn.textContent = originalButtonText;
            exportPdfBtn.disabled = false;
            // Restore original button display
            originalDisplays.forEach(item => {
                item.element.style.display = item.display;
            });
            // Restore overflow-x
            tableContainer.style.overflowX = originalOverflowX;
            // Remove the signature block
            msraOverviewSection.removeChild(signatureBlock);
        });
    });

    // Event listener for "Import MSRA from CSV" button
    importCsvBtn.addEventListener('click', () => {
        csvFileInput.click();
    });

    // Event listener for when a file is selected in the input
    csvFileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) {
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const csvContent = e.target.result;
                parseAndImportCsv(csvContent);
                // Replaced alert with custom message box or modal if available
                const message = 'MSRA data imported successfully!';
                alert(message);
                dataModified = false; // Data loaded, reset flag
                updateUnsavedChangesWarning(); // Update warning message
            } catch (error) {
                console.error('Error importing CSV:', error);
                // Replaced alert with custom message box or modal if available
                const message = 'Failed to import CSV. Please ensure it is a valid MSRA export file.';
                alert(message);
            } finally {
                csvFileInput.value = '';
            }
        };
        reader.readAsText(file);
    });

    // Function to parse CSV content and update MSRA data
    const parseAndImportCsv = (csvContent) => {
        console.log("Starting CSV import parsing...");
        const lines = csvContent.split(/\r?\n/).map(line => line.trim()).filter(line => line.length > 0);
        console.log("Filtered lines:", lines);

        // Initialize msraData
        msraData = {
            jobInfo: {
                jobName: '',
                preparedBy: '',
                contractorName: '',
                contractorPersonName: '',
                datePrepared: '',
                typeOfWork: [],
                ppesUsed: []
            },
            tasks: []
        };
        taskIdCounter = 1;
        riskIdCounter = 1;

        let currentLineIndex = 0;

        // Helper to clean and unquote a CSV part
        const cleanCsvPart = (part) => {
            let cleaned = part.trim();
            // Remove outer quotes if present
            if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
                cleaned = cleaned.substring(1, cleaned.length - 1);
            }
            return cleaned.replace(/""/g, '"'); // Unescape double quotes
        };

        // Parse MSRA Header Information
        if (lines[currentLineIndex] && lines[currentLineIndex].includes('Method Statement and Risk Assessment')) {
            currentLineIndex++; // Move past the title line
        }

        // Parse fixed job info fields
        const fixedInfoMap = {
            "Job Name": "jobName",
            "Prepared By": "preparedBy",
            "Contractor Name": "contractorName",
            "Contractor Person Name": "contractorPersonName",
            "Date of Preparation": "datePrepared"
        };

        for (const keyBase in fixedInfoMap) {
            if (lines[currentLineIndex]) {
                const line = lines[currentLineIndex];
                // Use simple split for these lines, as they are not expected to have commas within quoted values
                const parts = line.split(',');
                
                if (parts.length > 1) {
                    // Clean the key part and remove trailing colon if present
                    const parsedKey = cleanCsvPart(parts[0]).replace(/:$/, '');
                    if (parsedKey === keyBase) {
                        // Join the rest of the parts (in case value has commas) and clean the value
                        msraData.jobInfo[fixedInfoMap[keyBase]] = cleanCsvPart(parts.slice(1).join(','));
                    }
                }
                currentLineIndex++;
            }
        }

        // Parse Type of Work
        if (lines[currentLineIndex]) {
            const workTypeHeaderLine = lines[currentLineIndex];
            // Check for "Type of Work:" (quoted or unquoted)
            if (cleanCsvPart(workTypeHeaderLine.split(',')[0]) === 'Type of Work:') { // Simple split here
                currentLineIndex++;
                if (lines[currentLineIndex]) {
                    const workTypeDataLine = lines[currentLineIndex];
                    const workTypeOptions = workTypeHeaderLine.split(',').slice(1).map(cleanCsvPart); // Simple split here
                    const workTypeSelections = workTypeDataLine.split(',').slice(1).map(cleanCsvPart); // Simple split here

                    msraData.jobInfo.typeOfWork = [];
                    workTypeOptions.forEach((option, index) => {
                        if (workTypeSelections[index] && workTypeSelections[index].toUpperCase() === 'YES') {
                            msraData.jobInfo.typeOfWork.push(option);
                        }
                    });
                    currentLineIndex++;
                }
            }
        }

        // Parse PPEs to be Used
        if (lines[currentLineIndex]) {
            const ppeHeaderLine = lines[currentLineIndex];
            // Check for "PPEs to be Used:" (quoted or unquoted)
            if (cleanCsvPart(ppeHeaderLine.split(',')[0]) === 'PPEs to be Used:') { // Simple split here
                currentLineIndex++;
                if (lines[currentLineIndex]) {
                    const ppeDataLine = lines[currentLineIndex];
                    const ppeOptions = ppeHeaderLine.split(',').slice(1).map(cleanCsvPart); // Simple split here
                    const ppeSelections = ppeDataLine.split(',').slice(1).map(cleanCsvPart); // Simple split here

                    msraData.jobInfo.ppesUsed = [];
                    ppeOptions.forEach((option, index) => {
                        if (ppeSelections[index] && ppeSelections[index].toUpperCase() === 'YES') {
                            msraData.jobInfo.ppesUsed.push(option);
                        }
                    });
                    currentLineIndex++;
                }
            }
        }

        // Skip any empty line before the main table header
        while (lines[currentLineIndex] === '') {
            currentLineIndex++;
        }

        // Find the main table header row dynamically
        let tableHeaderFound = false;
        const expectedTableHeaderKeywords = ['Task ID', 'Task Description', 'Risk ID', 'Countermeasure'];
        
        if (lines[currentLineIndex]) {
            const rawLine = lines[currentLineIndex];
            // Use simple split for header lines, then clean each part.
            const parts = rawLine.split(',');
            const cleanedParts = parts.map(cleanCsvPart);

            // Check if the cleaned parts contain all expected keywords
            const containsAllKeywords = expectedTableHeaderKeywords.every(keyword => 
                cleanedParts.some(part => part.toLowerCase().includes(keyword.toLowerCase())) // Use .includes for more flexibility
            );

            if (containsAllKeywords) {
                tableHeaderFound = true;
                currentLineIndex++; // Move past the table header
            }
        }

        if (!tableHeaderFound) {
            throw new Error(`CSV table header row not found. Expected keywords: "${expectedTableHeaderKeywords.join(', ')}"`);
        }
        console.log("Main table header found at filtered index:", currentLineIndex - 1);

        // Parse Task and Risk Data
        let maxTaskId = 0;
        let maxRiskId = 0;

        for (let i = currentLineIndex; i < lines.length; i++) {
            const line = lines[i];
            console.log(`Processing data line ${i}: ${line}`);
            // Use the regex split for data rows, as they can have internal commas in descriptions that are quoted
            const parts = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);
            const cleanedParts = parts.map(cleanCsvPart);
            console.log("Cleaned parts for current line:", cleanedParts);

            // Safely access parts, providing default empty strings if index is out of bounds
            const taskId = parseInt(cleanedParts[0] || '');
            const taskDescription = cleanedParts[1] || '';
            const riskId = parseInt(cleanedParts[2] || '');
            const riskDescription = cleanedParts[3] || '';
            const likelihoodLabel = cleanedParts[4] || '';
            const impactLabel = cleanedParts[5] || '';
            const initialLevel = cleanedParts[6] || '';
            const countermeasure = cleanedParts[7] || '';
            const newLikelihoodLabel = cleanedParts[8] || '';
            const newImpactLabel = cleanedParts[9] || '';
            const residualLevel = cleanedParts[10] || '';

            if (!isNaN(taskId) && taskId > maxTaskId) maxTaskId = taskId;
            if (!isNaN(riskId) && riskId > maxRiskId) maxRiskId = riskId;

            let task = msraData.tasks.find(t => t.id === taskId);
            if (!task) {
                if (!isNaN(taskId) && taskDescription) {
                    task = {
                        id: taskId,
                        description: taskDescription,
                        risks: []
                    };
                    msraData.tasks.push(task);
                    console.log(`Created new task object: Task ${task.id}: ${task.description}`);
                } else {
                    console.log(`Skipping task creation for line ${i} due to invalid Task ID or Description.`);
                    continue;
                }
            }

            if (!isNaN(riskId) && riskDescription) {
                const risk = {
                    id: riskId,
                    description: riskDescription,
                    likelihood: (likelihoodLabel && !isNaN(parseInt(getLikelihoodValueFromLabel(likelihoodLabel)))) ? parseInt(getLikelihoodValueFromLabel(likelihoodLabel)) : null,
                    likelihoodLabel: likelihoodLabel,
                    impact: (impactLabel && !isNaN(parseInt(getImpactValueFromLabel(impactLabel)))) ? parseInt(getImpactValueFromLabel(impactLabel)) : null,
                    impactLabel: impactLabel,
                    initialLevel: initialLevel,
                    countermeasure: countermeasure,
                    newLikelihood: (newLikelihoodLabel && !isNaN(parseInt(getLikelihoodValueFromLabel(newLikelihoodLabel)))) ? parseInt(getLikelihoodValueFromLabel(newLikelihoodLabel)) : null,
                    newLikelihoodLabel: newLikelihoodLabel,
                    newImpact: (newImpactLabel && !isNaN(parseInt(getImpactValueFromLabel(newImpactLabel)))) ? parseInt(getImpactValueFromLabel(newImpactLabel)) : null,
                    newImpactLabel: newImpactLabel,
                    residualLevel: residualLevel
                };
                task.risks.push(risk);
                console.log(`Added risk ${risk.id} to task ${task.id}`);
            } else {
                console.log(`Skipping risk addition for line ${i} due to invalid Risk ID or Risk Description.`);
            }
        }

        msraData.tasks.sort((a, b) => a.id - b.id);
        msraData.tasks.forEach((task, index) => {
            task.id = index + 1;
            task.risks.sort((a, b) => a.id - b.id);
        });

        taskIdCounter = maxTaskId + 1;
        riskIdCounter = maxRiskId + 1;
        console.log("Final msraData after import:", msraData);
        console.log("Next taskIdCounter:", taskIdCounter, "Next riskIdCounter:", riskIdCounter);

        renderMSRAInfoSummary();
        renderMSRATable();
    };

    // Helper functions to get numerical values from labels for import
    const getLikelihoodValueFromLabel = (label) => {
        // Predefined map for efficiency and robustness
        const likelihoodMap = {
            'A': 5, 'B': 4, 'C': 3, 'D': 2, '1': 1, 'E': 1 
        };
        return likelihoodMap[label] || ''; // Return empty string if not found
    };

    const getImpactValueFromLabel = (label) => {
        // Predefined map for efficiency and robustness
        const impactMap = {
            '1': 1, '2': 2, '3': 3, '4': 4, '5': 5
        };
        return impactMap[label] || ''; // Return empty string if not found
    };

    // Function to update the visibility of the unsaved changes warning
    const updateUnsavedChangesWarning = () => {
        const hasJobInfo = Object.values(msraData.jobInfo).some(value => value !== '' && value !== null && (Array.isArray(value) ? value.length > 0 : true));
        const hasTasks = msraData.tasks.length > 0;

        if (dataModified || hasJobInfo || hasTasks) {
            document.getElementById('unsaved-changes-warning').style.display = 'block';
        } else {
            document.getElementById('unsaved-changes-warning').style.display = 'none';
        }
    };

    // Add beforeunload event listener for warning
    window.addEventListener('beforeunload', (event) => {
        // Check if there's any data in jobInfo or tasks
        const hasJobInfo = Object.values(msraData.jobInfo).some(value => value !== '' && value !== null && (Array.isArray(value) ? value.length > 0 : true));
        const hasTasks = msraData.tasks.length > 0;

        if (dataModified || hasJobInfo || hasTasks) {
            event.preventDefault(); // Standard for browser to show confirmation
            event.returnValue = ''; // For older browsers
        }
    });

    // --- CSV Converter Functions (Integrated) ---

    /**
     * Parses a single CSV line, correctly handling quoted fields and escaped quotes.
     * @param {string} line - The CSV line to parse.
     * @returns {string[]} An array of parsed fields.
     */
    function parseCsvLine(line) {
        const fields = [];
        let inQuote = false;
        let currentField = '';

        for (let i = 0; i < line.length; i++) {
            const char = line[i];

            if (char === '"') {
                // Check for escaped quote ""
                if (inQuote && line[i + 1] === '"') {
                    currentField += '"'; // Add a single quote for ""
                    i++; // Skip the next quote
                } else {
                    inQuote = !inQuote; // Toggle inQuote state
                }
            } else if (char === ',' && !inQuote) {
                fields.push(currentField);
                currentField = '';
            } else {
                currentField += char;
            }
        }
        fields.push(currentField); // Add the last field
        return fields;
    }

    /**
     * Converts the CSV content from Excel format back to the original format.
     * This function handles date format, adding quotes, removing trailing commas,
     * and converting lines that are only commas into blank lines.
     * @param {string} csvContent - The raw CSV content from Excel.
     * @returns {string} The converted CSV content.
     */
    function convertCsvFormat(csvContent) {
        const lines = csvContent.split(/\r?\n/);
        const outputLines = [];
        let inDataSection = false;

        // Helper to clean and unquote a CSV part (re-defined for this scope to match behavior)
        const cleanCsvPartLocal = (part) => {
            let cleaned = part.trim();
            if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
                cleaned = cleaned.substring(1, cleaned.length - 1);
            }
            return cleaned.replace(/""/g, '"'); // Unescape double quotes
        };

        for (let i = 0; i < lines.length; i++) {
            let line = lines[i];

            line = line.trim();
            line = line.replace(/,+$/, ''); // Remove any trailing commas

            if (line === '' && !inDataSection) {
                outputLines.push('');
                continue;
            }

            // Date Format Conversion (DD-MM-YYYY to YYYY-MM-DD)
            if (line.startsWith('Date of Preparation:')) { // Check for unquoted prefix
                const parts = parseCsvLine(line); // Use robust parser
                if (parts.length === 2) {
                    const datePart = cleanCsvPartLocal(parts[1]);
                    const dateRegex = /^(\d{2})-(\d{2})-(\d{4})$/; // DD-MM-YYYY
                    const match = datePart.match(dateRegex);
                    if (match) {
                        const [_, day, month, year] = match;
                        outputLines.push(`Date of Preparation:,${year}-${month}-${day}`);
                        continue;
                    }
                }
            }

            // Detect start of data section more robustly
            const dataHeaderKeywords = ['Task ID', 'Task Description', 'Risk ID', 'Risk Description'];
            const currentLineParts = parseCsvLine(line).map(cleanCsvPartLocal);
            const isDataHeader = dataHeaderKeywords.every(keyword => 
                currentLineParts.some(part => part.toLowerCase().includes(keyword.toLowerCase()))
            );

            if (isDataHeader) {
                inDataSection = true;
                // Reconstruct header line with quotes for each part as per Original.csv style
                const quotedHeaderParts = currentLineParts.map(part => `"${part}"`);
                outputLines.push(quotedHeaderParts.join(','));
                continue;
            }

            // Add Double Quotes to specific fields in data rows
            if (inDataSection && line !== '') {
                const fields = parseCsvLine(line);
                const quotedFields = [];

                // Indices that need to be quoted based on Original.csv
                const indicesToQuote = [1, 3, 7]; // Task Description, Risk Description, Countermeasure

                for (let j = 0; j < fields.length; j++) {
                    let field = fields[j];
                    // Clean the field from any Excel-introduced outer quotes before processing
                    field = cleanCsvPartLocal(field);

                    if (indicesToQuote.includes(j)) {
                        if (field !== '') {
                            quotedFields.push(`"${field.replace(/"/g, '""')}"`); // Escape internal quotes and wrap
                        } else {
                            quotedFields.push(field); // Keep empty fields as empty
                        }
                    } else {
                        // For other fields, if they contain a comma or quote, they must be quoted.
                        // Otherwise, add as is.
                        if (field.includes(',') || field.includes('"')) {
                            const escapedField = field.replace(/"/g, '""');
                            quotedFields.push(`"${escapedField}"`);
                        } else {
                            quotedFields.push(field);
                        }
                    }
                }
                outputLines.push(quotedFields.join(','));
            } else {
                // For other lines (like "Job Name:", "Type of Work:", "PPEs to be Used:")
                // or lines before the data section header, add them with appropriate quoting.
                // This part tries to make them look like "Original.csv"
                const parts = parseCsvLine(line);
                if (parts.length > 1 && (parts[0].endsWith(':') || parts[0].startsWith('"') && parts[0].endsWith(':"'))) {
                    // This is likely a key-value pair line like "Job Name:,test" or ""Job Name:"","test""
                    const cleanedKey = cleanCsvPartLocal(parts[0]);
                    const cleanedValue = cleanCsvPartLocal(parts.slice(1).join(',')); // Join rest for value
                    outputLines.push(`"${cleanedKey}","${cleanedValue}"`);
                } else if (parts.length > 1 && (cleanCsvPartLocal(parts[0]) === 'Type of Work:' || cleanCsvPartLocal(parts[0]) === 'PPEs to be Used:')) {
                    // This is a checkbox header line
                    const headerParts = parts.map(cleanCsvPartLocal);
                    outputLines.push(headerParts.map(part => `"${part}"`).join(','));
                }
                else {
                    // For lines like ",Yes,No,Yes..." or other simple lines, keep as is
                    outputLines.push(line);
                }
            }
        }
        return outputLines.join('\n');
    }


    // Event listener for "Convert Excel CSV" button
    convertExcelCsvBtn.addEventListener('click', () => {
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = '.csv';
        fileInput.style.display = 'none';
        document.body.appendChild(fileInput);

        fileInput.addEventListener('change', (event) => {
            const file = event.target.files[0];
            if (!file) {
                alert('No file selected.');
                document.body.removeChild(fileInput);
                return;
            }

            if (file.type !== 'text/csv') {
                alert('Invalid file type. Please select a CSV file.');
                document.body.removeChild(fileInput);
                return;
            }

            alert('Converting CSV...');

            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const content = e.target.result;
                    const convertedContent = convertCsvFormat(content);
                    const blob = new Blob([convertedContent], { type: 'text/csv;charset=utf-8;' });
                    const url = URL.createObjectURL(blob);

                    const tempLink = document.createElement('a');
                    tempLink.href = url;

                    const originalFileName = file.name;
                    const parts = originalFileName.split('.');
                    const extension = parts.pop();
                    const baseName = parts.join('.');
                    tempLink.download = `${baseName}_converted.${extension}`;

                    document.body.appendChild(tempLink);
                    tempLink.click();
                    document.body.removeChild(tempLink);
                    URL.revokeObjectURL(url);

                    alert('Conversion successful! File downloaded.');
                } catch (error) {
                    console.error('Error during conversion:', error);
                    alert('An error occurred during conversion. Please check the file format.');
                } finally {
                    document.body.removeChild(fileInput);
                }
            };
            reader.onerror = () => {
                alert('Failed to read file.');
                document.body.removeChild(fileInput);
            };
            reader.readAsText(file);
        });

        fileInput.click();
    });


    // Initial render calls
    renderMSRAInfoSummary();
    renderMSRATable();
    updateUnsavedChangesWarning(); // Call on initial load to set visibility
});