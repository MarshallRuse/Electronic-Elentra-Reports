const { 
    dialog, 
    Menu, 
    MenuItem, 
    getCurrentWindow
} = require('electron').remote;
const Store = require('electron-store');
const { PythonShell } = require('python-shell');
const path = require('path');
const customTitleBar = require('custom-electron-titlebar');

// Pyshell is instantiated with these options each time a generate-button is clicked.
// The instance is terminated once a '100' prog message is received.
let pyOptions = {
    pythonPath: path.join('src','server','py','venv','Scripts','python.exe'),
    scriptPath: path.join('src','server','py'),
    pythonOptions: ['-u'],
    mode: 'json'
}

const store = new Store();

let state = {
    FormatExtract: {
        extractDataFilePath: "",
        Options: {}
    },
    GenerateReport: {
        Options: {}
    },
    ReportSettings: {
        pathSeparator: path.sep,
    },
};

state = Object.assign(state, store.store);


// Enable right click inspect element for easier debugging
let rightClickPosition = null;
const menu = new Menu()
const menuItem = new MenuItem({
  label: 'Inspect Element',
  click: () => {
    getCurrentWindow().inspectElement(rightClickPosition.x, rightClickPosition.y)
  }
})
menu.append(menuItem)

window.addEventListener('contextmenu', (e) => {
    e.preventDefault()
    rightClickPosition = {x: e.x, y: e.y}
    menu.popup(getCurrentWindow())
}, false)

// Customize the title bar
new customTitleBar.Titlebar({
    backgroundColor: customTitleBar.Color.fromHex('#6d7fcc'),
    menuPosition: 'bottom',
    titleHorizontalAlignment: 'left'
});

// Collapsable Sidebar Controls
const sidebar = document.getElementById("sidebar");
const closeIcon = document.createElement("img");
closeIcon.src = "src\\contents\\icons\\close.svg";
closeIcon.classList.add("icon-button");
const menuIcon = document.createElement("img");
menuIcon.src = "src\\contents\\icons\\menu.svg";
menuIcon.classList.add("icon-button");

const btnSidebarCollapse = document.getElementById("sidebarCollapse");
btnSidebarCollapse.addEventListener("click", () => {
    sidebar.classList.toggle("active");
    if (sidebar.classList.contains("active")){
        btnSidebarCollapse.innerHTML = "";
        btnSidebarCollapse.appendChild(menuIcon);
    } else {
        btnSidebarCollapse.innerHTML = "";
        btnSidebarCollapse.appendChild(closeIcon);
    }
});

// Menu Controls
const menuFormatExtract = document.getElementById("menuFormatExtract");
const menuReportSettings = document.getElementById("menuReportSettings");

menuFormatExtract.addEventListener("click", () => changeAppPage("formatExtract"));
menuGenerateReport.addEventListener("click", () => changeAppPage("generateReport"));
menuReportSettings.addEventListener("click", () => changeAppPage("reportSettings"));

// Progress Toast
const progressToast = document.getElementById("progressToast");
const progressToastHeader = document.getElementById("toastHeaderText");
const progressToastText = document.getElementById("toastBodyProgressText");
const progressBar = document.getElementById("reportProgressBar");
const btnProgressToastClose = document.getElementById("btnProgressToastClose");

// All Generate Buttons
const generateButtons = document.getElementsByClassName("generate-button");
// Format Extract Button
const btnFormatExtract = document.getElementById("btnFormatExtract");
// Generate Report Button
const btnGenerateReport = document.getElementById("btnGenerateReport");


/***************************
 * 
 * Format Extract page
 * 
 **************************/
// File Chooser Controls
const FE_inpExtractData = document.getElementById("FE_inpExtractDataFile");
const FE_lblExtractData = document.getElementById("FE_lblExtractDataFile");
FE_inpExtractData.addEventListener("click", async () => {
    FE_lblExtractData.innerText = "";
    let file = await dialog.showOpenDialog({
        title: 'Choose Elentra Extract',
        filters: [{
            name: 'csv, xl*',
            extensions: ['csv', 'xl*']
        }]
    });
    if (file.canceled){
        FE_lblExtractData.innerText = "Choose File";
        state.FormatExtract.extractDataFilePath = "";
        if (!btnFormatExtract.hasAttribute("disabled")){
            btnFormatExtract.disabled = true;
        } 

    } else {
        FE_lblExtractData.innerHTML = `<span class="overflow-text">${file.filePaths[0]}</span>`;
        state.FormatExtract.extractDataFilePath = file.filePaths[0];
        if (btnFormatExtract.hasAttribute("disabled") && state.FormatExtract.lookupTableFilePath !== ""){
            btnFormatExtract.removeAttribute("disabled");
        }
    }
});

const FE_inpLookupTable = document.getElementById("FE_inpLookupTableFile");
state["FormatExtract"]["lookupTableFilePath"] = store.get('FormatExtract.lookupTableFilePath', "");
const FE_lblLookupTable = document.getElementById("FE_lblLookupTableFile");
if (state.FormatExtract.lookupTableFilePath !== ""){
    FE_lblLookupTable.innerText = state.FormatExtract.lookupTableFilePath;
}
FE_inpLookupTable.addEventListener("click", async () => {
    FE_lblLookupTable.innerText = "";
    let file = await dialog.showOpenDialog({
        title: 'Choose a Lookup Table',
        filters: [{
            name: 'csv, xl*',
            extensions: ['csv', 'xl*']
        }]
    });
    if (file.canceled){
        FE_lblLookupTable.innerText = "Choose File";
        state.FormatExtract.lookupTableFilePath = "";
        if (!btnFormatExtract.hasAttribute("disabled")){
            btnFormatExtract.disabled = true;
        } 
    } else {
        FE_lblLookupTable.innerHTML = `<span class="overflow-text">${file.filePaths[0]}</span>`;
        state.FormatExtract.lookupTableFilePath = file.filePaths[0];
        if (btnFormatExtract.hasAttribute("disabled") && state.FormatExtract.extractDataFilePath !== ""){
            btnFormatExtract.removeAttribute("disabled");
        }
    }
});

// Format Extract Button & Save Changed Options Modal controls
btnFormatExtract.addEventListener("click", () => {

    // Use the optional argument in place for GenerateReport
    state.FormatExtract.Options.createSpinoffExtract = true;

    if (!objectsEqual(objectMask(state.FormatExtract, ["extractDataFilePath"]), store.store.FormatExtract)){
        console.log("Store on generate: ", store.store.FormatExtract);
        console.log("State on generate: ", objectMask(state.FormatExtract, ["extractDataFilePath"]));
        $('#saveChangesModal').modal({backdrop: 'static'});
        const btnSaveOptionsChanges = document.getElementById("btnSaveOptionsChanges");
        
        btnSaveOptionsChanges.addEventListener("click", () => {
            store.set({
                FormatExtract: objectMask(state.FormatExtract, ["extractDataFilePath"])
            });
            $('#saveChangesModal').modal('hide');
            createSuccessAlert("Options saved for next time!");            
        });
    }

    progressToastHeader.innerText = "Format Extract Progress";
    progressBarReset();
    $("#progressToast").toast("show");
    for (let i = 0; i < generateButtons.length; i++){
        generateButtons[i].setAttribute("disabled", true);
    }

    const pyshell = new PythonShell('CreateReports.py', pyOptions);
    pythonCommand(pyshell, 'createFormattedExtract', { ...state.FormatExtract, ...state.ReportSettings});
});

// Option Toggle Values
const FE_tglRemoveUnsubmitted = document.getElementById('FE_tglRemoveUnsubmitted');
state.FormatExtract.Options.removeUnsubmitted = store.get('FormatExtract.Options.removeUnsubmitted', true);
FE_tglRemoveUnsubmitted.checked = state.FormatExtract.Options.removeUnsubmitted;
FE_tglRemoveUnsubmitted.addEventListener('change', (e) => {
    state.FormatExtract.Options.removeUnsubmitted = e.target.checked;
});

const FE_tglRemoveProcedures = document.getElementById('FE_tglRemoveProcedures');
state.FormatExtract.Options.removeProcedures = store.get('FormatExtract.Options.removeProcedures', true);
FE_tglRemoveProcedures.checked = state.FormatExtract.Options.removeProcedures;
FE_tglRemoveProcedures.addEventListener('change', (e) => {
    state.FormatExtract.Options.removeProcedures = e.target.checked;
});

const FE_tglRemoveEmptyColumns = document.getElementById('FE_tglRemoveEmptyColumns');
state.FormatExtract.Options.removeEmptyColumns = store.get('FormatExtract.Options.removeEmptyColumns', true);
FE_tglRemoveEmptyColumns.checked = state.FormatExtract.Options.removeEmptyColumns;
FE_tglRemoveEmptyColumns.addEventListener('change', (e) => {
    state.FormatExtract.Options.removeEmptyColumns = e.target.checked;
});

const FE_tglFormatAsTable = document.getElementById('FE_tglFormatAsTable');
state.FormatExtract.Options.spinoffExtractAsTable = store.get('FormatExtract.Options.spinoffExtractAsTable', true);
FE_tglFormatAsTable.checked = state.FormatExtract.Options.spinoffExtractAsTable;
FE_tglFormatAsTable.addEventListener('change', (e) => {
    state.FormatExtract.Options.spinoffExtractAsTable = e.target.checked;
});



/***************************
 * 
 * Generate Report page
 * 
 **************************/

//  // Data Source Radio Button 
// const GR_rbUseRawExtract = document.getElementById("GR_rbUseRawExtract"); 
// const GR_rbUseFormattedExtract = document.getElementById("GR_rbUseFormattedExtract"); 

// File Chooser controls
const GR_inpExtractData = document.getElementById("GR_inpExtractDataFile");
const GR_lblExtractData = document.getElementById("GR_lblExtractDataFile");
GR_inpExtractData.addEventListener("click", async () => {
    GR_lblExtractData.innerText = "";
    let file = await dialog.showOpenDialog({
        title: 'Choose Elentra Extract',
        filters: [{
            name: 'csv, xl*',
            extensions: ['csv', 'xl*']
        }]
    });
    if (file.canceled){
        GR_lblExtractData.innerText = "Choose File";
        state.GenerateReport.extractDataFilePath = "";
        if (!btnGenerateReport.hasAttribute("disabled")){
            btnGenerateReport.disabled = true;
        } 

    } else {
        GR_lblExtractData.innerHTML = `<span class="overflow-text">${file.filePaths[0]}</span>`;
        state.GenerateReport.extractDataFilePath = file.filePaths[0];
        if (btnGenerateReport.hasAttribute("disabled") && state.GenerateReport.lookupTableFilePath !== ""){
            btnGenerateReport.removeAttribute("disabled");
        }
    }
});

const GR_inpLookupTable = document.getElementById("GR_inpLookupTableFile");
state["GenerateReport"]["lookupTableFilePath"] = store.get('GenerateReport.lookupTableFilePath', "");
const GR_lblLookupTable = document.getElementById("GR_lblLookupTableFile");
if (state.GenerateReport.lookupTableFilePath !== ""){
    GR_lblLookupTable.innerText = state.GenerateReport.lookupTableFilePath;
}
GR_inpLookupTable.addEventListener("click", async () => {
    GR_lblLookupTable.innerText = "";
    let file = await dialog.showOpenDialog({
        title: 'Choose a Lookup Table',
        filters: [{
            name: 'csv, xl*',
            extensions: ['csv', 'xl*']
        }]
    });
    if (file.canceled){
        GR_lblLookupTable.innerText = "Choose File";
        state.GenerateReport.lookupTableFilePath = "";
        if (!btnGenerateReport.hasAttribute("disabled")){
            btnGenerateReport.disabled = true;
        }
    } else {
        GR_lblLookupTable.innerHTML = `<span class="overflow-text">${file.filePaths[0]}</span>`;
        state.GenerateReport.lookupTableFilePath = file.filePaths[0];
        if (btnGenerateReport.hasAttribute("disabled") && state.GenerateReport.extractDataFilePath !== ""){
            btnGenerateReport.removeAttribute("disabled");
        }
    }
});

btnGenerateReport.addEventListener("click", () => {

    // Use the optional argument in place for GenerateReport
    //state.GenerateReport.Options.createSpinoffExtract = true;

    if (!objectsEqual(objectMask(state.GenerateReport, ["extractDataFilePath"]), store.store.GenerateReport)){
        console.log("Store on generate: ", store.store.GenerateReport);
        console.log("State on generate: ", objectMask(state.GenerateReport, ["extractDataFilePath"]));
        $('#saveChangesModal').modal({backdrop: 'static'});
        const btnSaveOptionsChanges = document.getElementById("btnSaveOptionsChanges");
        
        btnSaveOptionsChanges.addEventListener("click", () => {
            store.set({
                GenerateReport: objectMask(state.GenerateReport, ["extractDataFilePath"])
            });
            $('#saveChangesModal').modal('hide');
            createSuccessAlert("Generate Report options saved for next time!");            
        });
    }

    progressToastHeader.innerText = "Generate Report Progress";
    progressBarReset();
    $("#progressToast").toast("show");
    for (let i = 0; i < generateButtons.length; i++){
        generateButtons[i].setAttribute("disabled", true);
    }

    const pyshell = new PythonShell('CreateReports.py', pyOptions);
    pythonCommand(pyshell, 'createGeneratedReport', { ...state.GenerateReport, ...state.ReportSettings});
});

// Option Toggle Values
const GR_tglRemoveUnsubmitted = document.getElementById('GR_tglRemoveUnsubmitted');
state.GenerateReport.Options.removeUnsubmitted = store.get('GenerateReport.Options.removeUnsubmitted', true);
GR_tglRemoveUnsubmitted.checked = state.GenerateReport.Options.removeUnsubmitted;
GR_tglRemoveUnsubmitted.addEventListener('change', (e) => {
    state.GenerateReport.Options.removeUnsubmitted = e.target.checked;
});

const GR_tglRemoveProcedures = document.getElementById('GR_tglRemoveProcedures');
state.GenerateReport.Options.removeProcedures = store.get('GenerateReport.Options.removeProcedures', true);
GR_tglRemoveProcedures.checked = state.GenerateReport.Options.removeProcedures;
GR_tglRemoveProcedures.addEventListener('change', (e) => {
    state.GenerateReport.Options.removeProcedures = e.target.checked;
});

const GR_tglRemoveEmptyColumns = document.getElementById('GR_tglRemoveEmptyColumns');
state.GenerateReport.Options.removeEmptyColumns = store.get('GenerateReport.Options.removeEmptyColumns', true);
GR_tglRemoveEmptyColumns.checked = state.GenerateReport.Options.removeEmptyColumns;
GR_tglRemoveEmptyColumns.addEventListener('change', (e) => {
    state.GenerateReport.Options.removeEmptyColumns = e.target.checked;
});

const GR_tglCreateSpinoffExtract = document.getElementById('GR_tglCreateSpinoffExtract');
state.GenerateReport.Options.createSpinoffExtract = store.get('GenerateReport.Options.createSpinoffExtract', true);
GR_tglCreateSpinoffExtract.checked = state.GenerateReport.Options.createSpinoffExtract;
GR_tglCreateSpinoffExtract.addEventListener('change', (e) => {
    state.GenerateReport.Options.createSpinoffExtract = e.target.checked;
    if (!state.GenerateReport.Options.createSpinoffExtract && state.GenerateReport.Options.spinoffExtractAsTable){
        state.GenerateReport.Options.spinoffExtractAsTable = false;
        GR_tglSpinoffExtractAsTable.checked = false;
    }
});

const GR_tglSpinoffExtractAsTable = document.getElementById('GR_tglSpinoffExtractAsTable');
state.GenerateReport.Options.spinoffExtractAsTable = store.get('GenerateReport.Options.spinoffExtractAsTable', true);
GR_tglSpinoffExtractAsTable.checked = state.GenerateReport.Options.spinoffExtractAsTable;
GR_tglSpinoffExtractAsTable.addEventListener('change', (e) => {
    state.GenerateReport.Options.spinoffExtractAsTable = e.target.checked;
    if (state.GenerateReport.Options.spinoffExtractAsTable && !state.GenerateReport.Options.createSpinoffExtract){
        state.GenerateReport.Options.createSpinoffExtract = true;
        GR_tglCreateSpinoffExtract.checked = true;
    }
});

const GR_tglIncludeExtractData = document.getElementById('GR_tglIncludeExtractData');
state.GenerateReport.Options.includeExtractDataInReport = store.get('GenerateReport.Options.includeExtractDataInReport', true);
GR_tglIncludeExtractData.checked = state.GenerateReport.Options.includeExtractDataInReport;
GR_tglIncludeExtractData.addEventListener('change', (e) => {
    state.GenerateReport.Options.includeExtractDataInReport = e.target.checked;
});


/***************************
 * 
 * Report Settings page
 * 
 **************************/
// Report Settings Save Button
const btnSaveReportSettings = document.getElementById("btnSaveReportSettings");

// File Chooser Controls
const inpSaveReportFolder = document.getElementById("inpSaveReportFolder");
state.ReportSettings.saveReportFolderPath = store.get('ReportSettings.saveReportFolderPath', path.dirname(__dirname).split(path.sep).slice(0,-1).join(path.sep));
const lblSaveReportFolder = document.getElementById("lblSaveReportFolder");
if (state.ReportSettings.saveReportFolderPath !== ""){
    lblSaveReportFolder.innerText = state.ReportSettings.saveReportFolderPath;
}
inpSaveReportFolder.addEventListener("click", async () => {
    lblSaveReportFolder.innerText = "";
    let file = await dialog.showOpenDialog({
        title: 'Choose a Save Location for Report',
        filters: [{
            name: 'csv, xl*',
            extensions: ['csv', 'xl*']
        }],
        properties: ['openDirectory', 'createDirectory']
    });
    if (file.canceled){
        lblSaveReportFolder.innerText = "Choose Folder";
        state.ReportSettings.saveReportFolderPath = store.get('saveReportFolderPath', path.dirname(__dirname));;
    } else {
        lblSaveReportFolder.innerHTML = `<span class="overflow-text">${file.filePaths[0]}</span>`;
        state.ReportSettings.saveReportFolderPath = file.filePaths[0];
        btnSaveReportSettings.disabled = false;
    }
});

// Save button actions
btnSaveReportSettings.addEventListener("click", () => {
    store.set({
        ReportSettings: state.ReportSettings
    }) 
    createSuccessAlert("Report Settings saved for next time!");
    btnSaveReportSettings.disabled = true;         
});


/***************************
 * 
 * Utility Functions
 * 
 **************************/

function objectsEqual(obj1, obj2){
    
    const keys1 = Object.keys(obj1);
    const keys2 = Object.keys(obj2);

    if (keys1.length !== keys2.length) {
        return false;
    }

    for (let key of keys1) {
        const val1 = obj1[key];
        const val2 = obj2[key];
        const areObjects = (val1 !== null && typeof val1 === 'object') && (val2 !== null && typeof val2 === 'object')
        if (
            areObjects && !objectsEqual(val1, val2) || 
            !areObjects && val1 !== val2
        ){
            console.log("NOT EQUAL: ", val1, val2);
            return false
        }
    }

    return true;
}

function objectMask(obj, keysToHide){

    let objCopy = {};
    objCopy = Object.assign(objCopy, obj);
    for (let key of keysToHide){
        delete objCopy[key];
    }
    return objCopy;
}

function createSuccessAlert(alertText){
    let wrapperDiv = document.createElement('div');
    wrapperDiv.classList.add('alert-wrapper', 'slide-in-right');
    let alertDiv = document.createElement('div');
    alertDiv.classList.add('alert', 'alert-success');
    alertDiv.setAttribute('role', 'alert');
    alertDiv.innerText = alertText;

    wrapperDiv.appendChild(alertDiv);

    document.body.appendChild(wrapperDiv);

    setTimeout(() => {
        wrapperDiv.classList.remove('slide-in-right');
        wrapperDiv.classList.add('slide-out-left');
        setTimeout(() => {
            wrapperDiv.remove();
        }, 1000)
    }, 3000)
}

function progressBarFinished(){
    
    progressBar.classList.remove("progress-bar-animated", "progress-bar-striped");
    progressBar.classList.add("bg-success");
    btnProgressToastClose.style.display = "inline";
    progressToast.classList.add("progress-toast-success");
}

function progressBarReset(){
    progressToast.classList.remove("progress-toast-success");
    progressBar.setAttribute("aria-valuenow", "0");
    progressBar.style.width = `0%`;
    progressBar.classList.remove("bg-success");
    progressBar.classList.add("progress-bar-animated", "progress-bar-striped");
    btnProgressToastClose.style.display = "none";
}

function unsupportedValsModalReset(){
    let modalBod = document.getElementById("fixUnsupportedEntrustmentsModalBody");
    modalBod.innerHTML = "";
}

function changeAppPage(pageID){
    let shownPage = document.getElementById(pageID);
    let appPages = document.getElementsByClassName('appPage');
    for (let i = 0; i < appPages.length; i++){
        appPages[i].style.display = "none";
    }
    shownPage.style.display = "block";

    // change active menu item
    let sidebarMenuItems = document.getElementsByClassName("sidebarMenuItem");
    for (let i = 0; i < sidebarMenuItems.length; i++){
        sidebarMenuItems[i].classList.remove("active");
    }

    let menuStr = "menu";
    let menuItemID = menuStr.concat(pageID.charAt(0).toUpperCase() + pageID.slice(1));
    // console.log("menuItemID: ", menuItemID);
    let activeMenuItemATag = document.getElementById(menuItemID);
    let activeMenuItemLITag = activeMenuItemATag.parentNode;
    activeMenuItemLITag.classList.add("active");
}



function pythonCommand(pyshell, pythonFunction, appState){
    console.log("APP STATE: ", appState);

    let funcAndState = {
        func: pythonFunction
    }
    pyshell.send(Object.assign(funcAndState, appState));

    pyshell.on('stderr', function(stderr){
        console.log("PYTHON ERROR: ", stderr);
    });

    // let someNum = 0
    pyshell.on('message', function(message) {
        // The below commented-out illustrates that this function instance stick around -
        // someNum keeps the same context upon multiple iterations of message.  Not sure how,
        // investigate later

        // console.log("now: ", Date.now());
        // console.log('somenum: ', someNum);
        // someNum += 1;
        if (message.type === "log"){
            console.log("Python Message:");
            console.log("\t" + JSON.stringify(message.message));
        } else if (message.type === "progMsg"){
            console.log("\t" + "progMsg:");
            console.log("\t\t" + JSON.stringify(message.message));
            progressToastText.innerText = message.message;
        } else if (message.type === "prog"){
            progressBar.setAttribute("aria-valuenow", message.message);
            progressBar.style.width = `${message.message}%`;

            if (message.message === "100"){
                progressBarFinished();
                for (let i = 0; i < generateButtons.length; i++){
                    generateButtons[i].removeAttribute("disabled");
                }

                pyshell.end(function(err, code, signal){
                    if (err) throw err;
                        console.log('The exit code was: ' + code);
                        console.log('The exit signal was: ' + signal);
                        console.log('finished');
                });
                
            }
        } else if (message.type === "requireInput"){
            if ("unsupportedVals" in message.message){
                let modalBod = document.getElementById("fixUnsupportedEntrustmentsModalBody");
                let inputGroupDiv = document.createElement('div');
                inputGroupDiv.classList.add('input-group', 'mb-3');

                let selectedChoices = [];

                for (let i = 0; i < message.message["unsupportedVals"].length; i++){
                    let prependDiv = document.createElement('div');
                    prependDiv.classList.add('input-group-prepend');

                    let label = document.createElement('label');
                    label.classList.add('input-group-text');
                    label.setAttribute('for', `inputGroupSelect${i}`);
                    label.innerText = message.message["unsupportedVals"][i];

                    prependDiv.appendChild(label);

                    let selection = document.createElement('select');
                    selection.classList.add('custom-select');
                    selection.id = `inputGroupSelect${i}`;

                    let opt0 = document.createElement('option');
                    opt0.selected = true;
                    opt0.innerText = "Choose an Entrustment Level"

                    let opt1 = document.createElement('option');
                    opt1.value = "1. Intervention";
                    opt1.innerText = "1. Intervention";
                        
                    let opt2 = document.createElement('option');
                    opt2.value = "2. Direction";
                    opt2.innerText = "2. Direction";

                    let opt3 = document.createElement('option');
                    opt3.value = "3. Support";
                    opt3.innerText = "3. Support";

                    let opt4 = document.createElement('option');
                    opt4.value = "4. Competent";
                    opt4.innerText = "4. Competent";

                    let opt5 = document.createElement('option');
                    opt5.value = "5. Proficient";
                    opt5.innerText = "5. Proficient";

                    let opt6 = document.createElement('option');
                    opt6.value = "";
                    opt6.innerText = "<Blank>";

                    [opt0, opt1, opt2, opt3, opt4, opt5, opt6].forEach((opt) => {
                        selection.appendChild(opt);
                    });

                    inputGroupDiv.appendChild(prependDiv);
                    inputGroupDiv.appendChild(selection);
                    selectedChoices.push([label, selection]);
                }

                modalBod.appendChild(inputGroupDiv);
                let btnSaveUnsupportedValFixes = document.getElementById("btnSaveUnsupportedValFixes");

                // This syntax below of declaring the function within the same scope,
                // and then setting it as a value with bound arguments to a wrapper function
                // on addEventListener is done to allow the removeEventListener function inside 
                // fixUnsupportedVals to work.  Anonymous functions (used to pass arguments
                // to functions in addEventListener) cannot be removed, so a named function with
                // bound arguments is needed.  The listener needs to be removed so that fresh parameters
                // can be passed in, since each time the event listener is added it keeps the context of whenever it was
                // declared, and in the case of the pyshell argument was trying to send messages after the pyshell 
                // instance had terminated.  Now, a new pyshell instance is passed each time an event listener is added.
                function fixUnsuportedVals(ps, selChoices){
                    let choicesJSON = {"selectedChoices": {}};
                
                    selChoices.forEach((pair) => {
                        // the labels inner text is key, the selection value is value
                        choicesJSON.selectedChoices[pair[0].innerText] = pair[1].value;
                    });
                    ps.send(choicesJSON);
                    $('#fixUnsupportedEntrustmentsModal').modal('hide');
                    unsupportedValsModalReset();
                    btnSaveUnsupportedValFixes.removeEventListener('click', wrapperFunc);
                }

                btnSaveUnsupportedValFixes.addEventListener("click", wrapperFunc = fixUnsuportedVals.bind(btnSaveUnsupportedValFixes, pyshell, selectedChoices));

                $('#fixUnsupportedEntrustmentsModal').modal({backdrop: 'static'});
            }

        }
    });
}