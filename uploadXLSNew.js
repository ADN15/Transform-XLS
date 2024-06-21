/*!

JSZip - A Javascript class for generating and reading zip files
<http://stuartk.com/jszip>

(c) 2009-2014 Stuart Knightley <stuart [at] stuartk.com>
Dual licenced under the MIT license or GPLv3. See https://raw.github.com/Stuk/jszip/master/LICENSE.markdown.

JSZip uses the library pako released under the MIT license :
https://github.com/nodeca/pako/blob/master/LICENSE
*/
var getScriptPromisify = (src) => {
    return new Promise(resolve => {
        fetch(src)
            .then(response => response.text())
            .then(scriptText => {
                const script = document.createElement('script');
                script.textContent = scriptText;
                document.head.appendChild(script);
                resolve();
            })
            .catch(error => {
                console.error(`Failed to load script: ${src}`, error);
                resolve();
            });
    });
}

(function () {
    const template = document.createElement('template');
    template.innerHTML = `
    <style>
    :host {
    font-size: 13px;
    font-family: arial;
    overflow: auto;
    }
    </style>
    <section hidden>
    <article>
    <label for="fileUpload">Upload</label>
        <span></span><button id="remove">Remove</button>
    </article>
    <input hidden id="fileUpload" type="file" accept=".xls,.xlsx,.xlsm" />
    </section>
    `;

    class UploadXLSValidation2 extends HTMLElement {
        constructor() {
            super();

            // HTML objects
            this.attachShadow({ mode: 'open' });
            this.shadowRoot.appendChild(template.content.cloneNode(true));
            this._input = this.shadowRoot.querySelector('input');
            this._remove = this.shadowRoot.querySelector('#remove');

            // XLS related objects
            this._sheetNames = null; // Holds array of Sheet Names
            this._data = null; // Holds JSON Array returned from XLS sheet

            // Ensure event listeners are only added once
            this._boundHandleFileSelect = this.handleFileSelect.bind(this);
            this._boundHandleRemove = this.handleRemove.bind(this);
        }

        showFileSelector() {
            this.handleRemove(); // Remove any existing files
            console.log("In ShowFileSelector()");
            this._input.click();
        }

        getData(sheetName) {
            return this._data[sheetName];
        }

        getSheetNames() {
            return this._sheetNames;
        }

        getCurrentYear = () => {
            const today = new Date();
            const currentMonth = today.getMonth() + 1; // Months are zero-based in JavaScript (January is 0)
            const currentYear = today.getFullYear();
            return currentMonth <= 3 ? currentYear - 1 : currentYear;
        }

        getNextYear = () => {
            const today = new Date();
            const currentYear = today.getFullYear();
            const currentMonth = today.getMonth() + 1; // Months are zero-based in JavaScript (January is 0)
            return currentMonth <= 3 ? currentYear : currentYear + 1;
        }

        getVersionType = (rowYear) => {
            const currentYear = this.getCurrentYear();
            return rowYear === currentYear ? "public.Revised" : "public.Estimated";
        }

        setNames(sheetNames) {
            this._sheetNames = sheetNames;
        }

        setData(newData) {
            this._data = newData;
        }

        async parseExcel(file) {
            let extension = '';
            const fileName = file.name || file.fileName;
            const dotIndex = fileName.lastIndexOf('.');
            if (dotIndex !== -1) {
                extension = fileName.substring(dotIndex + 1).toLowerCase();
            }

            if (!extension) {
                console.error('Unsupported file type');
                return;
            }

            if (extension === 'xlsx') {
                await getScriptPromisify('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/jszip.js');
                await getScriptPromisify('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.js');
            } else if (extension === 'xls') {
                await getScriptPromisify('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js');
            } else {
                console.error('Unsupported file type');
                return;
            }

            const temp = this;
            var reader = new FileReader();

            reader.onload = function (e) {
                var data = e.target.result;
                var workbook;

                if (extension === 'xlsx' || extension === 'xls') {
                    workbook = XLSX.read(data, { type: 'binary' });
                }

                var sheetNames = [];
                var sheetData = {};

                var today = new Date();
                var currentMonth = today.getMonth() + 1; // Months are zero-based in JavaScript (January is 0)
                var currentYear = today.getFullYear();

                workbook.SheetNames.forEach(function (sheetName) {
                    console.log("this is sheetname");
                    console.log(sheetName);
                    var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                    var json_object = JSON.stringify(XL_row_object);
                    var rowData = JSON.parse(json_object);
                    console.log("currently code result: ");
                    console.log(rowData);
                });

                var sheetName = workbook.SheetNames[0];
                console.log("this is sheetname getting from SheetNames[0]");
                console.log(sheetName);
                var sheet = workbook.Sheets[sheetName];
                console.log("this is sheet getting from workbook.Sheets[sheetName]");
                console.log(sheetName);
                var parsedData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
                console.log("my code before transform result: ");
                console.log(parsedData);

                var revisedCFY = parsedData[0];
                var revisedAndEstimatedYears = [];
                revisedAndEstimatedYears.push(currentMonth <= 3 ? currentYear - 1 : currentYear);
                revisedAndEstimatedYears.push(currentMonth <= 3 ? currentYear : currentYear + 1);

                var summaryRevise = 0;
                var summaryEstimate = 0;

                var tempFundPod = "";

                // Prepare data for export
                var exportData = parsedData.slice(1).map((row, rowIndex) => {
                    var getMissing = XLSX.utils.sheet_to_json(sheet, { range: 1, defval: "" })[rowIndex];
                    // Convert to string and check if it contains a comma
                    const revisedCFYString = String(row["Revised-CFY"]).replace(",", ""); // Replace comma with empty
                    let revisedCFY = parseFloat(revisedCFYString); // Parse as float
                    let stsCFY = false;
                    let reason = "Pass validation";
                    let rowNumber = rowIndex + 3

                    // Validate the amount according to the rules
                    if (typeof revisedCFY !== 'number' || isNaN(revisedCFY)) {
                        if (revisedCFYString === "undefined" || revisedCFYString === "") {
                            revisedCFY = 0;
                            stsCFY = true;
                            reason = "Blank";
                        } else {
                            revisedCFY = 0; // If not a valid number or not ending with 00, set amount to 0
                            stsCFY = true;
                            reason = "Numeric only";
                        }
                    } else if (revisedCFY < 0) {
                        revisedCFY = 0; // If not a valid number or not ending with 00, set amount to 0
                        stsCFY = true;
                        reason = "Cannot be negative";
                    } else if (revisedCFY % 100 !== 0 || !/^[0-9]+$/.test(revisedCFY.toString())) {
                        revisedCFY = 0; // If not a valid number or not ending with 00, set amount to 0
                        stsCFY = true;
                        reason = "Nearest 100s, No decimal";
                    }

                    if (rowNumber - 3 === 0) {
                        tempFundPod = getMissing["Funding Pot"];
                        summaryRevise = summaryRevise + revisedCFY;
                    } else {
                        if (tempFundPod === getMissing["Funding Pot"]) {
                            summaryRevise = summaryRevise + revisedCFY;
                        } else {
                            tempFundPod = getMissing["Funding Pot"];
                            summaryRevise = revisedCFY;
                        }
                    }

                    return {
                        MINVIEW: getMissing["Cost Centre"],
                        Budget: getMissing["Funding Pot"],
                        Account: getMissing["Accounts"],
                        Date: revisedAndEstimatedYears[0],
                        Version: "public.Revised",
                        Amount: revisedCFY,
                        Status: stsCFY.toString(),
                        Remark: reason,
                        summaryRevise: summaryRevise,
                        summaryEstimate: 0,
                        Row: rowNumber
                    };
                });

                let extendedExportData = [
                    ...exportData,
                    ...parsedData.slice(1).map((row, rowIndex) => {
                        const getMissing = XLSX.utils.sheet_to_json(sheet, { range: 1, defval: "" })[rowIndex];
                        const estimatedNFYString = String(row["Estimated-NFY"]).replace(",", ""); // Replace comma with empty
                        let estimatedNFY = parseFloat(estimatedNFYString); // Parse as float;
                        let stsNFY = false;
                        let reason = "Pass Validation";
                        let rowNumber = rowIndex + 3

                        if (typeof estimatedNFY !== 'number' || isNaN(estimatedNFY)) {
                            if (estimatedNFYString === "undefined" || estimatedNFYString === "") {
                                estimatedNFY = 0;
                                stsNFY = true;
                                reason = "Blank";
                            } else {
                                estimatedNFY = 0; // If not a valid number or not ending with 00, set amount to 0 
                                stsNFY = true;
                                reason = "Numeric only";
                            }
                        } else if (estimatedNFY < 0) {
                            estimatedNFY = 0; // If not a valid number or not ending with 00, set amount to 0
                            stsNFY = true;
                            reason = "Cannot be negative";
                        } else if (estimatedNFY % 100 !== 0 || !/^[0-9]+$/.test(estimatedNFY.toString())) {
                            estimatedNFY = 0; // If not a valid number or not ending with 00, set amount to 0
                            stsNFY = true;
                            reason = "Nearest 100s, No decimal";
                        }

                        if (rowNumber - 3 === 0) {
                            tempFundPod = getMissing["Funding Pot"];
                            summaryEstimate = summaryEstimate + estimatedNFY;
                        } else {
                            if (tempFundPod === getMissing["Funding Pot"]) {
                                summaryEstimate = summaryEstimate + estimatedNFY;
                            } else {
                                tempFundPod = getMissing["Funding Pot"];
                                summaryEstimate = estimatedNFY;
                            }
                        }

                        return {
                            MINVIEW: getMissing["Cost Centre"],
                            Budget: getMissing["Funding Pot"],
                            Account: getMissing["Accounts"],
                            Date: revisedAndEstimatedYears[1],
                            Version: "public.Estimated",
                            Amount: estimatedNFY,
                            Status: stsNFY.toString(),
                            Remark: reason,
                            summaryRevise: 0,
                            summaryEstimate: summaryEstimate,
                            Row: rowNumber
                        };
                    })
                ];

                var tempSheetName = 'Drawdown_Table';
                sheetNames.push(tempSheetName);
                sheetData[tempSheetName] = extendedExportData;

                sheetNames.push(sheetName);
                sheetData[sheetName] = extendedExportData;

                console.log("showing all sheetData[]");
                console.log(sheetData);

                console.log("after transform result: ");
                console.log(extendedExportData);

                temp.setData(sheetData);
                temp.setNames(sheetNames);
                temp.handleRemove();
                temp.dispatch('onFileUpload');
            };

            reader.onerror = function (ex) {
                console.log(ex);
            };

            reader.readAsBinaryString(file);
        }

        handleRemove() {
            const el = this._input;
            const file = el.files[0];
            el.value = "";
            this.dispatch('change', file);
        }

        handleFileSelect(evt) {
            console.log(Date.now());
            var files = evt.target.files;
            this.parseExcel(files[0]);
        }

        dispatch(event, arg) {
            this.dispatchEvent(new CustomEvent(event, { detail: arg }));
        }

        connectedCallback() {
            this._input.addEventListener('change', this._boundHandleFileSelect);
            this._remove.addEventListener('click', this._boundHandleRemove);
        }

        disconnectedCallback() {
            this._input.removeEventListener('change', this._boundHandleFileSelect);
            this._remove.removeEventListener('click', this._boundHandleRemove);
        }
    }

    window.customElements.define('com-sap-sample-uploadxlsvalidation2', UploadXLSValidation2);
})();
