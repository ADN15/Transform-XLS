/*!

JSZip - A Javascript class for generating and reading zip files
<http://stuartk.com/jszip>

(c) 2009-2014 Stuart Knightley <stuart [at] stuartk.com>
Dual licenced under the MIT license or GPLv3. See https://raw.github.com/Stuk/jszip/master/LICENSE.markdown.

JSZip uses the library pako released under the MIT license :
https://github.com/nodeca/pako/blob/master/LICENSE
*/
var getScriptPromisify = (src) => {
    return new Promise((resolve, reject) => {
      const existingScript = document.querySelector(`script[src="${src}"]`);
      if (existingScript) {
        resolve(); // already loaded
        return;
      }
  
      const script = document.createElement('script');
      script.src = src;
      script.onload = () => resolve();
      script.onerror = () => reject(new Error(`Failed to load script: ${src}`));
      document.head.appendChild(script);
    });
  };


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

    class MainWebComponent extends HTMLElement{
        constructor(){
            super();

            //HTML objects
            this.attachShadow({mode:'open'});
            this.shadowRoot.appendChild(template.content.cloneNode(true));
            this._input = this.shadowRoot.querySelector('input');
            this._remove = this.shadowRoot.querySelector('#remove');

    
            //XLS related objects
            this._sheetNames=null; //holds array of Sheet Names
            this._data=null; //holds JSON Array returned from XLS sheet
        }

        /**
         * This method displays the file selector to the end-user by executing the click event on the HTML object stored in the this._input variable
         * The rest of the upload is handled in the onChange() event of the input control stored in the connectedCallback() function. The onChange() event
         * calls the loadCSV() function and passes in the CSV file as a parameter
         */
        showFileSelector(){
            this.handleRemove(); //remove any existing files, required if this action is run multiple times in the same session
            console.log("In ShowFileSelector()");
            this._input.click();
        }

        //retrieve the data in the CSV file
        getData(sheetName){
            return this._data[sheetName];
        }

        getSheetNames(){
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

        setNames(sheetNames){
            this._sheetNames=sheetNames;
        }

        setData(newData){
            this._data=newData;
        }

        async parseExcel(file) {
            
            await getScriptPromisify('https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js');
            const XLSX = window.XLSX;
            //await getScriptPromisify('https://cdn.sheetjs.com/xlsx-0.20.0/package/jszip.min.js');
            const temp = this;

            var reader = new FileReader();

            reader.onload = function(e) {
                var data = e.target.result;
                var workbook = XLSX.read(data, {
                    type: 'array',
                    cellDates: true,
                    cellNF: false,
                    cellText: false,
                    bookVBA: true  // ✅ this is the key!
                });

                var sheetNames =[];
                var sheetData = {};

                var today = new Date();
                var currentMonth = today.getMonth() + 1; // Months are zero-based in JavaScript (January is 0)
                var currentYear = today.getFullYear();

                // updated macro detection
                const hasMacros = !!(workbook.vbaraw || workbook.vbaProject);

                console.log(hasMacros ? "✅ This workbook contains macros." : "❌ This workbook does NOT contain macros.");

                workbook.SheetNames.forEach(function(sheetName) {
                    // Here is your object
                    var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                    var json_object = JSON.stringify(XL_row_object);
                    var rowData = JSON.parse(json_object);
                    console.log("currently code result: ");
                    console.log(rowData);
                });

                var sheetName = workbook.SheetNames[0];
                console.log("sheet Name : "+sheetName);
                var sheet = workbook.Sheets[sheetName];
                var parsedData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
                console.log("my code before transform result: ");
                console.log(parsedData);

                var revisedCFY = parsedData[0];
                var revisedAndEstimatedYears = [];
                revisedAndEstimatedYears.push(currentMonth <= 3 ? currentYear - 1 : currentYear);
                revisedAndEstimatedYears.push(currentMonth <= 3 ? currentYear : currentYear + 1);

                //if (revisedCFY["Estimated NFY"]) {
                //    revisedAndEstimatedYears.push(revisedCFY["Revised CFY"]);
                //    revisedAndEstimatedYears.push(revisedCFY["Estimated NFY"]);
                //} else {
                //    revisedAndEstimatedYears.push(revisedCFY["Revised CFY"]);
                //}

                // Prepare data for export
                var exportData = parsedData.slice(1).map((row, rowIndex) => {
                    var getMissing = XLSX.utils.sheet_to_json(sheet, { range: 1, defval: "" })[rowIndex];
                    // Convert to string and check if it contains a comma
                    const revisedCFYString = String(row["Revised-CFY"]).replace(",", ""); // Replace comma with empty
                    let revisedCFY = parseFloat(revisedCFYString); // Parse as float or default to 0
                    let stsCFY = false
                    let reason = "Pass Validation";
                    let rowNumber = rowIndex + 3

                    // Validate the amount according to the rules
                    //if (revisedCFY < 0 || revisedCFY % 100 !== 0) {
                    //    revisedCFY = 0; // If negative or not ending with 00, set amount to 0
                    //    stsCFY = true;
                    //}

                    // Validate the amount according to the rules
                    if (typeof revisedCFY !== 'number' || isNaN(revisedCFY)) {
                        revisedCFY = 0; // If not a valid number or not ending with 00, set amount to 0
                        stsCFY = true
                        reason = "Numeric only";
                    }

                    else if (revisedCFY < 0) {
                        revisedCFY = 0; // If not a valid number or not ending with 00, set amount to 0
                        stsCFY = true
                        reason = "Cannot be negative";
                    }

                    else if (revisedCFY % 100 !== 0 || !/^[0-9]+$/.test(revisedCFY.toString())) {
                        revisedCFY = 0; // If not a valid number or not ending with 00, set amount to 0
                        stsCFY = true
                        reason = "Nearest 100s, No decimal";
                    }

                    return {
                            MINVIEW: getMissing["Cost Centre"],
                            Budget: getMissing["Funding Pot"],
                            Account:getMissing["Accounts"],
                            //Account: row.Measures.split(" ")[0],
                            Date: revisedAndEstimatedYears[0],
                            Version: "public.Revised",
                            Amount: revisedCFY,
                            Status: stsCFY.toString(),
                            Remark: reason,
                            Row:rowNumber
                    };
                });
            
            
                // Add the second set of data with NextYear and Estimated NFY
                let extendedExportData = [];
                //if(revisedCFY["Estimated NFY"]){
                // Add the second set of data with NextYear and Estimated NFY
                extendedExportData = [
                    ...exportData,
                    ...parsedData.slice(1).map((row, rowIndex) => {
                        const getMissing = XLSX.utils.sheet_to_json(sheet, { range: 1, defval: "" })[rowIndex];
                        // Convert to string and check if it contains a comma
                        const estimatedNFYString = String(row["Estimated-NFY"]).replace(",", ""); // Replace comma with empty
                        let estimatedNFY = parseFloat(estimatedNFYString); // Parse as float or default to 0
                        let stsNFY = false;
                        let reason = "Pass Validation"
                        let rowNumber = rowIndex + 3

                        // Validate the amount according to the rules
                        //if (estimatedNFY < 0 || estimatedNFY % 100 !== 0) {
                        //    estimatedNFY = 0; // If negative or not ending with 00, set amount to 0
                        //    stsNFY = true;
                        //}
                    
                        // Validate the amount according to the rules
                        if (typeof estimatedNFY !== 'number' || isNaN(estimatedNFY)) {
                            estimatedNFY = 0; // If not a valid number or not ending with 00, set amount to 0
                            stsNFY = true
                            reason = "Numeric only";
                        }

                        else if (estimatedNFY < 0 ) {
                            estimatedNFY = 0; // If not a valid number or not ending with 00, set amount to 0
                            stsNFY = true;
                            reason = "Cannot be negative";
                        }

                        else if (estimatedNFY % 100 !== 0 || !/^[0-9]+$/.test(estimatedNFY.toString())) {
                            estimatedNFY = 0; // If not a valid number or not ending with 00, set amount to 0
                            stsNFY = true;
                            reason = "Nearest 100s, No decimal";
                        }


                        return {
                            MINVIEW: getMissing["Cost Centre"],
                            Budget: getMissing["Funding Pot"],
                            Account:getMissing["Accounts"],
                            //Account: row.Measures.split(" ")[0],
                            Date: revisedAndEstimatedYears[1],
                            Version: "public.Estimated",
                            Amount: estimatedNFY,
                            Status: stsNFY.toString(),
                            Remark: reason,
                            Row:rowNumber
                        };
                    })
                ];
                //}else{
                //    extendedExportData = [...exportData];
                //}

                sheetNames.push(sheetName);
                sheetData[sheetName]=extendedExportData;
                console.log("after transform result: ");

                console.log(extendedExportData);

                temp.setData(sheetData);
                temp.setNames(sheetNames);
                temp.handleRemove();
                temp.dispatch('onFileUpload');
            };

            reader.onerror = function(ex) {
                console.log(ex);
            };
      
            //reader.readAsBinaryString(file);
            reader.readAsArrayBuffer(file);
        }
       


    //events

        //triggered when a user removes the Excel file
        handleRemove() {
            const el = this._input;
            const file = el.files[0];
            el.value = "";
            this.dispatch('change', file);
        }
        handleFileSelect(evt) {
            console.log(Date.now()); //prints timestamp to console...for testing purposes only
            var files = evt.target.files; // FileList object
            
            this.setData(this.parseExcel(files[0]));
            
        }

        dispatch(event, arg) {
            this.dispatchEvent(new CustomEvent(event, {detail: arg}));
        }


        /**
         * standard Web Component function used to add event listeners
         */
        connectedCallback(){
            this._input.addEventListener('change',(e)=>this.handleFileSelect(e));
            this._remove.addEventListener('click',()=>this.handleRemove());
        
        }

    }

    window.customElements.define('new-upload-xls',MainWebComponent);
})()