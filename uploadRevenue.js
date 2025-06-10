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

    class UploadRevenueXLSMain extends HTMLElement{
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
          const temp = this;
        
          var reader = new FileReader();
        
          reader.onload = function(e) {
            const data = e.target.result;
            const workbook = XLSX.read(data, {
              type: 'array',
              cellDates: true,
              cellNF: false,
              cellText: false,
              bookVBA: true // ✅ this is the key!
            });
        
            const sheetNames = [];
            const sheetData = {};
            const today = new Date();
            const currentMonth = today.getMonth() + 1;
            const currentYear = today.getFullYear();
        
            // ✅ Check for macros
            const hasMacros = !!(workbook.vbaraw || workbook.vbaProject);
            console.log(hasMacros ? "✅ This workbook contains macros." : "❌ This workbook does NOT contain macros.");
        
            const targetSheetName = "RevenueInput";
        
            // Declare sheet outside if-block
            let sheet = null;
        
            if (workbook.SheetNames.includes(targetSheetName)) {
              sheet = workbook.Sheets[targetSheetName];
        
              const XL_row_object = XLSX.utils.sheet_to_row_object_array(sheet);
              const json_object = JSON.stringify(XL_row_object);
              const rowData = JSON.parse(json_object);
        
              console.log(`Data from "${targetSheetName}" sheet:`);
              console.log(rowData);
            } else {
              console.log(`❌ Sheet "${targetSheetName}" not found. Skipping processing.`);
              return; // 🚫 Stop execution if sheet doesn't exist
            }
        
            const a1Value = sheet['A1'] ? sheet['A1'].v : null;
            const Yr = sheet['C7'] ? sheet['C7'].v : currentYear;
        
            if (a1Value !== "iBudget3RevUploadFile") {
              console.log("❌ Error: The file is not valid");
              return;
            } else {
              const parsedData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
              console.log("Raw parsed data:", parsedData);
        
              const revisedAndEstimatedYears = [];
              revisedAndEstimatedYears.push(currentMonth <= 3 ? currentYear - 1 : currentYear);
              revisedAndEstimatedYears.push(currentMonth <= 3 ? currentYear : currentYear + 1);
        
              // Get export data from row 12 (index 11)
              const exportData = parsedData.slice(10).map((row, rowIndex) => {
                const getMissing = XLSX.utils.sheet_to_json(sheet, { range: 11, defval: "" })[rowIndex];
                const revisedCFYString = String(row["Revised CFY"]).replace(/,/g, "");
                const revisedCFY = parseFloat(revisedCFYString) || 0;
        
                return {
                  MINVIEW: getMissing["CC"],
                  Budget: getMissing["Funding Pot"],
                  Account: getMissing["Account"],
                  Date: Yr,
                  Version: "public.Revised",
                  Amount: revisedCFY
                };
              });
        
              const extendedExportData = [
                ...exportData,
                ...parsedData.slice(10).map((row, rowIndex) => {
                  const getMissing = XLSX.utils.sheet_to_json(sheet, { range: 11, defval: "" })[rowIndex];
                  const estimatedNFYString = String(row["Estimated NFY"]).replace(/,/g, "");
                  const estimatedNFY = parseFloat(estimatedNFYString) || 0;
        
                  return {
                    MINVIEW: getMissing["CC"],
                    Budget: getMissing["Funding Pot"],
                    Account: getMissing["Account"],
                    Date: Yr + 1,
                    Version: "public.Estimated",
                    Amount: estimatedNFY
                  };
                })
              ];
        
              // Save processed data
              sheetNames.push(targetSheetName);
              sheetData[targetSheetName] = extendedExportData;
        
              console.log("✅ Final transformed data:");
              console.log(extendedExportData);
            }
        
            // Call component methods
            temp.setData(sheetData);
            temp.setNames(sheetNames);
            temp.handleRemove();
            temp.dispatch('onFileUpload');
          };
        
          reader.onerror = function(ex) {
            console.log("❌ File read error:", ex);
          };
        
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

    window.customElements.define('upload-revenue-xls',UploadRevenueXLSMain);
})()
