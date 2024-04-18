import { useState } from "react";
import * as XLSX from "xlsx";

import './App.css';

function App() {
  const [data, setData] = useState([]);
  const [parsedDataString, setParsedDataString] = useState("");
  /* const [selectedRow, setSelectedRow] = useState(null); */

  const getCurrentYear = () => {
    const today = new Date();
    const currentMonth = today.getMonth() + 1; // Months are zero-based in JavaScript (January is 0)
    const currentYear = today.getFullYear();

    return currentMonth <= 3 ? currentYear - 1 : currentYear;
  }

  const getNextYear = () => {
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1; // Months are zero-based in JavaScript (January is 0)

    return currentMonth <= 3 ? currentYear : currentYear + 1;
  }

  const processFileDuplicate = (data) => {
    const XLSX = require('xlsx');
    const workbook = XLSX.read(data, { type: "binary" });
    const sheetName = workbook.SheetNames[0];
	  const sheet = workbook.Sheets[sheetName];
    const parsedData = XLSX.utils.sheet_to_json(sheet, {defval: ""});
    console.log("parseData:");
    console.log(parsedData);

    const revisedCFY = parsedData[0];
    const revisedAndEstimatedYears = [];

    revisedAndEstimatedYears.push(getCurrentYear());
    revisedAndEstimatedYears.push(getNextYear());

    // Prepare data for export
    var exportData = parsedData.slice(1).map((row, rowIndex) => {
      var getMissing = XLSX.utils.sheet_to_json(sheet, { range: 1, defval: "" })[rowIndex];
      // Convert to string and check if it contains a comma
      const revisedCFYString = String(row["Revised-CFY"]).replace(",", ""); // Replace comma with empty
      let  revisedCFY = parseFloat(revisedCFYString); // Parse as float
      let stsCFY = false;
      let reason = "Pass validation";
      let rowNumber = rowIndex + 3

      // Validate the amount according to the rules
      if (typeof revisedCFY !== 'number' || isNaN(revisedCFY)) {
          //validate to check is cell have value or not ( blank will pass )
          if(revisedCFYString.trim() === ""){
              revisedCFY = 0; // If not a valid number or not ending with 00, set amount to 0
              stsCFY = true;
              reason = "Pass Validation";
          }else{
              revisedCFY = 0; // If not a valid number or not ending with 00, set amount to 0
              stsCFY = true;
              reason = "Numeric only";
          }
      }

      else if (revisedCFY < 0) {
          revisedCFY = 0; // If not a valid number or not ending with 00, set amount to 0
          stsCFY = true;
          reason = "Cannot be negative";
      }

      else if (revisedCFY % 100 !== 0 || !/^[0-9]+$/.test(revisedCFY.toString())) {
          revisedCFY = 0; // If not a valid number or not ending with 00, set amount to 0
          stsCFY = true;
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
            let estimatedNFY =  parseFloat(estimatedNFYString); // Parse as float;
            let stsNFY = false;
            let reason = "Pass Validation";
            let rowNumber = rowIndex + 3

            // Validate the amount according to the rules
            if (typeof estimatedNFY !== 'number' || isNaN(estimatedNFY)) {
                //validate to check is cell have value or not ( blank will pass )
                if(estimatedNFYString.trim() === ""){
                    estimatedNFY = 0; // If not a valid number or not ending with 00, set amount to 0
                    stsNFY = true;
                    reason = "Pass Validation";
                }else{
                    estimatedNFY = 0; // If not a valid number or not ending with 00, set amount to 0 
                    stsNFY = true;
                    reason = "Numeric only";
                }
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
    

    console.log(extendedExportData);

     // Add custom formatting for headers
     const headers = ["MINVIEW", "Budget", "Account", "Date", "Version", "Amount", "status", "remark", "row"];
     extendedExportData.unshift(headers);
 
     const wsData = extendedExportData.map(row => Object.values(row));
     const ws = XLSX.utils.aoa_to_sheet(wsData);
 
     const wb = XLSX.utils.book_new();
     XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
     XLSX.writeFile(wb, "exported_data.xlsx");


     // Initialize an object to store the summary
      const summary = {};

      // Iterate through extendedExportData starting from index 1
      for (let i = 1; i < extendedExportData.length; i++) {
          const data = extendedExportData[i];
          const key = `${data.Budget}_${data.Date}`;
          
          // If the key doesn't exist in the summary, initialize it with 0
          if (!summary[key]) {
              summary[key] = 0;
          }
          
          // Add the Amount to the corresponding summary key
          summary[key] += data.Amount;
      }

      // Output the summary
      console.log(summary);

    return extendedExportData;
  }

  const handleFileUpload = (e) => {
    const reader = new FileReader();
    reader.readAsBinaryString(e.target.files[0]);
    reader.onload = (e) => {
      const data = e.target.result;
      processFileDuplicate(data);
    };
  }

  return (
    <div className="App">
      <input 
        type="file" 
        accept=".xlsx, .xls" 
        onChange={handleFileUpload} 
      />


      {parsedDataString && (
        <div>
          <p>Values from parsedData (stringified JSON):</p>
          <pre>{parsedDataString}</pre>
        </div>
      )}

    </div>
  );
}

export default App;
