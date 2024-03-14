import { useState } from "react";
import * as XLSX from "xlsx";

import './App.css';

function App() {
  const [data, setData] = useState([]);
  const [parsedDataString, setParsedDataString] = useState("");
  /* const [selectedRow, setSelectedRow] = useState(null); */

  const processFileDuplicate = (data) => {
    const workbook = XLSX.read(data, { type: "binary" });
    const sheetName = workbook.SheetNames[0];
	  const sheet = workbook.Sheets[sheetName];
    const parsedData = XLSX.utils.sheet_to_json(sheet, {defval: ""});

    console.log(parsedData);

    const revisedCFY = parsedData[0];
    const revisedAndEstimatedYears = [];

    if (revisedCFY["Estimated NFY"]) {
        revisedAndEstimatedYears.push(revisedCFY["Revised CFY"]);
        revisedAndEstimatedYears.push(revisedCFY["Estimated NFY"]);
    } else {
        revisedAndEstimatedYears.push(revisedCFY["Revised CFY"]);
    }

    // Prepare data for export
    const exportData = parsedData.slice(2).map((row, rowIndex) => {
      const getMissing = XLSX.utils.sheet_to_json(sheet, { range: 2, defval: "" })[rowIndex];
       // Convert to string and check if it contains a comma
       const revisedCFYString = String(row["Revised CFY"]).replace(",", ""); // Replace comma with dot
       let revisedCFY = parseFloat(revisedCFYString) || 0; // Parse as float or default to 0
       let stsCFY = false;

       // Validate the amount according to the rules
       if (revisedCFY < 0 || revisedCFY % 100 !== 0) {
        revisedCFY = 0; // If negative or not ending with 00, set amount to 0
        stsCFY = true;
       }

      return {
        Account: row.Measures.split(" ")[0],
        Budget: getMissing["Funding Pot"].split(" ")[0],
        Date: revisedAndEstimatedYears[0],
        MINVIEW: getMissing["Ministry View"].split(" ")[0],
        Version: "public.Revised",
        Amount: revisedCFY,
        status: stsCFY
      };
  });


    let extendedExportData = [];
    if(revisedCFY["Estimated NFY"]){
      // Add the second set of data with NextYear and Estimated NFY
      extendedExportData = [
        ...exportData,
        ...parsedData.slice(2).map((row, rowIndex) => {
        const getMissing = XLSX.utils.sheet_to_json(sheet, { range: 2, defval: "" })[rowIndex];
        // Convert to string and check if it contains a comma
        const estimatedNFYString = String(row["Estimated NFY"]).replace(",", ""); // Replace comma with dot
        let estimatedNFY = parseFloat(estimatedNFYString) || 0; // Parse as float or default to 0
        let stsNFY = false;
        // Validate the amount according to the rules
        if (estimatedNFY < 0 || estimatedNFY % 100 !== 0) {
            estimatedNFY = 0; // If negative or not ending with 00, set amount to 0
            stsNFY = true;
        }
        
          return {
              Account: row.Measures.split(" ")[0],
              Budget: getMissing["Funding Pot"].split(" ")[0],
              Date: revisedAndEstimatedYears[1],
              MINVIEW: getMissing["Ministry View"].split(" ")[0],
              Version: "public.Estimated",
              Amount: estimatedNFY,
              status: stsNFY
          };
          })
      ];
    }else{
      extendedExportData = [...exportData];
    }

    console.log(extendedExportData);

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
