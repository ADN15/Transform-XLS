import React, { useState } from "react";
import * as XLSX from "xlsx";

import './App.css';

function App() {
  const [data, setData] = useState([]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
      const data = e.target.result;
      const fileExtension = file.name.split('.').pop().toLowerCase();
      const supportedFormats = ['xls', 'xlsx', 'xlsm'];

      if (!supportedFormats.includes(fileExtension)) {
        console.error("Unsupported file format");
        return;
      }

      const workbook = XLSX.read(data, { type: "binary" });
      const sheetName = "BudgetDrawdown";

      if (!workbook.SheetNames.includes(sheetName)) {
        alert(`Sheet "${sheetName}" tidak ditemukan dalam file Excel.`);
        return;
      }

      const sheet = workbook.Sheets[sheetName];
      const range = XLSX.utils.decode_range(sheet['!ref']); // ✅ get range
      //const rawParsedData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      // ✅ Filter visually non-empty rows starting from Excel row 12
     const parsedData = [];
      const startRow = 11; // 0-based index for Excel row 12
      let rowCount = 0;

      for (let r = startRow; r <= range.e.r; r++) {
        let hasData = false;

        for (let c = range.s.c; c <= range.e.c; c++) {
          const cellAddress = XLSX.utils.encode_cell({ r, c });
          const cell = sheet[cellAddress];

          if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
            hasData = true;
            break;
          }
        }

        if (hasData) {
          const rowJson = XLSX.utils.sheet_to_json(sheet, {
            range: { s: { r, c: range.s.c }, e: { r, c: range.e.c } },
            header: 1,
            defval: "",
          })[0];

          // Convert array row to object using headers (assumes headers at row 11)
          const headers = XLSX.utils.sheet_to_json(sheet, {
            range: 10,
            header: 1,
            defval: "",
          })[0];

          const rowObj = {};
          headers.forEach((key, i) => {
            rowObj[key] = rowJson[i] ?? "";
          });

          parsedData.push(rowObj);
          rowCount++;
        } else {
          break; // stop at first visually blank row
        }
      }

      console.log(`Found ${rowCount} rows with data starting from row 12.`);

      console.log("Filtered parsedData (non-blank rows from row 12):");
      console.log(parsedData);

      const a1Value = sheet['A1'] ? sheet['A1'].v : null;
      const today = new Date();
      const Yr = sheet['C7'] ? sheet['C7'].v : today.getFullYear();

      if (a1Value !== "iBudget3InputFile") {
        console.log("Error: The file is not valid ❌");
        alert("Error: The file is not valid ❌");
        return;
      }

      let sheetData = {};
      let sheetNames = [];

      const exportData = parsedData.slice(10).map((row, rowIndex) => {
        const getMissing = XLSX.utils.sheet_to_json(sheet, { range: 11, defval: "" })[rowIndex];
        const revisedCFYString = String(row["Revised CFY"]).replace(",", "");
        let revisedCFY = parseFloat(revisedCFYString);

        return {
          MINVIEW: getMissing["CC"],
          Budget: getMissing["Funding Pot"],
          Account: getMissing["Accounts"],
          Date: Yr,
          Version: "public.Revised",
          Amount: revisedCFY
        };
      });

      const extendedExportData = [
        ...exportData,
        ...parsedData.slice(10).map((row, rowIndex) => {
          const getMissing = XLSX.utils.sheet_to_json(sheet, { range: 11, defval: "" })[rowIndex];
          const estimatedNFYString = String(row["Estimated NFY"]).replace(",", "");
          let estimatedNFY = parseFloat(estimatedNFYString);

          return {
            MINVIEW: getMissing["CC"],
            Budget: getMissing["Funding Pot"],
            Account: getMissing["Accounts"],
            Date: Yr + 1,
            Version: "public.Estimated",
            Amount: estimatedNFY
          };
        })
      ];

      const tempSheetName = 'Drawdown_Table';

      sheetNames.push(tempSheetName);
      sheetData[tempSheetName] = extendedExportData;

      sheetNames.push(sheetName);
      sheetData[sheetName] = extendedExportData;

      console.log("Final transformed data:");
      console.log(extendedExportData);

      setData(extendedExportData);
    };

    reader.readAsBinaryString(file);
  };

  return (
    <div className="App">
      <input
        type="file"
        accept=".xlsx, .xls, .xlsm"
        onChange={handleFileUpload}
      />

      {/* Display data if available */}
      {data.length > 0 && (
        <table>
          <thead>
            <tr>
              <th>MINVIEW</th>
              <th>Budget</th>
              <th>Account</th>
              <th>Date</th>
              <th>Version</th>
              <th>Amount</th>
            </tr>
          </thead>
          <tbody>
            {data.map((row, index) => (
              <tr key={index}>
                <td>{row.MINVIEW}</td>
                <td>{row.Budget}</td>
                <td>{row.Account}</td>
                <td>{row.Date}</td>
                <td>{row.Version}</td>
                <td>{row.Amount}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

export default App;
