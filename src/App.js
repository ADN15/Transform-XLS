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
      let workbook;
      if (file.name.endsWith('.xls')) {
        workbook = XLSX.read(data, { type: "binary" });
      } else if (file.name.endsWith('.xlsx')) {
        workbook = XLSX.read(data, { type: "binary" });
      } else {
        // Handle unsupported file format
        console.error("Unsupported file format");
        return;
      }

      workbook.SheetNames.forEach(function(sheetName) {
          // Here is your object
          console.log("this is sheetname");
          console.log(sheetName);
          var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
          var json_object = JSON.stringify(XL_row_object);
          var rowData = JSON.parse(json_object);
          console.log("currently code result: ");
          console.log(rowData);
      });

      const sheetName = workbook.SheetNames[0];
      console.log("this is sheetname getting from SheetNames[0]");
      console.log(sheetName);
      const sheet = workbook.Sheets[sheetName];
      console.log("this is sheet getting from workbook.Sheets[sheetName]");
      console.log(sheetName);
      const parsedData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      console.log("my code before transform result: ");
      console.log(parsedData);

      var sheetNames =[];
      console.log("declare new var sheetNames[]");
      console.log(sheetNames);
      var sheetData = {};

      const today = new Date();
      const currentMonth = today.getMonth() + 1; // Months are zero-based in JavaScript (January is 0)
      const currentYear = today.getFullYear();

      const revisedAndEstimatedYears = [
          currentMonth <= 3 ? currentYear - 1 : currentYear,
          currentMonth <= 3 ? currentYear : currentYear + 1
      ];

      let summaryRevise = 0;
      let summaryEstimate = 0;
      let tempFundPod = "";

      const exportData = parsedData.slice(1).map((row, rowIndex) => {
        const getMissing = XLSX.utils.sheet_to_json(sheet, { range: 1, defval: "" })[rowIndex];
        const revisedCFYString = String(row["Revised-CFY"]).replace(",", "");
        let revisedCFY = parseFloat(revisedCFYString);
        let stsCFY = false;
        let reason = "Pass validation";
        let rowNumber = rowIndex + 3;

        if (typeof revisedCFY !== 'number' || isNaN(revisedCFY)) {
          if (revisedCFYString === "undefined" || revisedCFYString === "") {
              revisedCFY = 0;
          } else {
              revisedCFY = 0;
              stsCFY = true;
              reason = "Numeric only";
          }
      } else if (revisedCFY < 0 || revisedCFY % 100 !== 0 || !/^[0-9]+$/.test(revisedCFY.toString())) {
          revisedCFY = 0;
          stsCFY = true;
          reason = revisedCFY < 0 ? "Cannot be negative" : "Nearest 100s, No decimal";
      }

      if (rowNumber === 3) {
          tempFundPod = getMissing["Funding Pot"];
          summaryRevise += revisedCFY;
      } else {
          if (tempFundPod === getMissing["Funding Pot"]) {
              summaryRevise += revisedCFY;
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
          const estimatedNFYString = String(row["Estimated-NFY"]).replace(",", "");
          let estimatedNFY = parseFloat(estimatedNFYString);
          let stsNFY = false;
          let reason = "Pass Validation";
          let rowNumber = rowIndex + 3;

          if (typeof estimatedNFY !== 'number' || isNaN(estimatedNFY)) {
            if (estimatedNFYString === "undefined" || estimatedNFYString === "") {
                estimatedNFY = 0;
            } else {
                estimatedNFY = 0;
                stsNFY = true;
                reason = "Numeric only";
            }
          } else if (estimatedNFY < 0 || estimatedNFY % 100 !== 0 || !/^[0-9]+$/.test(estimatedNFY.toString())) {
              estimatedNFY = 0;
              stsNFY = true;
              reason = estimatedNFY < 0 ? "Cannot be negative" : "Nearest 100s, No decimal";
          }

          if (rowNumber === 3) {
              tempFundPod = getMissing["Funding Pot"];
              summaryEstimate += estimatedNFY;
          } else {
              if (tempFundPod === getMissing["Funding Pot"]) {
                  summaryEstimate += estimatedNFY;
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
      /*
      const headers = ["MINVIEW", "Budget", "Account", "Date", "Version", "Amount", "status", "remark", "summary revise", "summary estimate", "row"];
      extendedExportData.unshift(headers);

      const wsData = extendedExportData.map(row => Object.values(row));
      const ws = XLSX.utils.aoa_to_sheet(wsData);

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      XLSX.writeFile(wb, "exported_data.xlsx");
      */

      var tempSheetName = 'Drawdown_Table';

      //manualy add data
      sheetNames.push(tempSheetName);
      sheetData[tempSheetName]=extendedExportData;

      console.log("disini mau push sheetNames.push(sheetName)");
      console.log(sheetName);
      sheetNames.push(sheetName);
      sheetData[sheetName]=extendedExportData;

      console.log("showing all sheetData[]");
      console.log(sheetData);

      setData(extendedExportData);
      console.log("after transform result: ");
      console.log(extendedExportData);
    };

    reader.readAsBinaryString(file);
  };

  return (
    <div className="App">
      <input
        type="file"
        accept=".xlsx, .xls"
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
              <th>Status</th>
              <th>Remark</th>
              <th>Summary Revise</th>
              <th>Summary Estimate</th>
              <th>Row</th>
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
                <td>{row.Status}</td>
                <td>{row.Remark}</td>
                <td>{row.summaryRevise}</td>
                <td>{row.summaryEstimate}</td>
                <td>{row.Row}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

export default App;
