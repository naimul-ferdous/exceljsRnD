import Excel from "exceljs";
import React, { useCallback, useEffect, useRef, useState } from "react";
import { read, utils, writeFileXLSX } from "xlsx";
import "./App.css";

export default function SheetJSReactHTML() {
  /* the component state is an HTML string */
  const [__html, setHtml] = useState("");
  /* the ref is used in export */
  const tbl = useRef(null);

  /* Fetch and update the state once */
  useEffect(() => {
    (async () => {
      const f = await (
        await fetch("https://sheetjs.com/pres.xlsx")
      ).arrayBuffer();
      const wb = read(f); // parse the array buffer
      const ws = wb.Sheets[wb.SheetNames[0]]; // get the first worksheet
      const data = utils.sheet_to_html(ws); // generate HTML
      setHtml(data); // update state
    })();
  }, []);

  /* get live table and export to XLSX */
  const exportFile = useCallback(() => {
    const elt = tbl.current.getElementsByTagName("TABLE")[0];
    const wb = utils.table_to_book(elt);
    writeFileXLSX(wb, "SheetJSReactHTML.xlsx");
  }, [tbl]);

  const exportTableToXLSX = () => {
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    // Merge 5 columns for the title
    worksheet.mergeCells("A2:E2");
    const titleCell = worksheet.getCell("A2");
    titleCell.value = "Work/Activity, Financial Management Procurement Plan";
    titleCell.alignment = { vertical: "top", horizontal: "center", wrapText: true };
    titleCell.font = { bold: true, size: 16 };

    worksheet.mergeCells("A3:A4");
    const col1HeadingCell = worksheet.getCell("A3");
    col1HeadingCell.value = "Activities";
    col1HeadingCell.alignment = { vertical: "top", horizontal: "center", wrapText: true };
    col1HeadingCell.font = { bold: false, size: 12 };

    col1HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };


    worksheet.mergeCells("B3:B4");
    const col2HeadingCell = worksheet.getCell("B3");
    col2HeadingCell.value = "Expense Items [give details: specification, rates, numbers, etc.]";
    col2HeadingCell.alignment = { vertical: "top", horizontal: "center", wrapText: true };
    col2HeadingCell.font = { bold: false, size: 12 };

    col2HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col2HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    col1HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    worksheet.mergeCells("C3:C4");
    const col3HeadingCell = worksheet.getCell("C3");
    col3HeadingCell.value = "Economic Code";
    col3HeadingCell.alignment = { vertical: "top", horizontal: "center", wrapText: true };
    col3HeadingCell.font = { bold: false, size: 12 };

    col3HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col3HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    worksheet.mergeCells("D3:D4");
    const col4HeadingCell = worksheet.getCell("D3");
    col4HeadingCell.value = "Expense Items [give details: specification, rates, numbers, etc.]";
    col4HeadingCell.alignment = { vertical: "top", horizontal: "center", wrapText: true };
    col4HeadingCell.font = { bold: false, size: 12 };

    col4HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col4HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };


    worksheet.mergeCells("E3:E4");
    const col5HeadingCell = worksheet.getCell("E3");
    col5HeadingCell.value = "Unit Cost (BDT)";
    col5HeadingCell.alignment = { vertical: "top", horizontal: "center", wrapText: true };
    col5HeadingCell.font = { bold: false, size: 12 };

    col5HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col5HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    worksheet.columns = [
      { key: "A", width: 40 },
      { key: "B", width: 32 },
      { key: "C", width: 30 },
      { key: "D", width: 30 },
      { key: "E", width: 30 },
      { key: "F", width: 30 },
      { key: "G", width: 30 },
      { key: "H", width: 30 },
      { key: "I", width: 30 },
      { key: "J", width: 30 },
      { key: "K", width: 30 },
      { key: "L", width: 30 },
    ];

    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "example.xlsx";
      a.click();
    });
  };

  return (
    <>
      <button onClick={exportTableToXLSX}>Export XLSX</button>
      {/* <div ref={tbl} dangerouslySetInnerHTML={{ __html }} /> */}
      <div ref={tbl}>
        <table>
          <thead>
            <tr>
              <td
                style={{ border: "2px solid black", padding: "8px" }}
                colSpan={2}
              >
                Hello
              </td>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style={{ border: "2px solid black", padding: "8px" }}>
                Hello
              </td>
              <td style={{ border: "2px solid black", padding: "8px" }}>
                Therew
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </>
  );
}
