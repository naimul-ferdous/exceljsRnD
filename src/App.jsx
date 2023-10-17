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
    titleCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    titleCell.font = { bold: true, size: 16 };

    worksheet.mergeCells("A3:A4");
    const col1HeadingCell = worksheet.getCell("A3");
    col1HeadingCell.value = "Activities";
    col1HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col1HeadingCell.font = { bold: false, size: 10 };

    col1HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    worksheet.mergeCells("B3:B4");
    const col2HeadingCell = worksheet.getCell("B3");
    col2HeadingCell.value =
      "Expense Items [give details: specification, rates, numbers, etc.]";
    col2HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col2HeadingCell.font = { bold: false, size: 10 };

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
    col3HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col3HeadingCell.font = { bold: false, size: 10 };

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
    col4HeadingCell.value =
      "Expense Items [give details: specification, rates, numbers, etc.]";
    col4HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col4HeadingCell.font = { bold: false, size: 10 };

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
    col5HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col5HeadingCell.font = { bold: false, size: 10 };

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

    worksheet.mergeCells("F3:G3");
    const col6HeadingCell = worksheet.getCell("F3");
    col6HeadingCell.value = "Year 1";
    col6HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col6HeadingCell.font = { bold: false, size: 10 };

    col6HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col6HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    const col7HeadingCell = worksheet.getCell("F4");
    col7HeadingCell.value = "Quantity";
    col7HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col7HeadingCell.font = { bold: false, size: 10 };

    col7HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col7HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    const col8HeadingCell = worksheet.getCell("G4");
    col8HeadingCell.value = "Cost (BDT)";
    col8HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col8HeadingCell.font = { bold: false, size: 10 };

    col8HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col8HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    worksheet.mergeCells("H3:I3");
    const col16HeadingCell = worksheet.getCell("H3");
    col16HeadingCell.value = "Year 2";
    col16HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col16HeadingCell.font = { bold: false, size: 10 };

    col16HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col16HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    const col17HeadingCell = worksheet.getCell("H4");
    col17HeadingCell.value = "Quantity";
    col17HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col17HeadingCell.font = { bold: false, size: 10 };

    col17HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col17HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    const col18HeadingCell = worksheet.getCell("I4");
    col18HeadingCell.value = "Cost (BDT)";
    col18HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col18HeadingCell.font = { bold: false, size: 10 };

    col18HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col18HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    worksheet.mergeCells("J3:K3");
    const col116HeadingCell = worksheet.getCell("J3");
    col116HeadingCell.value = "Total";
    col116HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col116HeadingCell.font = { bold: false, size: 10 };

    col116HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col116HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    const col117HeadingCell = worksheet.getCell("J4");
    col117HeadingCell.value = "Quantity";
    col117HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col117HeadingCell.font = { bold: false, size: 10 };

    col117HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col117HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    const col118HeadingCell = worksheet.getCell("K4");
    col118HeadingCell.value = "Cost (BDT)";
    col118HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col118HeadingCell.font = { bold: false, size: 10 };

    col118HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col118HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    worksheet.mergeCells("L3:L4");
    const col1116HeadingCell = worksheet.getCell("L3");
    col1116HeadingCell.value = "Category (Goods/Works/Services)";
    col1116HeadingCell.alignment = {
      vertical: "top",
      horizontal: "center",
      wrapText: true,
    };
    col1116HeadingCell.font = { bold: false, size: 10 };

    col1116HeadingCell.border = {
      top: { style: "thin", color: { argb: "000000" } },
      left: { style: "thin", color: { argb: "000000" } },
      bottom: { style: "thin", color: { argb: "000000" } },
      right: { style: "thin", color: { argb: "000000" } },
    };

    // Set background color
    col1116HeadingCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF2CC" },
    };

    worksheet.columns = [
      { key: "A", width: 10 },
      { key: "B", width: 20 },
      { key: "C", width: 10 },
      { key: "D", width: 30 },
      { key: "E", width: 10 },
      { key: "F", width: 10 },
      { key: "G", width: 10 },
      { key: "H", width: 10 },
      { key: "I", width: 10 },
      { key: "J", width: 10 },
      { key: "K", width: 10 },
      { key: "L", width: 15 },
    ];

    // Add an array of rows
    const rows = [
      [
        "Modernization of teaching & learning environment",
        "A1. Purchase of computers, networks, CCTV (FR), for students- (30)",
        "58469265",
        "VDI (Virtual Desktop Interface)-A Central Server (parallel processors, RAMs, storage), no. of VDI interfaces, no. of monitors, and keyboards; can be operated in LAN without internet, can be operated over internet too. Desktop Computers (VDI), for concurrent user 25",
        "25000",
        "18",
        "450,000.00",
        "0",
        "-",
        "18",
        "450,000.00",
        "Goods",
      ],
    ];
    // add new rows and return them as array of row objects
    const newRows = worksheet.addRows(rows);

    worksheet.getCell("A5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("B5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("C5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("D5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("E5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("F5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("G5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("H5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("I5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("J5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("K5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };
    worksheet.getCell("L5").alignment = {
      wrapText: true,
      vertical: "top",
      horizontal: "left",
    };

    const otherRows = [
      [
        "58469265",
        "Laptops",
        "25000",
        "18",
        "450,000.00",
        "0",
        "-",
        "18",
        "450,000.00",
        "Goods",
      ],

      [
        "58469265",
        "Laptops",
        "25000",
        "18",
        "450,000.00",
        "0",
        "-",
        "18",
        "450,000.00",
        "Goods",
      ],
      [
        "58469265",
        "Laptops",
        "25000",
        "18",
        "450,000.00",
        "0",
        "-",
        "18",
        "450,000.00",
        "Goods",
      ],
    ];

    // Add an array of rows with inherited style
    // These new rows will have same styles as last row
    // and return them as array of row objects
    const newRowsStyled = worksheet.addRows(otherRows, "i");

    console.log(newRowsStyled);

    worksheet.mergeCells("B5:B8");

    worksheet.mergeCells("A5:A8");

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
