// const xlsx = require("node-xlsx");
const ExcelJS = require("./excel4node/excel4node"); // const the excel4node package
const fs = require("fs");
const blueCode = "#2b6bd2";
const greenCode = "#79b088";
const redCode = "#BD0303";
const {
  arrAgeingDatas,
  arrFilter,
  arrItemList,
  arrMainTitle,
  blnMultipleSheet,
  intUserID,
  objColumns,
  strMainTittle,
  strSheetName,
} = require("./json");
// const errHandler from "../../../core/common/helpers/errHandler";
// const { logoDownload } from "./logoDownload";

/**
 * ... (other parameters remain the same)
 */
async function getJSONToExcel({
  arrItemList = [],
  objColumns = null,
  arrFilter = [],
  arrMainTitle = null,
  strMainTittle = null,
  intUserID = 0,
  strSheetName = [],
  arrAgeingDatas = [],
  blnMultipleSheet = false,
} = {}) {
  try {
    // ... (previous code remains the same)

    // // Initialize the Excel workbook and worksheet
    // const workbook = new ExcelJS.Workbook();
    // const worksheet = workbook.addWorksheet("Sheet1");

    // Initialize a style for bold text
    // const boldStyle = workbook1.createStyle({
    //   font: {
    //     bold: true,
    //   },
    // });

    let arrItemLists = [];
    let arrObjColumns = [];
    let arrFilters = [];
    let arrMainTitles = [];
    let arrStrMainTittles = [];
    let arrstrSheetName = [];
    let arrForAgeingDatas = [];

    if (!blnMultipleSheet) {
      arrItemLists.push(arrItemList);
      arrObjColumns.push(objColumns);
      arrFilters.push(arrFilter);
      arrMainTitles.push(arrMainTitle);
      arrStrMainTittles.push(strMainTittle);
      arrstrSheetName.push(strSheetName?.length ? strSheetName : "sheet");
      arrForAgeingDatas.push(arrAgeingDatas);

      //Get all details in arrays
      arrItemList = arrItemLists;
      objColumns = arrObjColumns;
      arrFilter = arrFilters;
      arrMainTitle = arrMainTitles;
      strMainTittle = arrStrMainTittles;
      strSheetName = arrstrSheetName;
      arrAgeingDatas = arrForAgeingDatas;
    }
    let arrExcelExportData = [];
    //Iteratre through all data
    const workbook = new ExcelJS.Workbook();
    // Initialize a style for bold text
    const boldStyle = ({ color = "black", bold = false, objAlignment = {} }) => {
      const objStyle = {
        font: {
          bold,
          color: color,
        },
      };
      if (Object.keys(objAlignment).length) Object.assign(objStyle, { alignment: objAlignment });
      console.log(objStyle);
      return workbook.createStyle(objStyle);
    };
    for (let i = 0; i < arrItemList.length; i++) {
      const worksheet = workbook.addWorksheet(strSheetName[i] || "Sheet1");

      //Array of Arrays for Excel Data
      let arrExcelFormatList = [];
      let arrExcelFilterList = [];
      //Array of keys from objColounm
      let arrExcelKeys = [];
      //Array of Headings of Colounms
      let arrExcelHeaders = [];
      //Array of sub Headings of Colounms
      let arrExcelSubHeaders = [];
      let arrMerge = [];
      //Set Colounm Width based on length of heading
      //Set merging of colum and row
      // merge A2:A3 => "!merges" = ({ s: { r: 1, c: 0 }, e: { r: 2, c: 0 } })

      // Initialize a style for bold text
      // const boldStyle = workbook.createStyle({
      //   font: {
      //     bold: true,
      //   },
      // });

      // let headingRow = arrFilter.length + 1; // Set the heading row
      // let headingColumn = 1; // Set the heading column
      // let row = headingRow+ 1; // Set the data starting row
      // let column = 3; // Set the data starting column
      let column = 1;
      let row = 1;
      let objExcelOptions = {
        "!cols": [],
        "!merges": [],
      };
      if (objColumns[i]) {
        let mergeIndex = 0;
        let intColumnIndex = 1;
        //Make Formatted array from arrItemList based on objColounms
        objColumns[i][Object.keys(objColumns[i])[0]].forEach((objItem) => {
          if (objItem["blnShow"] == "true" || (objItem["blnShow"] == true && objItem["strKey"] != "dblRoe")) {
            let maxWidth = objItem["strHeader"]?.length + 5;
            for (let rowData of arrItemList[i]) {
              const cellContent = rowData[objItem["strKey"]];
              const contentWidth = cellContent ? cellContent.toString().length + 5 : 0;
              maxWidth = Math.max(maxWidth, contentWidth);
            }

            // Set the column width to the maximum content width
            worksheet.column(intColumnIndex).setWidth(maxWidth);

            intColumnIndex++;
            //set Coloumn width of sub headings
            if (objItem["arrChild"]?.length) {
              objItem["arrChild"].forEach((objItemChild) => {
                /**
                 * pushing subHeader data to arrExcelSubHeaders
                 */
                arrExcelSubHeaders.push({
                  strHeader: objItemChild["strHeader"],
                  intColSpan: objItemChild["intColSpan"] || 1,
                  intRowSpan: objItemChild["intRowSpan"] || 1,
                });

                arrExcelKeys.push(objItemChild["strKey"]);
                // objExcelOptions["!cols"].push({ wch: objItemChild["strHeader"].length + 3 });
                // worksheet.column(intColumnIndex).setWidth(objItemChild["strHeader"].length + 3);
                // objExcelOptions["!cols"].push({ wch: objItemChild["strHeader"].length + 3 });
              });
            } else {
              arrExcelSubHeaders.push({ strHeader: "" });
              arrExcelKeys.push(objItem["strKey"]);

              // worksheet
              //   .column(intColumnIndex)
              //   .setWidth(objItem["intExcelWidth"] ? objItem["intExcelWidth"] : objItem["strHeader"].length + 3);
              objExcelOptions["!cols"].push({
                wch: objItem["intExcelWidth"] ? objItem["intExcelWidth"] : objItem["strHeader"].length + 3,
              });
            }
            //Set Merging of column
            arrExcelHeaders.push({
              strHeader: objItem["strHeader"],
              intColSpan: objItem["intColSpan"] || 1,
              intRowSpan: objItem["intRowSpan"] || 1,
            });

            // arrMerge.push({
            //   s: { c: mergeIndex, r: 0 },
            //   e: { c: mergeIndex - 1 + objItem["intColSpan"], r: objItem["intRowSpan"] - 1 },
            // });
            for (let index = 1; index < objItem["intColSpan"]; index++) {
              arrExcelHeaders.push({ strHeader: "" });
            }
            // mergeIndex = mergeIndex + objItem["intColSpan"];
          }
        });
      }
      //main header settings
      for (let arrTitle of arrMainTitle[i]) {
        worksheet
          .cell(row, arrExcelHeaders.length / 2)
          .string(arrTitle[0])
          .style(boldStyle({ bold: true, color: blueCode }));
        row++;
      }
      //empty row
      worksheet.cell(row++).string("");
      row++;
      // setting filter data
      if (arrFilter[i]) {
        let intTotalCols = arrExcelHeaders.length;
        let arrFilterLeft = arrFilter[i].filter((obj) => obj["strPos"] == "Left");
        let arrFilterRight = arrFilter[i].filter((obj) => obj["strPos"] == "Right");
        let intFilterCount =
          arrFilterLeft.length > arrFilterRight.length ? arrFilterLeft.length : arrFilterRight.length;
        for (let intPos = 0; intPos < intFilterCount; intPos++) {
          if (arrFilterLeft[intPos]) {
            worksheet
              .cell(row, column++)
              .string(arrFilterLeft[intPos]["strName"])
              .style(boldStyle({ bold: true }));
            worksheet.cell(row, column++).string(arrFilterLeft[intPos]["strValue"]);
          }
          if (arrFilterRight[intPos]) {
            worksheet
              .cell(row, intTotalCols - 2)
              .string(arrFilterRight[intPos]["strName"])
              .style(boldStyle({ bold: true }));
            worksheet.cell(row, intTotalCols - 1).string(arrFilterRight[intPos]["strValue"]);
          }
          row++;
          column = 1;
        }
        row++;
      }

      //setting main Headers
      for (let objHeader of arrExcelHeaders) {
        if (objHeader["intColSpan"] && objHeader["intColSpan"] > 1) {
          const endColumn = column + objHeader["intColSpan"] - 1;
          worksheet
          .cell(row, column++, row, endColumn, true)
          .string(objHeader["strHeader"])
          .style(boldStyle({ bold: true, color: blueCode }));

        } else if (objHeader["intRowSpan"] && objHeader["intRowSpan"] > 1) {
          console.log({objHeader});

          const endRow = row + objHeader["intRowSpan"] - 1;
          worksheet
            .cell(row, column, endRow, column++, true)
            .string(objHeader["strHeader"])
            .style(boldStyle({ bold: true, color: blueCode }));
        } else
          worksheet
            .cell(row, column++)
            .string(objHeader["strHeader"])
            .style(boldStyle({ bold: true, color: blueCode }));
      }
      row++;
      column = 1;
      //setting subHeaders

      if (arrExcelSubHeaders.length) {
        for (let objSubHeader of arrExcelSubHeaders) {
          // const intTempColumn = column++;
          worksheet
            .cell(row, column++)
            .string(objSubHeader["strHeader"])
            .style(boldStyle({ bold: true }));
        }
        row++;
        column = 1;
      }

      arrExcelFormatList.push([]);
      // Set filter data
      arrExcelFormatList = arrExcelFormatList.concat(arrExcelFilterList);
      arrExcelFormatList.push([]);
      //set Headings for colounms
      let intheaderIndex = arrExcelFormatList.length;
      //Push column details for merging
      // arrMerge.forEach((obj) => {
      //   objExcelOptions["!merges"].push({
      //     s: { c: obj["s"]["c"], r: obj["s"]["r"] + intheaderIndex },
      //     e: { c: obj["e"]["c"], r: obj["e"]["r"] + intheaderIndex },
      //   });
      // });

      arrExcelFormatList.push(arrExcelHeaders);

      arrItemList[i].forEach((objItem) => {
        let objExcelFormatedItem = [];

        arrExcelKeys.forEach((strKeyItem) => {
          if(typeof objItem[strKeyItem] === 'undefined')
          objItem[strKeyItem] = '' 
          if (strKeyItem.includes("_#cur")) {
            const formatString = getFormatString(objItem[strKeyItem]);
            worksheet
              .cell(row, column++)
              .number(objItem[strKeyItem])
              .style({
                numberFormat: formatString,
                font: {
                  bold: !!objItem["blnBold"],
                },
              });
          } else {
            if (typeof objItem[strKeyItem] === "string")
              worksheet
                .cell(row, column++)
                .string(objItem[strKeyItem])
                .style({
                  font: {
                    bold: !!objItem["blnBold"],
                  },
                });
            if (typeof objItem[strKeyItem] === "number") {
              const formatString = getFormatString(objItem[strKeyItem]);
              worksheet
                .cell(row, column++)
                .number(objItem[strKeyItem])
                .style({
                  numberFormat: formatString,
                  font: {
                    bold: !!objItem["blnBold"],
                  },
                });
            }
          }
        });
        row++;
        column = 1;
      });
      if (arrAgeingDatas.length && arrAgeingDatas[i].length > 0) {
        let arrageingDataForPush = arrAgeingDatas[i];
        for (let j of arrageingDataForPush.length) {
          arrExcelFormatList.push(arrageingDataForPush[j]);
        }
      }
      // for (let header of arrExcelHeaders) {
      //   worksheet.cell(headingRow, headingColumn++).string(header).style(boldStyle);
      // }
      const excelBuffer = await workbook.writeToBuffer();
      arrExcelExportData.push({ name: strSheetName[i], data: excelBuffer, options: objExcelOptions });
    }
    console.log(Buffer.concat(arrExcelExportData.map((sheet) => sheet.data)));
    fs.writeFileSync("outputsamp.xlsx", Buffer.concat(arrExcelExportData.map((sheet) => sheet.data)));
    return {
      type: "file",
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Last-Modified": new Date().toUTCString(),
      },
      body: Buffer.concat(arrExcelExportData.map((sheet) => sheet.data)),
    };
  } catch (error) {
    console.log({ as: "......................................", error });
  }
}

function getFormatString(num = 0) {
  const strLength = num % 1 != 0 ? num.toString().split(".")[1].length : 2;
  return "#,##0.".padEnd(6 + strLength, "0");
}

getJSONToExcel({
  arrItemList,
  objColumns,
  arrFilter,
  arrMainTitle,
  strMainTittle,
  intUserID,
  strSheetName,
  arrAgeingDatas,
  blnMultipleSheet,
})
  .then((res) => console.log({ res }))
  .catch((e) => console.error(e));

// console.log(s);
/**
 * ... (other functions remain the same)
 */
