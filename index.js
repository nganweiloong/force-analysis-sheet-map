"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const xlsx_1 = __importDefault(require("xlsx"));
const path_1 = __importDefault(require("path"));
const GRAVITY_SHEET = "Gravity2";
const LSA_SHEET = "LSA2";
function getHigherAbsoluteValue(a, b) {
    if (Math.abs(a) > Math.abs(b)) {
        return a;
    }
    else {
        return b;
    }
}
// Define the absolute file path
const filePath = path_1.default.join(__dirname, "test.xlsm");
// Read the XLSX file
const workbook = xlsx_1.default.readFile(filePath);
// Get the sheet names
const gravitySheet = workbook.Sheets[GRAVITY_SHEET];
const lsaSheet = workbook.Sheets[LSA_SHEET];
const gravityData = xlsx_1.default.utils.sheet_to_json(gravitySheet, {
    range: 2,
});
const lsaSheetData = xlsx_1.default.utils.sheet_to_json(lsaSheet, {
    range: 2,
});
let gravityDataOutput = {};
let lsaDataOutput = {};
for (const item of gravityData) {
    gravityDataOutput[`${item.Element}__${item["Element Station"]}`] = item;
}
for (const item of lsaSheetData) {
    lsaDataOutput[`${item.Element}__${item["Element Station"]}`] = item;
}
function generateData(gravityData, lsaData) {
    let dataKeyGravity = Object.keys(gravityData);
    let finalOutput = [];
    dataKeyGravity.forEach(rowKey => {
        const key = rowKey.split("__")[0];
        let rowData = {
            Story: gravityData[rowKey].Story,
            Beam: gravityData[rowKey].Beam,
            UniqueName: gravityData[rowKey]["UniqueName"],
            "Output Case": gravityData[rowKey]["Output Case"],
            "Case Type": gravityData[rowKey]["Case Type"],
            "Step Type": gravityData[rowKey]["Step Type"],
            Station: gravityData[rowKey].Station,
            P: gravityData[rowKey].P,
            V2: gravityData[rowKey].V2,
            V3: gravityData[rowKey].V3,
            T: gravityData[rowKey].T,
            M2: gravityData[rowKey].M2,
            M3: gravityData[rowKey].M3,
            Element: gravityData[rowKey].Element,
            "Element Station": gravityData[rowKey]["Element Station"],
            Location: gravityData[rowKey]["Location"],
        };
        // console.log(lsaData[rowKey]);
        console.log(rowData);
    });
}
//first iteration
// 33971-1__something
generateData(gravityDataOutput, lsaDataOutput);
// function generateData(sheetOne: SheetItem[], sheetTwo: SheetItem[]) {
//   if (sheetOne.length !== sheetTwo.length) {
//     console.log("The row length seems not tally, please check again ;D");
//     return;
//   }
//   const netTable: SheetItem[] = [];
//   for (let index = 0; index < sheetOne.length; index++) {
//     const gravityRowItem = sheetOne[index];
//     const lsaRowItem = sheetTwo[index];
//     let rowData: SheetItem = {
//       Story: gravityRowItem.Story,
//       Beam: gravityRowItem.Beam,
//       UniqueName: gravityRowItem["UniqueName"],
//       "Output Case": gravityRowItem["Output Case"],
//       "Case Type": gravityRowItem["Case Type"],
//       "Step Type": gravityRowItem["Step Type"],
//       Station: gravityRowItem.Station,
//       P: gravityRowItem.P,
//       V2: gravityRowItem.V2,
//       V3: gravityRowItem.V3,
//       T: gravityRowItem.T,
//       M2: gravityRowItem.M2,
//       M3: gravityRowItem.M3,
//       Element: gravityRowItem.Element,
//       "Element Station": gravityRowItem["Element Station"],
//       Location: gravityRowItem["Location"],
//     };
//     netTable.push(rowData);
//   }
// }
// generateData(gravityData, lsaSheetData);
