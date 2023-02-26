import xlsx, { Sheet } from "xlsx";
import path from "path";
import { NormalizedSheetItem, SheetItem, SheetItemWithSources } from "./type";
const GRAVITY_SHEET = "Gravity2";
const LSA_SHEET = "LSA2";

function getHigherAbsoluteValue(a: number, b: number) {
  if (Math.abs(a) > Math.abs(b)) {
    return { value: a, s: "Gravity" };
  } else {
    return { value: b, s: "LSA" };
  }
}
// Define the absolute file path
const filePath = path.join(__dirname, "test.xlsm");
// Read the XLSX file
const workbook = xlsx.readFile(filePath);
// Get the sheet names
const gravitySheet = workbook.Sheets[GRAVITY_SHEET];
const lsaSheet = workbook.Sheets[LSA_SHEET];
const gravityData = xlsx.utils.sheet_to_json<SheetItem>(gravitySheet, {
  range: 2,
});
const lsaSheetData = xlsx.utils.sheet_to_json<SheetItem>(lsaSheet, {
  range: 2,
});
let gravityDataOutput: NormalizedSheetItem = {};
let lsaDataOutput: NormalizedSheetItem = {};

for (const item of gravityData) {
  gravityDataOutput[`${item.Element}__${item["Element Station"]}`] = item;
}
for (const item of lsaSheetData) {
  lsaDataOutput[`${item.Element}__${item["Element Station"]}`] = item;
}

function generateData(
  gravityData: NormalizedSheetItem,
  lsaData: NormalizedSheetItem
) {
  let dataKeyGravity = Object.keys(gravityData);
  let finalOutput: SheetItem[] = [];

  dataKeyGravity.forEach(rowKey => {
    const key = rowKey.split("__")[0];
    const gravityRow = gravityData[rowKey];
    const lsaRow = lsaData[rowKey];
    let rowData: SheetItemWithSources = {
      Story: gravityRow.Story,
      Beam: gravityRow.Beam,
      UniqueName: gravityRow["UniqueName"],
      "Output Case": "Combined",
      "Case Type": gravityRow["Case Type"],
      "Step Type": gravityRow["Step Type"],
      Station: gravityRow.Station,
      P: getHigherAbsoluteValue(gravityRow.P, lsaRow.P).value,
      "P source": getHigherAbsoluteValue(gravityRow.P, lsaRow.P).s,
      V2: getHigherAbsoluteValue(gravityRow.V2, lsaRow.V2).value,
      "V2 source": getHigherAbsoluteValue(gravityRow.V2, lsaRow.V2).s,
      V3: getHigherAbsoluteValue(gravityRow.V3, lsaRow.V3).value,
      "V3 source": getHigherAbsoluteValue(gravityRow.V3, lsaRow.V3).s,
      T: getHigherAbsoluteValue(gravityRow.T, lsaRow.T).value,
      "T source": getHigherAbsoluteValue(gravityRow.T, lsaRow.T).s,
      M2: getHigherAbsoluteValue(gravityRow.M2, lsaRow.M2).value,
      "M2 source": getHigherAbsoluteValue(gravityRow.M2, lsaRow.M2).s,
      M3: getHigherAbsoluteValue(gravityRow.M3, lsaRow.M3).value,
      "M3 source": getHigherAbsoluteValue(gravityRow.M3, lsaRow.M3).s,
      Element: gravityRow.Element,
      "Element Station": gravityRow["Element Station"],
      Location: gravityRow["Location"],
    };
    finalOutput.push(rowData);
  });
  return finalOutput;
}

let jsonOutput = generateData(gravityDataOutput, lsaDataOutput);

const newWb = xlsx.utils.book_new();
const newWS = xlsx.utils.json_to_sheet(jsonOutput);
xlsx.utils.book_append_sheet(newWb, newWS, "New data");

xlsx.writeFile(newWb, "output2.xlsm");
