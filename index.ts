import xlsx from "xlsx";
import path from "path";
import { NormalizedSheetItem, SheetItem, SheetItemWithSources } from "./type";
import { formattedDate, generateData, getHigherAbsoluteValue } from "./utils";
const GRAVITY_SHEET = "Gravity";
const LSA_SHEET = "LSA";

const readline = require("readline");

(async () => {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  const fileName = await new Promise<string>(resolve => {
    rl.question(
      `Enter the target file name, eg: 'example.xlsm'.(Press enter for default file = 'BDA.xlsm')
      `,
      (name: string) => {
        resolve(name || "BDA.xlsm");
      }
    );
  });
  const filePath = path.join(`${__dirname}/InputFolder`, fileName);

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
    gravityDataOutput[
      `${item.Element}__${item["Element Station"]}__${item["Step Type"]}`
    ] = item;
  }
  for (const item of lsaSheetData) {
    lsaDataOutput[
      `${item.Element}__${item["Element Station"]}__${item["Step Type"]}`
    ] = item;
  }
  let jsonOutput = generateData(gravityDataOutput, lsaDataOutput);
  function generateFile() {
    const newWb = xlsx.utils.book_new();
    const newWS = xlsx.utils.json_to_sheet(jsonOutput);
    xlsx.utils.book_append_sheet(newWb, newWS, "New data");
    xlsx.writeFile(newWb, `./Output/output-${formattedDate}.xlsm`);
  }

  generateFile();
  rl.close();
})();
