"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const xlsx_1 = __importDefault(require("xlsx"));
const path_1 = __importDefault(require("path"));
const utils_1 = require("./utils");
const GRAVITY_SHEET = "Gravity";
const LSA_SHEET = "LSA";
const readline = require("readline");
(() => __awaiter(void 0, void 0, void 0, function* () {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    const fileName = yield new Promise(resolve => {
        rl.question(`Enter the target file name, eg: 'example.xlsm'.(Press enter for default file = 'BDA.xlsm')
      `, (name) => {
            resolve(name);
        });
    });
    const filePath = path_1.default.join(__dirname, fileName);
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
        gravityDataOutput[`${item.Element}__${item["Element Station"]}__${item["Step Type"]}`] = item;
    }
    for (const item of lsaSheetData) {
        lsaDataOutput[`${item.Element}__${item["Element Station"]}__${item["Step Type"]}`] = item;
    }
    let jsonOutput = (0, utils_1.generateData)(gravityDataOutput, lsaDataOutput);
    function generateFile() {
        const newWb = xlsx_1.default.utils.book_new();
        const newWS = xlsx_1.default.utils.json_to_sheet(jsonOutput);
        xlsx_1.default.utils.book_append_sheet(newWb, newWS, "New data");
        xlsx_1.default.writeFile(newWb, `output4-${utils_1.formattedDate}.xlsm`);
    }
    generateFile();
    rl.close();
}))();
