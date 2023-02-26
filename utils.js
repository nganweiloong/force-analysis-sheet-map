"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.generateData = exports.formattedDate = exports.getHigherAbsoluteValue = void 0;
function getHigherAbsoluteValue(a, b) {
    if (Math.abs(a) > Math.abs(b)) {
        return { value: a, s: "Gravity" };
    }
    else {
        return { value: b, s: "LSA" };
    }
}
exports.getHigherAbsoluteValue = getHigherAbsoluteValue;
exports.formattedDate = new Date()
    .toLocaleDateString("en-US", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
})
    .replace(/\//g, ""); // remove slashes
function generateData(gravityData, lsaData) {
    // console.log({ gravityData, lsaData });
    let dataKeyGravity = Object.keys(gravityData);
    let dataKeyLSA = Object.keys(lsaData);
    let finalOutput = [];
    dataKeyLSA.forEach(dataRowKeyLSA => {
        const lsaRow = lsaData[dataRowKeyLSA];
        const gravityRow = gravityData[dataRowKeyLSA];
        if (!gravityRow) {
            let rowData = {
                Story: lsaRow.Story,
                Beam: lsaRow.Beam,
                UniqueName: lsaRow["UniqueName"],
                "Output Case": "Combined",
                "Case Type": lsaRow["Case Type"],
                "Step Type": lsaRow["Step Type"],
                Station: lsaRow.Station,
                P: lsaRow.P,
                "P source": "LSA_",
                V2: lsaRow.V2,
                "V2 source": "LSA_",
                V3: lsaRow.V3,
                "V3 source": "LSA_",
                T: lsaRow.T,
                "T source": "LSA_",
                M2: lsaRow.M2,
                "M2 source": "LSA_",
                M3: lsaRow.M3,
                "M3 source": "LSA_",
                Element: lsaRow.Element,
                "Element Station": lsaRow["Element Station"],
                Location: lsaRow["Location"],
                "Row source": "LSA Standalone",
            };
            finalOutput.push(rowData);
        }
    });
    dataKeyGravity.forEach(dataRowKeyGravity => {
        const gravityRow = gravityData[dataRowKeyGravity];
        const lsaRow = lsaData[dataRowKeyGravity];
        if (lsaRow) {
            //comparable!
            console.count("writting row ");
            let rowData = {
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
                "Row source": "Compared",
            };
            finalOutput.push(rowData);
        }
        else {
            let rowData = {
                Story: gravityRow.Story,
                Beam: gravityRow.Beam,
                UniqueName: gravityRow["UniqueName"],
                "Output Case": "Combined",
                "Case Type": gravityRow["Case Type"],
                "Step Type": gravityRow["Step Type"],
                Station: gravityRow.Station,
                P: gravityRow.P,
                "P source": "Gravity_",
                V2: gravityRow.V2,
                "V2 source": "Gravity_",
                V3: gravityRow.V3,
                "V3 source": "Gravity_",
                T: gravityRow.T,
                "T source": "Gravity_",
                M2: gravityRow.M2,
                "M2 source": "Gravity_",
                M3: gravityRow.M3,
                "M3 source": "Gravity_",
                Element: gravityRow.Element,
                "Element Station": gravityRow["Element Station"],
                Location: gravityRow["Location"],
                "Row source": "GravityStandalone",
            };
            finalOutput.push(rowData);
        }
    });
    console.log(`
    /* -------------------------------------------------------------------------- */
    /*                                  DONE! :D                                  */
    /* -------------------------------------------------------------------------- */
    `);
    return finalOutput;
}
exports.generateData = generateData;
