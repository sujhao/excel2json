
const xlsx = require('node-xlsx');
const fs = require("fs")
const path = require("path")
// 
// const workSheetsFromFile = xlsx.parse('excel/testa.xlsx');
const workSheetsFromFile = xlsx.parse('excel/Z.装备配置.xlsx');

console.log("workSheetsFromFile====", workSheetsFromFile);
let outputPath = path.resolve() + "/json/"
console.log("outputPath====", outputPath);

for (let i = 0; i < workSheetsFromFile.length; i++) {
    let sheetObj = workSheetsFromFile[i];
    let sheetName = sheetObj["name"];
    let sheetData = sheetObj["data"];
    let formatList = [];
    let keyList = [];
    let dataList = [];
    for (let j = 0; j < sheetData.length; j++) {
        let oneRowData = sheetData[j];
        if (j == 1) {
            formatList = oneRowData
        }
        else if (j == 2) {
            keyList = oneRowData
        } else if (j > 1) {
            let oneRowObj = {};
            for (let k = 0; k < oneRowData.length; k++) {
                let cellData = oneRowData[k];
                let key = keyList[k];
                let format = formatList[k].toLowerCase();
                // console.log("cellData=======", key, cellData);
                let value = cellData;
                if (format == "array" || format == "object") {
                    try {
                        value = JSON.parse(cellData);
                    } catch (error) {
                        // console.log("格式化错误=", key, format, cellData);
                    }
                }
                oneRowObj[key] = value;
                dataList.push(oneRowObj);
            }
        }
    }
    console.log("keyList=", keyList);
    console.log("dataList=", dataList);
    let jsonStr = JSON.stringify(dataList);
    // console.log("jsonStr=", jsonStr);
    let filePath = outputPath + sheetName + ".json"
    console.log("filePath=", filePath);
    fs.writeFile(filePath, jsonStr, function (err) {
        if (err) {
            console.warn("writeFile==error", filePath);
        }
    });
}
console.log("end");