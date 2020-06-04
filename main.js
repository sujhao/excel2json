
const xlsx = require('node-xlsx');
const fs = require("fs")
const path = require("path")
// 
// const workSheetsFromFile = xlsx.parse('excel/testa.xlsx');
// const workSheetsFromFile = xlsx.parse('excel/Z.装备配置.xlsx');

// console.log("workSheetsFromFile====", workSheetsFromFile);
let outputPath = path.resolve() + "/json/"
let inputPath = path.resolve() + "/excel"
console.log("outputPath====", outputPath);

function exportOneExcel(workSheetsFromFile) {
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
}

function searchDir(inputPath) {
    //根据文件路径读取文件，返回文件列表  
    fs.readdir(inputPath, (err, files) => {
        if (err) {
            console.warn(err);
        } else {
            //遍历读取到的文件列表  
            files.forEach((filename) => {
                //获取当前文件的绝对路径  
                let filedir = path.join(inputPath, filename);
                console.log("filedir==", filedir);
                //根据文件路径获取文件信息，返回一个fs.Stats对象  
                fs.stat(filedir, (eror, stats) => {
                    if (eror) {
                        console.warn('获取文件stats失败');
                    } else {
                        let isFile = stats.isFile(); //是文件  
                        let isDir = stats.isDirectory(); //是文件夹  
                        if (isFile) {
                            console.log("filedir=============2", filedir, isFile);
                            let workSheetsFromFile = xlsx.parse(filedir);
                            exportOneExcel(workSheetsFromFile)
                        }
                        if (isDir) {
                            searchDir(filedir); //递归，如果是文件夹，就继续遍历该文件夹下面的文件  
                        }
                    }
                })
            });
        }
    });
}

function main() {
    searchDir(inputPath)
}

main();

console.log("end");