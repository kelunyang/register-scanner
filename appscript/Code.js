/*
2. MODE： debug, 寫過不能再寫, 直出, 根據狀態決定能不能再寫
*/
const _ = LodashGS.load();
const appProperties = PropertiesService.getScriptProperties();

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index')
  var html = template.evaluate()
    .setTitle(appProperties.getProperty('systemTitle'));
  
  var htmlOutput = HtmlService.createHtmlOutput(html);
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}
function buildList() {
  let sheetListID = appProperties.getProperty("sheetListID");
  let spreadsheet = SpreadsheetApp.openById(sheetListID);
  let sheet = spreadsheet.getSheetByName("條碼列表");
  let dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7);
  let dataValues = dataRange.getValues();

  let dataList = [];
  dataValues.forEach(function(row) {
    let item = {
      title: row[1],
      id: row[0],
      dueDate: row[2],
      duplicateWord: row[4],
      failWord: row[5],
      successWord: row[6],
      mode: row[3]
    };
    dataList.push(item);
  });
  return dataList;
}

function checkAuth(userID, password, sheetID) {
  const sheetListID = appProperties.getProperty("sheetListID");
  const spreadsheet = SpreadsheetApp.openById(sheetListID);
  const sheet = spreadsheet.getSheetByName("帳號密碼表");
  const dataRange = sheet.getDataRange();
  const dataValues = dataRange.getValues();

  const userMatch = _.find(dataValues, function(row) {
    const storedUserID = row[0];
    const storedSheetID = row[2];
    const storedPassword = row[1];
    if(storedUserID === userID) {
      if(storedSheetID === sheetID) {
        return storedPassword === password;
      }
    }
    return false;
  });

  if (userMatch) {
    let barcodeSheet = spreadsheet.getSheetByName('條碼列表');
    
    // 使用sheetID去過濾A欄，找到符合的列
    let barcodeData = barcodeSheet.getRange(2, 1, barcodeSheet.getLastRow() - 1, barcodeSheet.getLastColumn()).getValues();
    let matchedBarcodeRow = _.find(barcodeData, function(row) {
      return row[0] === sheetID;
    });
    if(!matchedBarcodeRow) {
      return false;
    } else {
      // 組裝成物件
      let sheetObj = {
        id: sheetID,
        timestamp: matchedBarcodeRow[2], // C欄
        mode: matchedBarcodeRow[3],      // D欄
        duplicateWord: matchedBarcodeRow[4], // E欄
        failWord: matchedBarcodeRow[5],      // F欄
        successWord: matchedBarcodeRow[6]    // G欄
      };
      let currentQuery = { inputColumns: [], viewableColumns: [], debugData: [], mode: matchedBarcodeRow[3] };

      // 開啟試算表，試算表ID等於sheetID
      const spreadsheetToOpen = SpreadsheetApp.openById(sheetID);
      
      // 讀取「程式寫入表」，讀取C欄開始的第一列到第三列，用物件輸出
      const configSheet = spreadsheetToOpen.getSheetByName("程式寫入表");
      const configValues = configSheet.getRange(1, 3, 4, configSheet.getLastColumn() - 2).getValues();
      const inputColumns = [];
      for (let i = 0; i < configValues[0].length; i++) {
        const column = {
          id: "column" + i,
          name: configValues[0][i],
          type: configValues[1][i],
          format: configValues[2][i],
          must: /M/.test(configValues[3][i]),
          value: "",
          status: "",
          state: /S/.test(configValues[3][i])
        };
        if(column.type === "P") { column.must = true; }
        inputColumns.push(column);
      }
      
      // Sort inputColumns: typeP should be first, then order from configValues
      const typePColumn = _.find(inputColumns, { type: "P" });
      if(typePColumn) {
        const remainingColumns = _.filter(inputColumns, (column) => {
          return column.type !== "P";
        });
        const sortedInputColumns = _.sortBy(remainingColumns, (column) => {
          return _.indexOf(configValues[0], column.name);
        });

        if (typePColumn) {
          sortedInputColumns.unshift(typePColumn);
        }

        currentQuery.inputColumns = sortedInputColumns;

        // 讀取「原始資料表」的第一列，逐欄把值放到currentQuery去
        const rawDataSheet = spreadsheetToOpen.getSheetByName("原始資料表");
        const rawData = rawDataSheet.getRange(1, 1, rawDataSheet.getLastRow(), rawDataSheet.getLastColumn()).getValues();
        let rawHeader = rawData.shift();
        const viewableColumns = rawHeader.map(function(header, index) {
          return { key: index, label: header };
        });
        currentQuery.viewableColumns = viewableColumns;
        if(/測試/.test(sheetObj.mode)) {
          const rawDB = rawData.map(function(row) {
            let dataObj = {};
            viewableColumns.forEach(function(viewableColumn) {
              dataObj[viewableColumn.label] = row[viewableColumn.key];
            });
            return dataObj;
          });
          currentQuery.debugData = rawDB;
        }
        return {
          status: true,
          result: currentQuery
        }
      } else {
        return {
          status: false,
          msg: "這份表單沒有設定主鍵，請通知管理員"
        };
      }
    }
  } else {
    return {
      status: false,
      msg: "帳號密碼錯誤"
    };
  }
}

function writeRecord(account, password, sheetID, inputColumns, selectedViewables) {
  // 開啟主控台試算表
  let sheetListID = PropertiesService.getScriptProperties().getProperty('sheetListID');
  let spreadsheet = SpreadsheetApp.openById(sheetListID);
  
  // 開啟主控台試算表的「帳號密碼表」
  let accountSheet = spreadsheet.getSheetByName('帳號密碼表');
  let data = accountSheet.getDataRange().getValues();
  
  // 比對account, password和sheetID
  let matchedRow = _.find(data, function(row) {
    return row[0] === account && row[1] === password && row[2] === sheetID;
  });
  
  if (matchedRow) {
    let barcodeSheet = spreadsheet.getSheetByName('條碼列表');
    
    // 使用sheetID去過濾A欄，找到符合的列
    let barcodeData = barcodeSheet.getRange(2, 1, barcodeSheet.getLastRow() - 1, barcodeSheet.getLastColumn()).getValues();
    let matchedBarcodeRow = _.find(barcodeData, function(row) {
      return row[0] === sheetID;
    });
    if(!matchedBarcodeRow) {
      return {
        status: false,
        result: "你是怎麼做到存取不存在的表單的？"
      };
    } else {
      // 組裝成物件
      let sheetObj = {
        id: sheetID,
        timestamp: matchedBarcodeRow[2], // C欄
        mode: matchedBarcodeRow[3],      // D欄
        duplicateWord: matchedBarcodeRow[4], // E欄
        failWord: matchedBarcodeRow[5],      // F欄
        successWord: matchedBarcodeRow[6]    // G欄
      };
      //把送進來的變數全部toString
      for(let i=0; i<inputColumns.length; i++) {
        inputColumns[i].value = _.toString(inputColumns[i].value);
      }
      // 通過比對後，開啟指定的試算表
      let targetSpreadsheet = SpreadsheetApp.openById(sheetID);
      let nameRow = [];
      let rawRow = false;
      // 掃描inputColumns找出type是P的物件
      let typePColumn = _.find(inputColumns, { 'type': 'P' });
      if (typePColumn) {
        // 開啟試算表的「原始資料表」
        let rawDataSheet = targetSpreadsheet.getSheetByName('原始資料表');
        let rawData = rawDataSheet.getDataRange().getValues();
        
        nameRow = rawData[0];
        let nameIndex = nameRow.indexOf(typePColumn.name);
        if(nameIndex === -1) {
          return {
            status: false,
            result: "輸入表和原始表的主鍵名稱不一致，請聯絡管理員重新設定"
          };
        }
        rawRow = _.find(rawData, function(row) {
          return _.toString(row[nameIndex]) === typePColumn.value;
        });
        
        // 檢查typePColumn的value是否在typePValues中
        if (!rawRow) {
          return {
            status: false,
            result: sheetObj.failWord
          };
        }
      } else {
        return {
          status: false,
          result: "你送入的欄位沒有主鍵?!?!?!?!"
        }
      }
      
      // 開啟「程式寫入表」，讀取C欄開始的各欄，從第一列讀到第三列，並組成 savedInput 陣列
      let programSheet = targetSpreadsheet.getSheetByName('程式寫入表');
      let savedInput = programSheet.getRange(1, 3, programSheet.getLastRow(), programSheet.getLastColumn() - 2).getValues();
      
      // for inputColumn 進行匹配
      for (let i = 0; i < inputColumns.length; i++) {
        let column = inputColumns[i];
        let savedColumn = _.indexOf(savedInput[0], column.name);
        
        if (savedColumn > -1) {
          // 執行 valField，如果不是 true 則回傳錯誤訊息
          let validationResult = valField(column, [savedInput[0][savedColumn],savedInput[1][savedColumn],savedInput[2][savedColumn],savedInput[3][savedColumn]], sheetObj.mode);
          if (!validationResult.status) {
            return {
              status: false,
              result: "格式檢查失敗，原因：" + validationResult.result
            };
          } else {
            column = validationResult.result
          }
        } else {
          // 找不到對應的 savedInput，回傳 false
          return {
            status: false,
            result: "你是怎麼做到傳入寫入表裡沒有的欄位的？"
          };
        }
      }
      // 檢查typePColumn的value是否在「程式寫入表」中出現過（報到機）
      if(/報到/.test(sheetObj.mode)) {
        let typePValue = typePColumn.value;
        let typePColumnIndex = _.indexOf(savedInput[0], typePColumn.name);
        if (typePColumnIndex !== -1) {
          let typePColumnValues = _.map(savedInput, function(row) {
            return _.toString(row[typePColumnIndex]);
          });
          if (_.includes(typePColumnValues, typePValue)) {
            return {
              status: false,
              result: sheetObj.duplicateWord
            };
          }
        }
      }
      if(/借用/.test(sheetObj.mode)) {
        let typePColumnIndex = _.indexOf(savedInput[0], typePColumn.name);
        // 找出InputColumn物件屬性state是true的column物件，陣列名稱為stateColumns
        let stateColumns = _.filter(inputColumns, { state: true });
        let matchedSavedInputRows = _.filter(savedInput, function(row) {
          return _.toString(row[typePColumnIndex]) === typePColumn.value;
        });
        
        // 如果找到符合條件的列
        if (matchedSavedInputRows.length > 0) {
          // 取出符合條件的所有列中的最後一列
          let lastRowValues = matchedSavedInputRows[matchedSavedInputRows.length - 1];
        
          // 比對stateColumn的index是否和目前inputColumns中state是true的value相符
          let isEveryStateMatch = _.every(stateColumns, function(column) {
            let index = _.indexOf(savedInput[0], column.name);
            return lastRowValues[index] === column.value;
          });
        
          // 完全相符(every)則回傳fasle
          if (isEveryStateMatch) {
            return {
              status: false,
              result: sheetObj.duplicateWord
            };
          }
        }
      }
      
      // 寫入「程式寫入表」
      let newRow = [new Date().getTime(), account];
      let sortedInputColumns = _.sortBy(inputColumns, (column) => {
        return _.indexOf(savedInput[0], column.name);
      });
      for(let i=0; i<sortedInputColumns.length; i++) {
        newRow.push(sortedInputColumns[i].value);
      }
      programSheet.appendRow(newRow);
      // 尋找符合 selectedViewables 的值，並放入 returnResult 陣列
      let returnResult = [];
      for(let i=0; i<selectedViewables.length; i++) {
        if(rawRow.length > selectedViewables[i]) {
          returnResult.push(rawRow[selectedViewables[i]]);
        }
      }

      return {
        status: true,
        result: returnResult
      }
    }
  } else {
    return {
      status: false,
      result: "帳號密碼驗證失敗，請聯絡管理員"
    };
  }
}

function valField(column, savedColumn, sheetMode) {
  column.type = savedColumn[1];
  column.format = savedColumn[2];
  column.must = /M/.test(savedColumn[3]);
  if(formatDetector('P', column)) { column.must = true }
  column.status = "";
  if(column.must) { //先檢查是否為空
    if(column.value === "") {
      return {
        status: false,
        result: column.name + "必需有值！"
      }
    }
  }
  let formatCheck = column.must;
  if(!column.must) {
    formatCheck = column.value !== "";
  }
  if(/借用/.test(sheetMode)) {
    formatCheck = /S/.test(savedColumn[3]);
  }

  if(formatCheck) {
    if(!formatDetector('P', column)) {
      if(formatDetector('T', column)) {
        if(column.format !== "") {
          let regexConfig = column.format.split("::");
          if(new RegExp(regexConfig[1]).test(column.value)) {
            column.value = column.value.replace(/台(北|中|南|灣)/,'臺$1');
          } else {
            return {
              status: false,
              result: column.name + "格式提示為「" + regexConfig[0] + "」"
            };
          }
        }
      } else if(formatDetector('S', column)) {
        let selections = column.format.split(";");
        if(selections.length > 0) {
          if(!_.includes(selections, column.value)) {
            return {
              status: false,
              result: column.name + "欄位的值你真的是用選單選出來的？（還是你沒選？）"
            }
          }
        } else {
          return {
            status: false,
            result: column.name + "這個欄位沒設定選項"
          }
        }
      }
    }
  }
  return {
    status: true,
    result: column
  };
}

function formatDetector(formatReg, column) {
  if((new RegExp(formatReg)).test(column.type)) {
    return true;
  }
  return false;
}