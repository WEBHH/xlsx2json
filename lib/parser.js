const config = require('../config.json');
const types = require('./types');
var _ = require('lodash');

const DataType = types.DataType;
const SheetType = types.SheetType;
const JsonType = types.JsonType;

// class StringBuffer {
//   constructor(str) {
//     this._str_ = [];
//     if (str) {
//       this.append(str);
//     }
//   }

//   toString() {
//     return this._str_.join("");
//   }

//   append(str) {
//     this._str_.push(str);
//   }
// }


/**
 * 解析workbook中所有sheet的设置
 * @param {*} workbook 
 */
function parseSettings(workbook) {

  /**
   * settings's schema
   * {
      type: SheetType.NORMAL,
      jsonType: JsonType.ARRAY,
      head: [{
          name: "json-key(column name)", //  字段名称
          type: "int/string/bool/[]/{}", // 字段类型
          isKey: isKey, // 当数据类型为hash时，该字段是否为key
        }]
      ,
      slaves: [],
      master: {
        name: '',
        type: DataType.OBJECT
      },
    } 
   */

  let settings = {};

  workbook.forEach(sheet => {

    //叹号开头的sheet不输出
    if (sheet.name.startsWith('!')) {
      return;
    }

    let sheet_name = sheet.name;
    let sheet_setting = {
      type: SheetType.NORMAL,
      jsonType: JsonType.ARRAY,
      master: null,
      slaves: [],
      head: []
    };

    // 是否主表
    if (sheet_name.indexOf('#master') >= 0) {
      sheet_name = sheet_name.split('#')[0];
      sheet_setting.type = SheetType.MASTER;
    }

    // 是否从表
    if (sheet_name.indexOf('@') >= 0) {
      sheet_setting.type = SheetType.SLAVE;
      let pair = sheet_name.split('@');
      sheet_name = pair[0].trim();
      sheet_setting.master = pair[1].trim();
      settings[sheet_setting.master].slaves.push(sheet_name);
    }

    // 解析字段类型行 & 字段名称行 --2020-09-30 vicky
    let type_row = sheet.data[config.xlsx.head - 2];
    let head_row = sheet.data[config.xlsx.head - 1];

    type_row.forEach((cell, index) => {
      let head = head_row[index];
      let type = cell.toString() || DataType.UNKNOWN; // 字段类型
      let name = ""; // 字段名称
      let isKey = false; // hash模式的key

      // 是否该表为hash格式 & 该字段为key
      if (head.indexOf('#') >= 0) {
        name = head.toString().substr(1);
        if (sheet_setting.jsonType !== JsonType.HASH) {
          sheet_setting.jsonType = JsonType.HASH;
          isKey = true;
        }
      } else {
        name = head.toString();
      }

      let head_setting = {
        name: name,
        type: type,
        isKey: isKey,
      };

      sheet_setting.head.push(head_setting);
    });

    settings[sheet_name] = sheet_setting;
  });

  console.log("sheet的设置: ", settings);

  return settings;
}


/**
 * 解析一个表(sheet)
 *
 * @param sheet 表的原始数据
 * @param setting 表的设置
 * @return Array or Object
 */
function parseSheet(sheet, setting) {

  let result = [];

  if (setting && setting.jsonType === JsonType.HASH) {
    result = {};
  }

  // console.log('  * sheet:', sheet.name, 'rows:', sheet.data.length, 'setting:', setting);

  // 从表头行开始解析
  for (let i_row = config.xlsx.head; i_row < sheet.data.length; i_row++) {

    let row = sheet.data[i_row];

    let parsed_row = parseRow(row, i_row, setting.head);

    if (setting.jsonType === JsonType.HASH) {

      let id_cell = _.find(setting.head, item => {
        return item.isKey;
      });

      if (!id_cell) {
        throw `在表${sheet.name}中获取不到id列`;
      }

      result[parsed_row[id_cell.name]] = parsed_row;

    } else {
      result.push(parsed_row);
    }
  }

  return result;
}

/**
 * 解析一行
 * @param {*} row 
 * @param {*} rowIndex 
 * @param {*} head 
 */
function parseRow(row, rowIndex, head) {

  let result = {};
  // let id;

  for (let index = 0; index < head.length; index++) {
    let cell = row[index];

    let name = head[index].name;
    let type = head[index].type;

    if (name.startsWith('!')) {
      continue;
    }

    if (cell === null || cell === undefined) {
      result[name] = getDefault(type);
      continue;
    }

    // console.log("=========");
    // console.log("解析行: ", row[index]);
    // console.log("解析行对应头: ", head[index]);

    switch (type) {
      case DataType.UNKNOWN: // int string boolean
        if (isNumber(cell)) {
          result[name] = Number(cell);
        } else if (isBoolean(cell)) {
          result[name] = toBoolean(cell);
        } else {
          result[name] = cell;
        }
        break;
      case DataType.DATE:
        if (isNumber(cell)) {
          //xlsx's bug!!!
          result[name] = numdate(cell);
        } else {
          result[name] = cell.toString();
        }
        break;
      case DataType.STRING:
        result[name] = cell.toString();
        break;
      case DataType.INT:
      case DataType.LONG:
      case DataType.FLOAT:
        //+xxx.toString() '+' means convert it to number
        if (isNumber(cell)) {
          result[name] = Number(cell);
        } else {
          console.warn("type error at [" + rowIndex + "," + index + "]," + cell + " is not a number");
        }
        break;
      case DataType.BOOL:
        result[name] = toBoolean(cell);
        break;
      case DataType.OBJECT:
        result[name] = JSON.parse(cell);
        break;
      case DataType.ARRAY:
      case DataType.ARRAY_INT:
      case DataType.ARRAY_STRING:
      case DataType.ARRAY_FLOAT:
        if (!cell.toString().startsWith('[')) {
          cell = `[${cell}]`;
        }
        result[name] = JSON.parse(cell);
        break;
      default:
        console.log('无法识别的类型:', '[' + rowIndex + ',' + index + ']', cell, typeof (cell));
        break;
    }
  }

  return result;
}

/**
 * convert value to boolean.
 */
function toBoolean(value) {
  return value.toString().toLowerCase() === 'true';
}

/**
 * is a number.
 */
function isNumber(value) {

  if (typeof value === 'number') {
    return true;
  }

  if (value) {
    return !isNaN(+value.toString());
  }

  return false;
}

/**
 * boolean type check.
 */
function isBoolean(value) {

  if (typeof (value) === "undefined") {
    return false;
  }

  if (typeof value === 'boolean') {
    return true;
  }

  let b = value.toString().trim().toLowerCase();

  return b === 'true' || b === 'false';
}

//fuck node-xlsx's bug
var basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
// var dnthresh = basedate.getTime() + (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
var dnthresh = basedate.getTime() + (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
// function datenum(v, date1904) {
// 	var epoch = v.getTime();
// 	if(date1904) epoch -= 1462*24*60*60*1000;
// 	return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
// }

function numdate(v) {
  var out = new Date();
  out.setTime(v * 24 * 60 * 60 * 1000 + dnthresh);
  return out;
}
//fuck over

// 获取字段类型的默认值
function getDefault(type) {
  let value = null;
  switch (type) {
    case DataType.STRING:
      value = "";
      break;
    case DataType.INT:
    case DataType.LONG:
    case DataType.FLOAT:
      value = 0;
      break;
    case DataType.ARRAY:
    case DataType.ARRAY_INT:
    case DataType.ARRAY_STRING:
    case DataType.ARRAY_FLOAT:
      value = [];
      break;
  }

  return value;
}


module.exports = {

  parseSettings: parseSettings,

  parseWorkbook: function (workbook, settings) {

    // console.log('settings >>>>>', JSON.stringify(settings, null, 2));
    // console.log("parseWorkboo==workbook== ", workbook);
    // console.log("parseWorkboo==settings== ", settings);

    let parsed_workbook = {};

    workbook.forEach(sheet => {

      if (sheet.name.startsWith('!')) {
        return;
      }

      let sheet_name = sheet.name;

      // 若为从表
      if (sheet_name.indexOf('@') >= 0) {
        sheet_name = sheet_name.split('@')[0].trim();
      }

      // 若为主表
      if (sheet_name.indexOf('#') >= 0) {
        sheet_name = sheet_name.split('#')[0].trim();
      }

      let sheet_setting = settings[sheet_name];
      let parsed_sheet = parseSheet(sheet, sheet_setting);

      parsed_workbook[sheet_name] = parsed_sheet;

    });

    for (let name in settings) {
      // 若为主表
      if (settings[name].type === SheetType.MASTER) {
        // 取主表数据
        let master_sheet = parsed_workbook[name];

        // 遍历从表
        settings[name].slaves.forEach(slave_name => {
          let slave_setting = settings[slave_name]; // 从表配置
          let slave_sheet = parsed_workbook[slave_name]; // 从表数据

          // console.log("slave 表中所有数据====", slave_sheet);

          // 遍历从表数据
          if (slave_setting.jsonType === JsonType.ARRAY) {
            slave_sheet.forEach(row => {
              let id = row["master"];
              delete row["master"];
              master_sheet[id][slave_name] = master_sheet[id][slave_name] || [];
              master_sheet[id][slave_name].push(row);
            });
          } else if (slave_setting.jsonType === JsonType.HASH) {
            for (let key in slave_sheet) {
              delete slave_sheet[key]["master"];
              master_sheet[key][slave_name] = slave_sheet[key];
            }
          }

          delete parsed_workbook[slave_name];
        });
      }
    }

    return parsed_workbook;
  }
};