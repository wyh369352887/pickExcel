/*
 * @Author: elephant.H
 * @Date:   2018-01-03 17:20:03
 * @Last Modified by:   elephant.H
 * @Last Modified time: 2018-01-04 17:53:59
 */
var X = XLSX;
var XW = {
    /* worker message */
    msg: 'xlsx',
    /* worker scripts */
    rABS: './xlsxworker2.js',
    norABS: './xlsxworker1.js',
    noxfer: './xlsxworker.js'
};

var rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";
if (!rABS) {
    document.getElementsByName("userabs")[0].disabled = true;
    document.getElementsByName("userabs")[0].checked = false;
}

var use_worker = typeof Worker !== 'undefined';
if (!use_worker) {
    document.getElementsByName("useworker")[0].disabled = true;
    document.getElementsByName("useworker")[0].checked = false;
}

var transferable = use_worker;
if (!transferable) {
    document.getElementsByName("xferable")[0].disabled = true;
    document.getElementsByName("xferable")[0].checked = false;
}

var wtf_mode = false;

function fixdata(data) {
    var o = "",
        l = 0,
        w = 10240;
    for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
    o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
    return o;
}

function ab2str(data) {
    var o = "",
        l = 0,
        w = 10240;
    for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint16Array(data.slice(l * w, l * w + w)));
    o += String.fromCharCode.apply(null, new Uint16Array(data.slice(l * w))); //打印这块的o的时候也出现了文字
    return o;
}

function s2ab(s) {
    var b = new ArrayBuffer(s.length * 2),
        v = new Uint16Array(b);
    //console.log(b); 此时b为空
    for (var i = 0; i != s.length; ++i) v[i] = s.charCodeAt(i);
    return [v, b];
}

function xw_noxfer(data, cb) {
    var worker = new Worker(XW.noxfer);
    worker.onmessage = function(e) {
        switch (e.data.t) {
            case 'ready':
                break;
            case 'e':
                console.error(e.data.d);
                break;
            case XW.msg:
                cb(JSON.parse(e.data.d));
                break;
        }
    };
    var arr = rABS ? data : btoa(fixdata(data));
    worker.postMessage({ d: arr, b: rABS });
}

function xw_xfer(data, cb) {
    var worker = new Worker(rABS ? XW.rABS : XW.norABS);
    worker.onmessage = function(e) {
        switch (e.data.t) {
            case 'ready':
                break;
            case 'e':
                console.error(e.data.d);
                break;
            default:
                xx = ab2str(e.data).replace(/\n/g, "\\n").replace(/\r/g, "\\r");
                console.log("done");
                cb(JSON.parse(xx));
                break; //从这一块打印出来的xx出现的文字
        }
    };
    if (rABS) {
        var val = s2ab(data);
        worker.postMessage(val[1], [val[1]]);
    } else {
        worker.postMessage(data, [data]);
    } //我运行到这一步了
    //console.log(data);  已经接近了
}

function xw(data, cb) {
    if (transferable) xw_xfer(data, cb);
    else xw_noxfer(data, cb);
}

function get_radio_value(radioName) {
    var radios = document.getElementsByName(radioName);
    for (var i = 0; i < radios.length; i++) {
        if (radios[i].checked || radios.length === 1) {
            //console.log(radios);
            return radios[i].value;
        }
    }
}
//这个是必须的
function to_json(workbook) {
    var result = {};
    workbook.SheetNames.forEach(function(sheetName) {
        var roa = X.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        if (roa.length > 0) {
            result[sheetName] = roa;
        }
    });
    return result;
}

function process_wb(wb) {
    var output = "";
    switch (get_radio_value("format")) {
        case "json":
            outputJson = JSON.stringify(to_json(wb), 2, 2);
            //字符串格式的数据
            output = to_json(wb);
            //json格式的数据
            // console.log(output);
            //这边是最终输出
            break;

    }

    // if (out.innerHTML === undefined) out.textContent = output;
    // else out.innerHTML = `<textarea>${outputJson}</textarea>`;
    // if (typeof console !== 'undefined') console.log("output", new Date());
    var excelPage = document.getElementById('excelPageNum').value;
    if (excelPage != '') {
        switch (target) {
            case 'getHoleExcelColumn':
            	var inputs = document.getElementById('innerOption').getElementsByTagName('input'),
            		columnName = [];
            	for(var i = 0; i < inputs.length; i++){
            		columnName.push(inputs[i].value);
            	}
                getHoleColumn(excelPage, columnName, output);
            break;
        }
    } else {
        alert('请先配置选项!');
    }
    //     columnName = '客户需提供的基本手续,社保记录,贷款期限,信用记录',
    //     lineSection = '5-100',
    //     lineStart = 2,//excel数据实际从第几行开始
    //     lineData = {
    //     	name:'仓储面积',
    //     	data:'2000'
    //     };
    // getHoleColumn(excelPage,columnName,output);
    // getSeriesLine(excelPage, lineSection, output,lineStart);
    // getLineByData(excelPage,lineData,output);

}

function getLineByData(excelPage, lineData, output) {
    var resData = {};
    var len = output[excelPage].length;
    for (var i = 0, n = 0; i < len; i++) {
        if (output[excelPage][i][lineData.name] == lineData.data) {
            resData[n] = output[excelPage][i];
            n++;
        }
    }
    resData = JSON.stringify(resData, 2, 2);
    if (out.innerHTML === undefined) out.textContent = resData;
    else out.innerHTML = `<textarea>${resData}</textarea>`;
    if (typeof console !== 'undefined') console.log("output", new Date());
}
//抓取符合条件的整列
function getHoleColumn(excelPage, columnName, output) {
    var cnArr = columnName;
    console.log(cnArr);
    var resColumn = {};
    for (var i = 0; i < output[excelPage].length; i++) {
        var resObj = {};
        for (var j = 0; j < cnArr.length; j++) {
            if (output[excelPage][i][cnArr[j]]) {
                resObj[cnArr[j]] = output[excelPage][i][cnArr[j]];
            } else {
                resObj[cnArr[j]] = '';
            }
        }
        resColumn[i] = resObj;
    }
    resColumn = JSON.stringify(resColumn, 2, 2);
    if (out.innerHTML === undefined) out.textContent = resColumn;
    else out.innerHTML = `<textarea>${resColumn}</textarea>`;
    if (typeof console !== 'undefined') console.log("output", new Date());
}
//抓取整列数据
function getSeriesLine(excelPage, lineSection, output, lineStart) {
    console.log(output);
    var lsArr = lineSection.split('-');
    var section = lsArr[1] - lsArr[0];
    var resLine = {};
    for (var i = lsArr[0] - lineStart, n = 0; n < section + 1; n++) {
        resLine[n] = output[excelPage][i];
        i++;
    }
    resLine = JSON.stringify(resLine, 2, 2);
    if (out.innerHTML === undefined) out.textContent = resLine;
    else out.innerHTML = `<textarea>${resLine}</textarea>`;
    if (typeof console !== 'undefined') console.log("output", new Date());
}
//抓取整行数据
var drop = document.getElementById('drop');

function handleDrop(e) {
    e.stopPropagation();
    e.preventDefault();
    var files = e.dataTransfer.files;
    var f = files[0]; {
        var reader = new FileReader();
        var name = f.name;
        reader.onload = function(e) {
            var data = e.target.result;
            var abcd = JSON.stringify(data);
            //console.log(abcd);
            if (use_worker) {
                xw(data, process_wb);
            } else {
                var wb;
                if (rABS) {
                    wb = X.read(data, { type: 'binary' });
                } else {
                    var arr = fixdata(data);
                    wb = X.read(btoa(arr), { type: 'base64' });
                }
                process_wb(wb);
            }
        };
        if (rABS) reader.readAsBinaryString(f);
        else reader.readAsArrayBuffer(f);
    }
}

function handleDragover(e) {
    e.stopPropagation();
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
}

if (drop.addEventListener) {
    drop.addEventListener('dragenter', handleDragover, false);
    drop.addEventListener('dragover', handleDragover, false);
    drop.addEventListener('drop', handleDrop, false);
}
var target = '',
    holeColumn = document.getElementById('holeColumn'),
    holeLine = document.getElementById('holeLine'),
    dataLine = document.getElementById('dataLine');
if (holeColumn.addEventListener && holeLine.addEventListener && dataLine.addEventListener) {
    holeColumn.addEventListener('click', han, false);
    holeLine.addEventListener('click', han, false);
    dataLine.addEventListener('click', han, false);
}

function han() {
    var id = this.id,
        optionObj = document.getElementById('option');
    switch (id) {
        case 'holeColumn':
            optionObj.innerHTML = '<input type="text" id="columnNum" placeholder="请输入提取的总列数"><div id="innerOption"></div>';
            var columnNum = document.getElementById('columnNum');
            columnNum.addEventListener('input', forHoleColumnInput, false);
            target = 'getHoleExcelColumn';
            break;
        case 'holeLine':
            console.log(2);
            break;
        case 'dataLine':
            console.log(3);
            break;
        default:
            alert('error');
            break;
    }
}

function forHoleColumnInput() {
    var num = document.getElementById('columnNum').value,
    	text = '<input type="text" class="column-name" placeholder="列的名称">',
    	res = '';
    for(var i = 0; i < num; i++){
    	res += text;
    }
    document.getElementById('innerOption').innerHTML = res;
}