/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
var __webpack_exports__ = {};
/*!******************************!*\
  !*** ./src/contentScript.js ***!
  \******************************/


const absenceList = ["廣川創太", "黒田賢人","税理士"];

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {

  if (request.message == "get YM") {
    let wndlct = window.location;
    let url = wndlct.host + wndlct.port + wndlct.pathname;
    if (url !== "dev.mirte.jp/attendance/edit" && url !== "mirte.jp/attendance/edit" && url !== "dev.mirte.jp/attendance/edit/update" && url !== "mirte.jp/attendance/edit/update") {
      sendResponse("no mirte");
    } else {
      let YM = document.getElementsByName("txtYM")[0].value;
      sendResponse(YM);
    }
  } else if (request.message == "sheet not found") {
    sendResponse("sheet not found");
  } else if (request.message == "version failed") {
    sendResponse("version failed");
  } else {

    let cannotFindOnMirteList = [];
    let actuallyAbsenceList = [];
    let StaffsOnMirteList = document.getElementsByClassName("table-responsive")[0].children[1];
    let excelDaySize;
    let mirteDaySize;

    //request.messageには、popup.jsから送信されたdatalist[]が送られてくる。
    for (let i = 0; i < StaffsOnMirteList.children.length; i += 2) {
      let noData = true;
      let staffData = StaffsOnMirteList.children[i].children;

      for (let j = 0; j < request.message.length; j++) {
        if (request.message[j].name.replace(' ', '') == staffData[0].innerHTML.replace(' ', '')) {
          excelDaySize = request.message[j].datas.length;
          // データが入っている列のみ抽出して、名前と院名を除く
          mirteDaySize = Array.prototype.slice.call(staffData).filter(function (value) {
            return value.innerHTML != "" && value.innerHTML != " " && value.innerHTML != "&nbsp;"
          }).length - 2;

          // エクセル側かミルテ側のどちらか日数の小さい方まで取り込む
          let daySize = Math.min(excelDaySize, mirteDaySize);

          noData = false;

          let dataList = request.message[j].datas;

          for (let k = 0; k < daySize; k++) {
            //〇＝出勤、✕＝休み
            if (dataList[k] == "〇") {
              staffData[k + 2].children[0].options[0].selected = true;
              staffData[k + 2].children[0].style.backgroundColor = "pink";
            }
            else if (dataList[k] == "✕") {
              staffData[k + 2].children[0].options[1].selected = true;
              staffData[k + 2].children[0].style.backgroundColor = "silver";
            }
            else {
              //暫定的に、〇、✕以外は”休”で反映させて対処している。
              staffData[k + 2].children[0].options[1].selected = true;
              staffData[k + 2].children[0].style.backgroundColor = "silver";
            }
          }
          // これ以降人物は一致しない
          break;
        }
      }

      if (noData) {
        // エクセル側かミルテ側のどちらか日数の小さい方まで取り込む
        let daySize = Math.min(excelDaySize, mirteDaySize);
        let absence = false;

        //欠勤リストに名前があるか判定
        for (let l = 0; l < absenceList.length; l++) {
          if (staffData[0].innerHTML.replace(' ', '') == absenceList[l].replace(' ', '')) {
            absence = true;
            break;
          }
        }

        if (absence) {
          //全て欠席入力
          for (let k = 0; k < daySize; k++) {
            staffData[k + 2].children[0].options[1].selected = true;
            staffData[k + 2].children[0].style.backgroundColor = "silver";
          }
          actuallyAbsenceList.push(staffData[0].innerHTML.replace(' ', ''));
        } else {
          // 欠勤リストにない場合は反映していない表示を出す
          cannotFindOnMirteList.push(StaffsOnMirteList.children[i].children[0].innerHTML);
        }
      }
    }
    sendResponse({
      cannotFindOnMirteList,
      actuallyAbsenceList,
    });
    return true;
  }
});
/******/ })()
;
//# sourceMappingURL=contentScript.js.map