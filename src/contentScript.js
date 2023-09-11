'use strict';

const absenceList = ["廣川創太", "坪井恵"];

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {

  if (request.message == "get YM") {
    let wndlct = window.location;
    let url = wndlct.host + wndlct.port + wndlct.pathname;
    if (url !== "dev.mirte.jp/attendance/edit" && url !== "mirte.jp/attendance/edit") {
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

    //request.messageには、popup.jsから送信されたdatalist[]が送られてくる。
    for (let i = 0; i < StaffsOnMirteList.children.length; i += 2) {
      let noData = true;

      for (let j = 0; j < request.message.length; j++) {
        let staffData = StaffsOnMirteList.children[i].children;
        if (request.message[j].name.replace(' ', '') == staffData[0].innerHTML.replace(' ', '')) {
          let excelDaySize = request.message[j].datas.length;
          // データが入っている列のみ抽出して、名前と院名を除く
          let mirteDaySize = Array.prototype.slice.call(staffData).filter(function (value) {
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
            }
            else if (dataList[k] == "✕") {
              staffData[k + 2].children[0].options[1].selected = true;
            }
            else {
              //暫定的に、〇、✕以外は”休”で反映させて対処している。
              staffData[k + 2].children[0].options[1].selected = true;
            }
          }
        }
        else {
          let excelDaySize = request.message[j].datas.length;
          // データが入っている列のみ抽出して、名前と院名を除く
          let mirteDaySize = Array.prototype.slice.call(staffData).filter(function (value) {
            return value.innerHTML != "" && value.innerHTML != " " && value.innerHTML != "&nbsp;"
          }).length - 2;
          // エクセル側かミルテ側のどちらか日数の小さい方まで取り込む
          let daySize = Math.min(excelDaySize, mirteDaySize);
          let absence = false;

          for (let l = 0; l < absenceList.length; l++) {
            if (staffData[0].innerHTML.replace(' ', '') == absenceList[l].replace(' ', '')) {
              absence = true;
              break;
            }
          }
          if (absence) {
            for (let k = 0; k < daySize; k++) {
              staffData[k + 2].children[0].options[1].selected = true;
            }
            actuallyAbsenceList.push(staffData[0].innerHTML.replace(' ', ''));
          }
        }
        break;
      }

      if (noData) {
        cannotFindOnMirteList.push(StaffsOnMirteList.children[i].children[0].innerHTML);
      }
    }
    sendResponse({
      cannotFindOnMirteList,
      actuallyAbsenceList,
    });
    return true;
  }
});