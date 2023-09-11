'use strict';

import '../node_modules/bootstrap';
import './popup.scss';
import * as XLSX from 'xlsx';

//chrome.tabs.sendMessage()内（13行目）で宣言した変数msgをhandleFileAsync()内でも参照したかった為グローバルで宣言。
let msg = null;
let YM = "";

window.onload = function(){
  document.getElementById("my-input-btn").addEventListener("click", () => {
    document.getElementById("input-file").click();
  });
  const myInput = document.getElementById("input-file");
  myInput.addEventListener('change', (event) => {
    if(event.target.files[0]){
      document.getElementById("filename").innerHTML = event.target.files[0].name;
    }
    else{
      document.getElementById("filename").innerHTML = "ファイルが選択されていません"
    }
  });

  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    let getYMmsg = "get YM";
    chrome.tabs.sendMessage(tabs[0].id, { message: getYMmsg }, (response) => {
      if (!response) {
        alert("エラーが発生しました。実行前にページを更新して下さい。");
        return;
      }
      else if (response == "no mirte"){
        alert("Mirteの勤怠情報ページを開いて下さい。");
      }
      else{
        YM = response;
      }
    });
  });
}

document.getElementById("btn-get").addEventListener("click", () => {
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    chrome.tabs.sendMessage(tabs[0].id, { message: msg }, (content) => {
      if (!content) {
        alert("エラーが発生しました。実行前にページを更新して下さい。");
        return;
      } else if (content === "sheet not found") {
        alert("勤怠表ファイルに対応するシートが見つかりません。");
        return;
      } else if (content === "version failed") {
        alert("シートのバージョンが異なります。最新のバージョンを使用してください。");
        return;
      }
      document.getElementById("target-table").innerHTML = content.cannotFindOnMirteList + "は勤怠表ファイルに名前が見つからないため、反映していません。";
      document.getElementById("target-table-absence").innerHTML = content.actuallyAbsenceList + "は全て”休み”にしています。";
    });
  });
});

async function handleFileAsync(e) {
  const file = e.target.files[0];
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);

  //勤怠情報ページの日付に対応したシートを読み込んでいる
  let targetSheetName = "R" + String(Number(YM.slice(0,4)) - 2018) + "." + String(Number(YM.slice(-2))) + "月";

  const worksheet = workbook.Sheets[targetSheetName];
  document.getElementById("worksheet-name").innerHTML = targetSheetName;

  if (worksheet == undefined) {
    msg = "sheet not found";
    return;
  }

  //最終的にcontentScript.jsへ送信する配列
  let datalist = [];
  let version = worksheet[XLSX.utils.encode_cell({c:0, r:0})];

  // セルA1にバージョン指定がなければ失敗
  if (!version || version.v !== "ver1.3") {
    msg = "version failed";
    return;
  }

  //xlsxの上から4行目から読み込みを始める。現時点では有効な勤怠データの入っている行（最下行から20番目の行）までに範囲指定している。
  for (let curr = 3; curr <= XLSX.utils.decode_range(worksheet['!ref']).e.r - 1; curr++){
    if (worksheet[XLSX.utils.encode_cell({c:4, r:curr})] == undefined) {
      break;
    }

    //名前と勤怠情報を入れる連想配列
    let rowlist = {};
    //勤怠情報を入れる配列
    let rowdatas = [];

    //横長タイプ（H31.3月）の出勤表の勤怠データ5列目〜34列目のデータを収集している
    for(let curc = 5; curc <= XLSX.utils.decode_range(worksheet['!ref']).e.c; curc++){
      if (worksheet[XLSX.utils.encode_cell({c:curc, r:curr})] == undefined) {
        break;
      }
      rowdatas.push(worksheet[XLSX.utils.encode_cell({c:curc, r:curr})].v);
    }
    //取得した行のスタッフ名
    rowlist.name = worksheet[XLSX.utils.encode_cell({c:4, r:curr})].v;
    //取得した行の勤怠情報
    rowlist.datas = rowdatas;
    //連想配列をリストに追加
    datalist.push(rowlist);
  }

  //contentScript.jsへ送信
  msg = datalist;
}

document.getElementById("input-file").addEventListener("change", handleFileAsync, false);