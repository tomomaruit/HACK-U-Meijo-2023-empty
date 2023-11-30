// アクセストークン,URL,スプレッドシートIDを定義 
const LINE_TOKEN = PropertiesService.getScriptProperties().getProperty("LINE_TOKEN"); // LINE Botのアクセストークン
const LINE_URL = 'https://api.line.me/v2/bot/message/reply'; // LINE Bot の要件に沿ったリンクを定義
const SHEET_ID = PropertiesService.getScriptProperties().getProperty("SHEET_ID"); // このGASが紐づいているスプレッドシートの定義

// スプレッドシート制御用
const ss = SpreadsheetApp.openById(SHEET_ID); //使用するスプレッドシート自体をssとして定義
const logs = ss.getSheetByName('logs'); // 設定シート定義
const sublogs = ss.getSheetByName('sublogs'); // 設定シート定義

// 関数に依存しない共通部分の定義
// ここでユーザからの入力を分ける
const sheetname = ['タワー75','共通講義棟南','共通講義棟北','共通講義棟東','研究実験棟4','11号館','12号館','DW','DN','DS'];
const username1 = ['T','S','N','E','R4','11','12','DW','DN','DS'];

// メッセージ送信時に起動するdoPost関数
function doPost(e) {
  // 受け取ったメッセージをJSON形式から抽出
  const json = JSON.parse(e.postData.contents);
  const reply_token = json.events[0].replyToken;
  const messageId = json.events[0].message.id;
  const messageType = json.events[0].message.type;
  const messageText = json.events[0].message.text;
  const userID = json.events[0].source.userId;

  //ログ用のを日付生成
  let date= new Date();
  logs.appendRow([date,userID,messageText]); // 日時、userID、送信された内容をログとして残す
  sublogs.appendRow([date,userID]); // 日時、userID、送信された内容をログとして残す

  if (typeof reply_token === 'underfined') {
    return;
  }
  //メッセージ解析してチュートリアルを始めるか解析
  if (messageText == 'スタート'){ // 「スタート」と入力されたら
    notify(reply_token); // チュートリアル開始
    sublogs.appendRow(['スタート']);
  }
  else{ // そうでなければ
    const splittext = messageText.split("\n"); // 配列に分ける
    judgetoolno(splittext,reply_token); // 配列0番目検証関数へ移行
  }
  return;
}

// 配列の0番目を読んで検索方法を解析してメッセージ送信まで行う関数
function judgetoolno(splittext,reply_token){
  if (splittext[0] == '1'){ // もし1なら検索方法1を起動
    const buildname = splittext[1]; // 建物名抽出
    const roomname = splittext[2]; // 部屋番号抽出
    const hairetsu = username1.length; // 配列の長さを定義
    for (i = 0; i <= hairetsu ; i++){
      if (buildname == username1[i]){ // LINEのメッセージと配列内の要素が一致したら
        const searchsheet = ss.getSheetByName(sheetname[i]); // 検索対象シートを定義
        sublogs.appendRow(['検索方法1','検索シート名：',sheetname[i],'教室番号：',roomname]); // 動作検証用のログ記入
        let result = []; // 結果保持用の配列
        const lastrow = searchsheet.getLastRow(); // 最終行が何行目なのか取得する
        for (i = 2 ; i<= lastrow; i++){ // 縦方向検索
          let searchcel = 'B' + i; // 検索するセルは確実にB列に存在するので形式指定する
          let sheetroomname = searchsheet.getRange(searchcel).getValue();
          if (roomname == sheetroomname){
            for (j = 4; j <= 38 ; j ++){ //横方向探索に切替
              let sheetclasscode = searchsheet.getRange(i,j).getValue(); // 検索対象セルの内容定義
              if (sheetclasscode === ''){ // 検索結果が空白なら
                let nullcellname = searchsheet.getRange(1,j).getValue(); // 同列の一行目教室名取得
                result.push(nullcellname); // 結果用配列の末尾に追加
              }
            }
          }
        }
        sublogs.appendRow(result); // 検索結果をログ出力
        let message = result.join('\n'); // 配列に入ったままでは送信できないので、改行してメッセージとして送信できる形にする
        sendLINE(reply_token,message); // 返信実行関数起動
      }
    }
  }

  else if (splittext[0] == '2'){ // もし2なら検索方法2を起動
    const buildname = splittext[1]; // 建物名抽出
    const whatdatetime = splittext[2]; // 曜日時限抽出
    const hairetsu = username1.length; // 配列の長さを定義
    for (i = 0; i <= hairetsu ; i++){
      if (buildname == username1[i]){ // LINEのメッセージと配列内の要素が一致したら
        const searchsheet = ss.getSheetByName(sheetname[i]); // 検索設定シート定義
        sublogs.appendRow(['検索方法2','検索シート名：',buildname,'検索時間：',whatdatetime]); // 動作検証用のログ記入
        let result = []; // 結果保持用の配列
        for (i = 4 ; i<= 38; i++){ // 4と38は不変かつどのシートでも不変なので変数で呼び出ししない
          let sheetdatename = searchsheet.getRange(1,i).getValue(); // スプシから取得した曜日時限
          if (whatdatetime == sheetdatename){ // LINEの送信内容と一致した時
            for (j = 2; j <= 40 ; j ++){ // 縦方向検索に切替
              let sheetclasscode = searchsheet.getRange(j,i).getValue(); // 検索対象セルの内容定義
              if (sheetclasscode === ''){ // 検索結果が空白なら
                let nullcellname = searchsheet.getRange(j,2).getValue(); // 同行の教室名取得
                result.push(nullcellname);
              }
            }
          }
        }
        sublogs.appendRow(result); 
        let message = result.join('\n'); 
        sendLINE(reply_token,message); 
      }
    }
  }
  else if (splittext[0] == '3'){
    const buildname = splittext[1]; // 建物名抽出
    const whatdatetime = splittext[2]; // 曜日時限抽出
    const roomname = splittext[3]; // 教室番号抽出
    const hairetsu = username1.length; // 配列の長さを定義
    for (i = 0; i <= hairetsu ; i++){
      if (buildname == username1[i]){ // LINEのメッセージと配列内の要素が一致したら
        const searchsheet = ss.getSheetByName(sheetname[i]); // 検索設定シート定義
        sublogs.appendRow(['検索方法3','検索シート名：',buildname,'曜日時限：',whatdatetime,'教室番号：',roomname]); // 動作検証用のログ記入
        let result = []; // 結果保持用の配列
        for (i = 4 ; i<= 38; i++){ // 4と38は不変かつどのシートでも不変なので変数で呼び出ししない
          let sheetdatename = searchsheet.getRange(1,i).getValue(); // スプシから取得した曜日時限
          if (whatdatetime == sheetdatename){ // LINEの送信内容と一致した時
            for (j = 2; j <= 40 ; j ++){ // 縦方向検索に切替
              let sheetclasscode = searchsheet.getRange(j,2).getValue(); // 検索対象セルの内容定義
              if (sheetclasscode == roomname){ // 検索結果が空白なら
                let classno = searchsheet.getRange(j,i).getValue(); // 同行の教室名取得
                result.push(classno);
              }
            }
          }
        }
        sublogs.appendRow(result); 
        let codemessage = result.join('\n'); 
        const message = 'https://gkmsyllabus.meijo-u.ac.jp/camweb/slbssbdr.do?value(risyunen)=2023&value(semekikn)=1&value(kougicd)=' + codemessage
        sendLINE(reply_token,message); 
      }
    }


    
  }
  else {
    errorcode = 100;
    error(errorcode,reply_token);
  }
}

// 説明用関数
function notify(reply_token) {
  const set_sheet = ss.getSheetByName('ユーザ説明用'); // 設定シート定義
  const msg = set_sheet.getRange("B1").getValue();
  const buildingmsg = set_sheet.getRange("B2").getValue();
  const ex1 = set_sheet.getRange("B3").getValue();
  const ex2 = set_sheet.getRange("B4").getValue();
  const ex3 = set_sheet.getRange("B5").getValue();
  // LINE側の要件に合わせる
  const option = {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + LINE_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': reply_token,
      'messages': [
        {
        'type': 'text',
        'text': msg
        },
        {
        'type': 'text',
        'text': buildingmsg
        },
        {
        'type': 'text',
        'text': ex1
        },
        {
        'type': 'text',
        'text': ex2
        },
        {
        'type': 'text',
        'text': ex3
        }
      ],
    }),
  }
  UrlFetchApp.fetch(LINE_URL,option);
  return;
}

// エラー用に使用する関数
function error(errorcode,reply_token) {
  const option = {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + LINE_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': reply_token,
      'messages': [
        {
        'type': 'text',
        'text': 'エラーが発生しました。最初からやり直してください。\nエラーコード:'+ errorcode
        },
      ],
    }),
  }
  UrlFetchApp.fetch(LINE_URL,option);
  return;
}

// LINE送信用関数
function sendLINE(reply_token,result) {
    const option = {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + LINE_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': reply_token,
      'messages': [{
        'type': 'text',
        'text': result,
      }],
    }),
  }
  UrlFetchApp.fetch(LINE_URL,option);
  return;
}
