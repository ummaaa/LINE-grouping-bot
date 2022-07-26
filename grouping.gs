var ACCESS_TOKEN = ''; // LINE botのチャンネルアクセストークン入力
var sheetUrl = ''; // 使うスプレッドシートのurlを入力
var sheetName = ''; // 使うシートの名前を入力
var ss = SpreadsheetApp.openByUrl(sheetUrl);
var roomsUpperLeftCorner = 'A2';
var usersInfoUpperLeftCorner = 'E2';
var lastRow = 1000;

function doPost(e) {

  // 送られてきた値をパース
  var event = JSON.parse(e.postData.contents).events[0];

  // 送信者のtypeを取得 
  var type = event.source.type;

  // 現在作られているルームの情報を取得
  var rooms = readSheetAsDict(sheetName, roomsUpperLeftCorner, 2);

  // 現在の各ユーザの情報を取得
  var usersInfo = readSheetAsDict(sheetName, usersInfoUpperLeftCorner, 3);

  // ユーザidを取得
  var userId, userName;
  if (type == 'user') {
    userId = event.source.userId;
    // userName = getUserName(userId); // LINEのアカウントを複数アカウント作れず実験できないので，LINEでの表示名は使わない
    // initUser(userName, userId);
  } else if (type == 'group') {
    pushMessage(event.source.groupId, 'このbotは個人アカウントでしか使用できません');
  } else if (type == 'room') {
    pushMessage(event.source.roomId, 'このbotは個人アカウントでしか使用できません');
  }


  // 送られてきたmessageを取得
  var message = event.message.text;

  userName = message.split(/\r\n|\n/)[0];
  if (userName == 'ヘルプ') {
    pushMessage(userId, help(6));
    return;
  }
  if (!(userName in usersInfo)) {
    initUser(userName, userId);
    reply = `${userName}さんこんにちは\n${help(7)}`;
    pushMessage(userId, reply);
    return;
  }

  message = message.split(/\r\n|\n/).slice(1).join(',')

  // messageによって行う処理を決める
  if (message == '状況確認') {
    var roomId = usersInfo[userName][1];
    if (String(roomId).length == 0) {
      var reply = `${userName}さんはまだルームに参加していません\n${help(2)}`
      pushMessage(userId, reply);
    }
    else {
      var roomSize = rooms[roomId][0].split(',').length;
      var members = rooms[roomId][0].replace(',', ', ');
      var reply = `${userName}さんはルーム${roomId}に参加しています\n現在のルーム${roomId}の参加メンバーは以下の${roomSize}人です\n${members}(敬称略)\n${help(4)}`;
      pushMessage(userId, reply);
    }
  }
  else if (message == 'ルーム作成') {
    var roomId = getRandomInt(1000, 10000);
    while (roomId in rooms) roomId = getRandomInt(1000, 10000);
    rooms[roomId] = ['', 0];
    saveDictToSheet(sheetName, rooms, roomsUpperLeftCorner);
    addUser(userName, roomId);
    var reply = `ルーム${roomId}を作成し参加しました\n${help(4)}\n${help(3)}`;
    pushMessage(userId, reply);
  }
  else if (isInt(message)) {
    if (message in rooms) {
      var roomId = message;
      var groupingNow = rooms[roomId][1];
      if (groupingNow == 0) {
        var reply = `ルーム${roomId}に参加しました\n${help(4)}\n${help(3)}`;
        addUser(userName, roomId);
        pushMessage(userId, reply);
      }
      else {
        var reply = `ルーム${roomId}は締め切られました`
        pushMessage(userId, reply);
      }
    }
    else {
      var reply = `ルーム${message}は存在しません\n${help(2)}\n${help(3)}`;
      pushMessage(userId, reply);
    }
  }
  else if (message == 'ルーム締切') {
    var roomId = usersInfo[userName][1];
    if (String(roomId).length == 0) {
      var reply = `まだルームに参加していません\n${help(2)}\n${help(3)}`;
      pushMessage(userId, reply);
      return;
    }
    rooms[roomId][1] = 1;
    saveDictToSheet(sheetName, rooms, roomsUpperLeftCorner);
    var roomSize = rooms[roomId][0].split(',').length;
    var reply = `ルーム${roomId}を締め切りました\n${help(5, roomSize)}`;
    pushMessage(userId, reply);
  }
  else {
    var roomId = usersInfo[userName][1];
    if (String(roomId).length == 0) {
      pushMessage(userId, `${userName}さん\n` + help(7));
      return;
    }
    var groupingNow = rooms[roomId][1];
    if (groupingNow == 0) {
      pushMessage(userId, `${userName}さん\n` + help(7));
      return;
    }
    // group分け中
    var groupingRes = grouping(roomId, message);
    if (!groupingRes) {
      var roomSize = rooms[roomId][0].split(',').length;
      var reply = `分け方が正しくありません\n${help(5, roomSize)}`;
      pushMessage(userId, reply);
    }
    else {
      var members = rooms[roomId][0].split(',');
      // ルームroomIdの参加者全員に結果を送信する
      members.forEach(member => {
        uid = usersInfo[member][0];
        pushMessage(uid, `${member}さん\n` + groupingRes);
      })
      // pushMessage(userId, groupingRes);
      deleteRoom(roomId);
    }
  }
}


// ユーザidがuserIdの人に，LINEでtextというメッセージを送信する
function pushMessage(userId, text) {

  var payload = {
    to: userId,
    'messages': [
      {
        'type': 'text',
        'text': text
      }
    ]
  };

  var options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + ACCESS_TOKEN,
    },
    'payload' : JSON.stringify(payload)
  };

  // lineメッセージ送信
  var pushUrl = 'https://api.line.me/v2/bot/message/push';
  UrlFetchApp.fetch(pushUrl, options);
}

// ユーザidからLINEでの表示名を取得: https://teratail.com/questions/304140
function getUserName(userId) {
  var endPoint = `https://api.line.me/v2/bot/profile/${userId}`;
  var res = UrlFetchApp.fetch(endPoint, {
    headers: {
      "Content-Type": "application/json; charset=UTF-8",
      Authorization: "Bearer " + ACCESS_TOKEN,
    },
    method: "GET",
  });

  var userName = JSON.parse(res.getContentText()).displayName;
  return userName;
}

// ランダムなmin以上max未満の整数を取得
function getRandomInt(min, max) {
  min = Math.ceil(min);
  max = Math.floor(max);
  return Math.floor(Math.random() * (max - min) + min);
}

// xが整数かどうか判定する
function isInt(x) {
  if (isNaN(x)) return false;
  else return Number.isInteger(parseFloat(x));
}

// order='x3'だったら3人組をできるだけ作る
// order='2:3,5:2'だったら2人組を3個，5人組を2個作る
// orderの内容がおかしかったらfalseを，大丈夫ならグループ分けの結果のstringを返す
function grouping(roomId, order) {

  if (String(order).length == 0) return false;

  var members = readSheetAsDict(sheetName, roomsUpperLeftCorner, 1)[roomId][0].split(',');
  var membersNum = members.length;

  if (order[0] == 'x') {
    if (!isInt(order.slice(1))) return false;
    var g = parseInt(order.slice(1));
    order = `${g}:${Math.floor(membersNum/g)},${membersNum%g}:1`;
  }

  order = order.split(',');
  var groupIdx = 1;
  var groups = [];
  order.forEach(o => {
    var n = o.split(':')[0]; // n人組を
    var m = o.split(':')[1]; // m個作る
    for (var i = 0; i < m; i++) {
      for (var j = 0; j < n; j++) {
        groups.push(groupIdx);
      }
      groupIdx++;
    }
  })

  if (groups.length != membersNum) return false;

  groups = shuffle(groups);
  var usersInfo = readSheetAsDict(sheetName, usersInfoUpperLeftCorner, 3);
  groupMember = {};
  for (var i = 0; i < membersNum; i++) {
    usersInfo[members[i]][2] = groups[i];
    if (!(groups[i] in groupMember)) groupMember[groups[i]] = members[i];
    else groupMember[groups[i]] += `, ${members[i]}`;
  }
  saveDictToSheet(sheetName, usersInfo, usersInfoUpperLeftCorner);
  groupingRes = '';
  for (var groupIdx in groupMember) {
    if (groupingRes.length > 0) groupingRes += '\n';
    groupingRes += `グループ${groupIdx}\n${groupMember[groupIdx]}`;
  }
  groupingRes = `ルーム${roomId}のグループ分けの結果は以下のようになりました(敬称略)\n` + groupingRes;
  return groupingRes;
}

// userNameさんのuserIdを登録する
function initUser(userName, userId) {
  var usersInfo = readSheetAsDict(sheetName, usersInfoUpperLeftCorner, 1);
  if (userName in usersInfo) return;
  usersInfo[userName] = [userId];
  saveDictToSheet(sheetName, usersInfo, usersInfoUpperLeftCorner);
}

// userNameさんを部屋roomIdに追加する
function addUser(userName, roomId) {
  var rooms = readSheetAsDict(sheetName, roomsUpperLeftCorner, 1);
  if (!(roomId in rooms)) return; // roomIdが存在しなかったら何もしない
  if (rooms[roomId][0].length > 0) rooms[roomId][0] += ',';
  rooms[roomId][0] += userName;
  var usersInfo = readSheetAsDict(sheetName, usersInfoUpperLeftCorner, 2);
  usersInfo[userName][1] = roomId;
  saveDictToSheet(sheetName, rooms, roomsUpperLeftCorner);
  saveDictToSheet(sheetName, usersInfo, usersInfoUpperLeftCorner);
}

// ルームroomIdを消す
function deleteRoom(roomId) {
  var rooms = readSheetAsDict(sheetName, roomsUpperLeftCorner, 2);
  if (!(roomId in rooms)) return;
  var usersInfo = readSheetAsDict(sheetName, usersInfoUpperLeftCorner, 3);
  var members = rooms[roomId][0].split(',');
  members.forEach(member => {
    delete usersInfo[member];
  })
  delete rooms[roomId];
  saveDictToSheet(sheetName, rooms, roomsUpperLeftCorner);
  saveDictToSheet(sheetName, usersInfo, usersInfoUpperLeftCorner);
}

// sheetNameの左上のセルをupperLeftCornerとするwidth行についてend行目までを，最左列をkey，それ以外のlistをvalueとするdictを返す．keyが空の行は無視する．
// asListがtrueのときは，valueを,で区切ったStringの配列とする
function readSheetAsDict(sheetName, upperLeftCorner, numValue=1, end=lastRow) {
  
  var res = {};

  var charL = upperLeftCorner[0]; // 左上隅のマスの列のアルファベット
  var idxL = parseInt(upperLeftCorner.slice(1)); // 左上隅のマスの行番号
  var columnL = fetchListColumn(sheetName, charL, idxL, end); // 最左の行の要素のリスト
  var filledRow = []; // 空じゃない行のindex
  for (var i = 0; i < columnL.length; i++) {
    if (String(columnL[i]).length > 0) {
      filledRow.push(i);
      res[columnL[i]] = [];
    }
  }
  
  for (var i = 1; i <= numValue; i++) {
    var charI = String.fromCharCode(charL.charCodeAt(0) + i);
    var columnI = fetchListColumn(sheetName, charI, idxL, end);
    filledRow.forEach(j => {
      res[columnL[j]].push(columnI[j]);
    })
  }
  
  return res;
}

// column(アルファベットで指定)列のstart行目からend行目までをリストで取得
function fetchListColumn(sheetName, column, start, end) {
  var sheet = ss.getSheetByName(sheetName);
  var query = `${column}${start}:${column}${end}`;
  var raw = sheet.getRange(query).getValues();
  var list = Array(end - start + 1);
  for (var i=0; i<list.length; i++) {
    list[i] = raw[i][0];
  }
  return list;
}

// dictをsheetNameのシートに保存する．upperLeftCornerは，保存するシートの領域の左上のセルを表す文字列('D3'みたいな)．
// dictのvalueはlistにする
function saveDictToSheet(sheetName, dict, upperLeftCorner) {
  var keys = Object.keys(dict);
  var num = keys.length;

  var content = [];
  for (var i=0; i<num;i++){
    content.push([]);
    content[i].push(keys[i]);
    dict[keys[i]].forEach(value => {
      content[i].push(value);
    })
  }
  
  var charL = upperLeftCorner[0]; // 左上隅のマスの列のアルファベット
  var charR = String.fromCharCode(charL.charCodeAt(0) + dict[keys[0]].length); // 右上隅のマスの列のアルファベット

  var idxL = parseInt(upperLeftCorner.slice(1)); // 左上隅のマスの行番号
  var idxR = idxL + num - 1; // 右下隅のマスの行番号

  var clearRange = `${charL}${idxL}:${charR}${lastRow}`; // これから書き込む場所から下
  var saveRange = `${charL}${idxL}:${charR}${idxR}`; // これから書き込む場所

  var sheet = ss.getSheetByName(sheetName);
  sheet.getRange(clearRange).clearContent(); // これから書き込む場所から下はクリアしておく
  sheet.getRange(saveRange).setValues(content); // 書き込む
}

// shuffleした配列を返す
function shuffle (array) {
  console.log(array);
  for (let i = array.length - 1; i >= 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}


function help(type, roomSize=null) {
  var res;
  if (type == 1) {
    res = '初めての方:\nユーザネームを送信してください\nこのbotで「1行目に~を入力し2行目に~~を入力して送信する」という表現は，1つのメッセージ中で改行して入力し，送信することを意味します\nヘルプと送信すると，使い方を見られます';
  }
  else if (type == 2) {
    res = '新しくルームを作成する場合は1行目にユーザネームを入力し，2行目に「ルーム作成」と入力して送信してください\n既存のルームに参加する場合は1行目にユーザネームを入力し，2行目に4桁のルーム番号を入力して送信してください';
  }
  else if (type == 3) {
    res = `自分および自分の参加しているルームの状況を知りたい場合は1行目にユーザネームを入力し，2行目に「状況確認」と入力して送信してください`;
  }
  else if (type == 4) {
    res = `参加者の募集を終える場合は1行目にユーザネームを入力し，2行目に「ルーム締切」と入力して送信してください`;
  }
  else if (type == 5) {
    res = `${roomSize}人の分け方を教えてください\n送信内容の1行目はユーザネームとし，2行目以降は以下の例にならってください\nできるだけたくさんの3人グループを作りたい場合: 2行目にx3と入力し送信する\n2人グループを3個，4人グループを5個作りたい場合: 2行目に2:3，3行目に4:5\nと入力して送信する(:は半角)`;
  }
  else if (type == 6) {
    res = help(1) + '\n\n初めてでない方:\n' + help(2) + '\n' + help(3);
  }
  else if (type == 7) {
    res = help(2) + '\n' + help(3); 
  }
  return res;
}

