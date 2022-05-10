
function test() {
  const name = "52002黃瑞得";
  const number = parseInt(name.slice(0, 5)) - 52000;
  const sheet_url = "https://docs.google.com/spreadsheets/d/1NQtJHJxXxg5WS3XAJ7gejJVOB_zvhhNe2T0_KwJ_s9o/edit?usp=sharing";
  const sheet_name = "Sheet1";
  const SpreadSheet = SpreadsheetApp.openByUrl(sheet_url);
  const sheet = SpreadSheet.getSheetByName(sheet_name);

  var status = "";
  for (let i = 1; i <= 13; i++) {
    status += get_status(sheet, i);
  }

  console.log(status);

  var date = new Date();
  var current_hour = Utilities.formatDate(date, "Asia/Taipei", "HH"); // get times
  var current_time = Utilities.formatDate(date, "Asia/Taipei", "MM/dd HH00");
  console.log(current_time);
};

function broadcast_remind() {
  var CHANNEL_ACCESS_TOKEN = "mHo57S8iRhU0qoijbs2Sb4Wu4VcIfyEqsZybKztdRYkpY4jkuIO56L97VpNJNZ8yRujvpWo7h1tGrfLOdWdRX9POy4b3JIZqs2j/11vBgB9nS+58o1zTzq3Pcs7q1gH23ooW+vH+QQPnZUboBjf+ZwdB04t89/1O/w1cDnyilFU=";
  var date = new Date();
  var current_hour = Utilities.formatDate(date, "Asia/Taipei", "HH");
  var remind_message = String(parseInt(current_hour) + 1) + "00 回報時間";

  push_to_line(CHANNEL_ACCESS_TOKEN, format_text_message(remind_message));

}

function broadcast_report_alert() {
  var CHANNEL_ACCESS_TOKEN = "mHo57S8iRhU0qoijbs2Sb4Wu4VcIfyEqsZybKztdRYkpY4jkuIO56L97VpNJNZ8yRujvpWo7h1tGrfLOdWdRX9POy4b3JIZqs2j/11vBgB9nS+58o1zTzq3Pcs7q1gH23ooW+vH+QQPnZUboBjf+ZwdB04t89/1O/w1cDnyilFU=";
  var sheet = get_google_sheet();
  var notdone = [];
  var tag_name = "";

  for (let i = 1; i <= 13; i++) {
    if (sheet.getRange(i + 1, 4).getValues()[0][0] == "") {
      notdone.push(sheet.getRange(i + 1, 7).getValues()[0][0]);
      // tag_name += "@" + sheet.getRange(i + 1, 1).getValues()[0][0] + sheet.getRange(i + 1, 2).getValues()[0][0] + " ";
      tag_name += i + " ";
    }
  }

  var date = new Date();
  var current_time = Utilities.formatDate(date, "Asia/Taipei", "HH00");
  var status = String(current_time) + "回報，報告班長，一班應到13員，實到" + (13 - notdone.length) + "員，";
  if (notdone.length != 0) {
    status += tag_name;
    status += "未完成回報，報告完畢";
    var report_msg = format_text_message(status);
    push_to_line(CHANNEL_ACCESS_TOKEN, report_msg);
  }
  else {
    return;
  }

}

function set_trigger(hhmm) {
  ScriptApp.newTrigger("broadcast_report_alert")
    .timeBased()
    .atHour(hhmm / 100)
    .nearMinute(hhmm % 100)
    .everyDays(1)
    .create()

  ScriptApp.newTrigger("broadcast_remind")
    .timeBased()
    .atHour((hhmm / 100) - 1)
    .nearMinute(hhmm % 100)
    .everyDays(1)
    .create()
}

function clear_status() {
  var sheet = get_google_sheet();
  sheet.getRangeList(['D2:F14']).clear();
}

function get_google_sheet() {
  const sheet_url = "https://docs.google.com/spreadsheets/d/1NQtJHJxXxg5WS3XAJ7gejJVOB_zvhhNe2T0_KwJ_s9o/edit?usp=sharing";
  const sheet_name = "Sheet1";
  const SpreadSheet = SpreadsheetApp.openByUrl(sheet_url);
  const sheet = SpreadSheet.getSheetByName(sheet_name);

  return sheet;
}

function get_status(sheet, number) {
  var id = sheet.getRange(number + 1, 1).getValues()[0][0];
  var name = sheet.getRange(number + 1, 2).getValues()[0][0];
  var phone = sheet.getRange(number + 1, 3).getValues()[0][0];
  var things = sheet.getRange(number + 1, 4).getValues()[0][0];
  var place = sheet.getRange(number + 1, 5).getValues()[0][0];
  var todo = sheet.getRange(number + 1, 6).getValues()[0][0];

  return String(id + " 姓名 " + name + "\n電話: " + phone + "\n做什麼事: " + things + "\n地點: " + place + "\n預計: " + todo + "\n\n");
};

function get_user_name(access_token, msg) {
  const user_id = msg.events[0].source.userId;
  const event_type = msg.events[0].source.type;

  switch (event_type) {
    case "user":
      var nameurl = "https://api.line.me/v2/bot/profile/" + user_id;
      break;
    case "group":
      var groupid = msg.events[0].source.groupId;
      var nameurl = "https://api.line.me/v2/bot/group/" + groupid + "/member/" + user_id;
      break;
  }

  try {
    var response = UrlFetchApp.fetch(nameurl, {
      "method": "GET",
      "headers": {
        "Authorization": "Bearer " + access_token,
        "Content-Type": "application/json"
      },
    });
    var namedata = JSON.parse(response);
    var reserve_name = namedata.displayName;
  }
  catch {
    reserve_name = "not avaliable";
  }
  return String(reserve_name)
}

function push_to_line(access_token, message) {
  var groupID = "C57faa1d5cea735f2fcecf12e429112c5";

  var url = "https://api.line.me/v2/bot/message/push";
  UrlFetchApp.fetch(url, {
    headers: {
      "Content-Type": "application/json; charset=UTF-8",
      Authorization: "Bearer " + access_token,
    },
    method: "post",
    payload: JSON.stringify({
      to: groupID,
      messages: message,
    }),
  });
}

function send_to_line(access_token, replyToken, message) {
  var url = "https://api.line.me/v2/bot/message/reply";
  UrlFetchApp.fetch(url, {
    headers: {
      "Content-Type": "application/json; charset=UTF-8",
      Authorization: "Bearer " + access_token,
    },
    method: "post",
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: message,
    }),
  });
}

function format_text_message(word) {
  let text_json = [
    {
      type: "text",
      text: word,
    },
  ];

  return text_json;
}

function doPost(e) {

  // LINE Messenging API Token
  var CHANNEL_ACCESS_TOKEN = "mHo57S8iRhU0qoijbs2Sb4Wu4VcIfyEqsZybKztdRYkpY4jkuIO56L97VpNJNZ8yRujvpWo7h1tGrfLOdWdRX9POy4b3JIZqs2j/11vBgB9nS+58o1zTzq3Pcs7q1gH23ooW+vH+QQPnZUboBjf+ZwdB04t89/1O/w1cDnyilFU=";

  var msg = JSON.parse(e.postData.contents);

  const replyToken = msg.events[0].replyToken;
  var userMessage = msg.events[0].message.text;
  var userName = get_user_name(CHANNEL_ACCESS_TOKEN, msg);
  var userNumber = parseInt(userName.slice(0, 5)) - 52000;
  var userId = msg.events[0].source.userId;
  var date = new Date();

  if (typeof replyToken === "undefined") {
    return;
  }

  // if (msg.events[0].source.type != "group"){
  //   send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message("This chatbot was designed for group chat"));
  // }


  // google sheet 
  // --------------------------------------------------------------------------------------
  var sheet = get_google_sheet();

  // system command
  // -------------------------------------------------------------------------------------
  if (userMessage === "#reset") {
    sheet.getRangeList(['D2:F14']).clear();
    send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message("data has been cleared"));
    return;
  }

  if (userMessage === "#help") {
    var msg = "README:\nhttps://github.com/huangjuite/report-linebot\n\ngoogle sheet:\nhttps://docs.google.com/spreadsheets/d/1NQtJHJxXxg5WS3XAJ7gejJVOB_zvhhNe2T0_KwJ_s9o/edit?usp=sharing";
    send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message(msg));
    return;
  }

  if (userMessage.includes("#alarm")) {
    var content = userMessage.split(" ");
    var reply_message = "";
    for (let i = 1; i < content.length; i++) {
      set_trigger(parseInt(content[i]));
      reply_message += content[i] + " ";
    }

    send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message("alarm " + reply_message + "set"));
    return;
  }

  if (userMessage === "#acknowledgement") {
    send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message("In remember of that rediculous April 2022.\n\t\t\t\tJui-Te Huang"));
    return;
  }

  // user command
  // -------------------------------------------------------------------------------------
  if (userMessage === "清查") {
    var current_time = Utilities.formatDate(date, "Asia/Taipei", "MM/dd HH00");
    var notdone = [];
    for (let i = 1; i <= 13; i++) {
      if (sheet.getRange(i + 1, 4).getValues()[0][0] == "") {
        notdone.push(String(i) + " ");
      }
    }

    var status = "第一班 " + current_time + " 回報\n應到13員 實到" + (13 - notdone.length) + "員 發燒0員\n";
    if (notdone.length != 0) {
      status += "缺：" + notdone + "\n";
    }

    status += "\n";
    for (let i = 1; i <= 13; i++) {
      status += get_status(sheet, i);
    }

    send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message(status));
    return;
  }

  if (userMessage === "清查頑劣份子") {
    var marked_id = [2, 3, 4, 5, 7, 11];
    var current_time = Utilities.formatDate(date, "Asia/Taipei", "MM/dd HH00");

    var status = "第一班 " + current_time + " 回報\n頑劣份子6員\n\n";
    marked_id.forEach(function display_ids(item, indx) {
      status += get_status(sheet, item);
    });

    send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message(status));
    return;
  }


  // report command
  // ------------------------------------------------------------------------------------
  if (!(userMessage.slice(0, 3) === "回報 ")) {
    return;
  }
  userMessage = userMessage.slice(3, userMessage.length);

  if (userName === "吳芷葳") {
    send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message("Hi 班長"));
    return;
  }

  content = userMessage.split(' ');
  var things = content[0];
  var plcae = content[1];
  var todo = content[2];

  send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message(userName + " " + things + " " + plcae + " " + todo));

  sheet.getRange(userNumber + 1, 4).setValue(things);
  sheet.getRange(userNumber + 1, 5).setValue(plcae);
  sheet.getRange(userNumber + 1, 6).setValue(todo);
  sheet.getRange(userNumber + 1, 7).setValue(userId);

  for (let i = 1; i <= 13; i++) {
    if (sheet.getRange(i + 1, 4).getValues()[0][0] == "") {
      // not done
      return;
    }
  }

  // all set
  var current_time = Utilities.formatDate(date, "Asia/Taipei", "MM/dd HH00");
  var status = "第一班 " + current_time + " 回報\n應到13員 實到13員 發燒0員\n\n";
  for (let i = 1; i <= 13; i++) {
    status += get_status(sheet, i);
  }

  push_to_line(CHANNEL_ACCESS_TOKEN, format_text_message(status));

}
