
function get_status(sheet, number) {
  var id = sheet.getRange(number + 1, 1).getValues()[0][0];
  var name = sheet.getRange(number + 1, 2).getValues()[0][0];
  var phone = sheet.getRange(number + 1, 3).getValues()[0][0];
  var things = sheet.getRange(number + 1, 4).getValues()[0][0];
  var place = sheet.getRange(number + 1, 5).getValues()[0][0];
  var todo = sheet.getRange(number + 1, 6).getValues()[0][0];

  return String(id + " 姓名 " + name + "\n電話 " + phone + "\n做什麼事 " + things + "\n地點 " + place + "\n預計 " + todo + "\n\n");
};

function test() {
  const name = "52002黃瑞得";
  const number = parseInt(name.slice(0, 5)) - 52000;
  const sheet_url = "https://docs.google.com/spreadsheets/d/1NQtJHJxXxg5WS3XAJ7gejJVOB_zvhhNe2T0_KwJ_s9o/edit?usp=sharing";
  const sheet_name = "Sheet1";
  const SpreadSheet = SpreadsheetApp.openByUrl(sheet_url);
  const sheet = SpreadSheet.getSheetByName(sheet_name);

  // sheet.getRangeList(['D2:F14']).clear();
  // sheet.getRange(number + 1, 4).setValue("吃飯");
  // sheet.getRangeList(['D2:F14']).setValue("吃飯");

  var status = "";
  for (let i = 1; i <= 13; i++) {
    status += get_status(sheet, i);
  }

  console.log(status);
  console.log(number);

  var date = new Date();
  var current_hour = Utilities.formatDate(date, "Asia/Taipei", "HH"); // get times
  var current_time = Utilities.formatDate(date, "Asia/Taipei", "MM/dd HH00");
  console.log(current_date);

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
  var date = new Date();

  if (typeof replyToken === "undefined") {
    return;
  }

  // google sheet 
  // --------------------------------------------------------------------------------------
  const sheet_url = "https://docs.google.com/spreadsheets/d/1NQtJHJxXxg5WS3XAJ7gejJVOB_zvhhNe2T0_KwJ_s9o/edit?usp=sharing";
  const sheet_name = "Sheet1";
  const SpreadSheet = SpreadsheetApp.openByUrl(sheet_url);
  const sheet = SpreadSheet.getSheetByName(sheet_name);


  // system command
  // -------------------------------------------------------------------------------------

  if (userMessage === "清查") {
    var current_time = Utilities.formatDate(date, "Asia/Taipei", "MM/dd HH00");

    var status = "第一班 "+current_time+" 回報\n應到13員 實到13員 發燒0員\n\n";
    for (let i = 1; i <= 13; i++) {
      status += get_status(sheet, i);
    }

    send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message(status));
    return;
  }

  if (userMessage === "reset") {
    sheet.getRangeList(['D2:F14']).clear();
    send_to_line(CHANNEL_ACCESS_TOKEN, replyToken, format_text_message("data has been cleared"));
    return;
  }


  // r command
  // ------------------------------------------------------------------------------------

  var current_hour = Utilities.formatDate(new Date(), "Asia/Taipei", "HH"); // get times

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


}
