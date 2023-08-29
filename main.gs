const scriptProperties = PropertiesService.getScriptProperties();
const appToken = scriptProperties.getProperty("APP_TOKEN")
const spreadSheet = SpreadsheetApp.getActiveSpreadsheet()

function main() {
  const emailList = convertJson(spreadSheet.getSheetByName("名簿").getDataRange().getValues(), ["Name", "Email"])
  emailList.forEach((t => {
    channelID = getChannelID(getSlackIdByEmail(t["Email"]))
    sendSlackDM(channelID)
  }))
}

function getSlackIdByEmail(email) {
  var options = {
    "method": "post",
    "contentType": "application/x-www-form-urlencoded",
    "payload": {
      "token": appToken,
      "email": email
    }
  }
  var response = UrlFetchApp.fetch('http://slack.com/api/users.lookupByEmail', options)
  var res = Json.parse(response)
}

// SlackのUserIDからDMのChannelIDを取得
function getChannelID(memberId) {
  var options = {
    "method": "post",
    "contentType": "application/x-www-form-urlencoded",
    "payload": {
      "token": appToken,
      "users": memberId
    }
  };
  var response = UrlFetchApp.fetch('https://slack.com/api/conversations.open', options);
  var res = JSON.parse(response);
  if (res.ok) {
    return res.channel.id;
  } else {
    Logger.log(res);
    return null;
  }
}

function sendSlackDM(channelID, message) {
  var message_options = {
    "method": "post",
    "contentType": "application/x-www-form-urlencoded",
    "payload": {
      "token": appToken,
      "channel": channelID,
      "text": message
    }
  };
  UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', message_options);
}

function convertJson(range, key) {
  const js = range.slice(1).map(
    function(row) {
      const obj = {};
      row.map(function(item, index) {
        obj[String(key[index])] = (!item) ? '' : item;
      });
      return obj;
    });
  return js;
}
