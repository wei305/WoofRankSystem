function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

function refreshRanking() {
  var url = "*****";
  var spread_sheet = SpreadsheetApp.openByUrl(url);
  var ws = spread_sheet.getSheetByName("May");
  var raw_data = createDataWithNames()
  Logger.log(raw_data)
  
  lastCol = ws.getLastColumn()
  for (var col = 2; col<lastCol+1; col++){
    header = ws.getRange(1, col).getValue()
    if (header == "總分"){
      break
    }
    game = readOldGame(col)
    if (game["win"].length == 0){
      continue
    }

    win_rank = 0
    lose_rank = 0

    for (var i=0; i<game["win"].length; i++){
      win_rank += raw_data[game["win"][i]]["ranking"]
    }

    for (var i=0; i<game["mvp"].length; i++){
      win_rank += raw_data[game["mvp"][i]]["ranking"]
    }

    win_ave_rank = win_rank/(game["win"].length+1)

    for (var i=0; i<game["lose"].length; i++){
      lose_rank += raw_data[game["lose"][i]]["ranking"]
    }

    lose_ave_rank = lose_rank/(game["lose"].length)

    win_ratio = 1 + (lose_ave_rank-win_ave_rank)/150

    for (var i=0; i<game["win"].length; i++){
      raw_data[game["win"][i]]["ranking"] += Math.round(win_ratio*20)
      raw_data[game["win"][i]]["win"] += 1
      raw_data[game["win"][i]]["games"] += 1
      raw_data[game["win"][i]]["win%"] = raw_data[game["win"][i]]["win"]/raw_data[game["win"][i]]["games"]
      raw_data[game["win"][i]]["mvp%"] = raw_data[game["win"][i]]["mvp"]/raw_data[game["win"][i]]["games"]
    }

    for (var i=0; i<game["mvp"].length; i++){
      raw_data[game["mvp"][i]]["ranking"] += Math.round(win_ratio*30)
      raw_data[game["mvp"][i]]["win"] += 1
      raw_data[game["mvp"][i]]["games"] += 1
      raw_data[game["mvp"][i]]["mvp"] += 1
      raw_data[game["mvp"][i]]["win%"] = raw_data[game["mvp"][i]]["win"]/raw_data[game["mvp"][i]]["games"]
      raw_data[game["mvp"][i]]["mvp%"] = raw_data[game["mvp"][i]]["mvp"]/raw_data[game["mvp"][i]]["games"]
    }

    for (var i=0; i<game["lose"].length; i++){
      raw_data[game["lose"][i]]["ranking"] -= Math.round(win_ratio*10)
      raw_data[game["lose"][i]]["lose"] += 1
      raw_data[game["lose"][i]]["games"] += 1
      raw_data[game["lose"][i]]["win%"] = raw_data[game["lose"][i]]["win"]/raw_data[game["lose"][i]]["games"]
      raw_data[game["lose"][i]]["mvp%"] = raw_data[game["lose"][i]]["mvp"]/raw_data[game["lose"][i]]["games"]
    }

    for (var i=0; i<game["host"].length; i++){
      raw_data[game["host"][i]]["ranking"] += 10
      raw_data[game["host"][i]]["host"] += 1
    }
    Logger.log(game)
    Logger.log(raw_data)
  }

  var keys = Object.keys(raw_data);

  data_array = []

  for(var i = 0; i < keys.length;i++){
    raw_data[keys[i]]["name"] = keys[i]
    data_array.push(raw_data[keys[i]])
  }

  data_array.sort((a, b) => (a.ranking < b.ranking) ? 1 : -1)

  Logger.log(data_array)

  return data_array
}

function createDataWithNames() {
  var url = "*****";
  var spread_sheet = SpreadsheetApp.openByUrl(url);
  var ws = spread_sheet.getSheetByName("May");
  var data = {}

  // add names
  last_row = ws.getLastRow()
  for (var row = 2; row<last_row+1; row++){
    data[ws.getRange(row, 1).getValue()] = {"ranking": 1200, "win": 0, "lose": 0, "games": 0, "mvp": 0, "host": 0, "win%": 0, "mvp%": 0}
  }

  return data
}

function readOldGame(col) {
  var url = "*****";
  var spread_sheet = SpreadsheetApp.openByUrl(url);
  var ws = spread_sheet.getSheetByName("May");

  last_row = ws.getLastRow()
  ret = {'win': [], 'lose': [], 'mvp' : [], 'host': []}
  for (var row=2; row<last_row+1; row++){
    score = ws.getRange(row, col).getValue()
    name = ws.getRange(row, 1).getValue()

    if (score == -1) {
      ret["lose"].push(name)
    }
    else if (score == 1){
      ret["host"].push(name)
    }
    else if (score == 2){
      ret["win"].push(name)
    }
    else if (score == 3){
      ret["mvp"].push(name)
    }
  }
  return ret
}