function ImportBGGPlays(
  username,
  minDate,
  maxDate,
  includePlaytime,
  includeGameYear,
  includeGameRank,
  includeInCollection,
  includeNewGame,
  includePlayers,
  includePlayerScoresAndRatings,
  includeComments
) {
  //SpreadsheetApp.getUi().alert('Play Import starting');

  // Delete previous data from spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clearContents();

  // Save the preferences
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperties({
    BGGUsername: username,
    BGGMinDate: minDate,
    BGGMaxDate: maxDate,
    BGGIncPlaytime: includePlaytime,
    BGGIncGameYear: includeGameYear,
    BGGIncGameRank: includeGameRank,
    BGGIncInCollection: includeInCollection,
    BGGNewGame: includeNewGame,
    BGGIncPlayers: includePlayers,
    BGGIncPlayerScoresAndRatings: includePlayerScoresAndRatings,
    BGGIncComments: includeComments,
  });

  var authorization = "Bearer [AUTH TOKEN]"; // Replace with your actual token
  
  var opt = {
    method: "get",
    headers: {
      authorization: authorization,
    },
  };
  var collectionUrl =
    "https://www.boardgamegeek.com/xmlapi2/collection?own=1&username=" +
    username +
    "&showprivate=1";

  var collectionResponse = UrlFetchApp.fetch(collectionUrl, opt);
  var collectionXml = collectionResponse.getContentText();

  var collectionDoc = XmlService.parse(collectionXml);
  var collectionRoot = collectionDoc.getRootElement();

  while (collectionRoot.getName() != "items") {
    // Sleep to wait for BGG to produce the list
    Utilities.sleep(15000);
    collectionResponse = UrlFetchApp.fetch(collectionUrl, opt);
    collectionXml = collectionResponse.getContentText();
    collectionDoc = XmlService.parse(collectionXml);
    collectionRoot = collectionDoc.getRootElement();
    //SpreadsheetApp.getUi().alert('Waiting on BGG Collection creation');
  }

  //Logger.log(collectionXml);

  var collectionItems = collectionRoot.getChildren();

  var exCollectionItems;

  // First call to the API to establish how many plays there are, and therefore the number of pages
  var url =
    "http://www.boardgamegeek.com/xmlapi2/plays?username=" +
    username +
    "&mindate=" +
    minDate +
    "&maxdate=" +
    maxDate +
    "&page=1";
  var xml = UrlFetchApp.fetch(url, opt).getContentText();
  var document = XmlService.parse(xml);
  var root = document.getRootElement();

  var totalNumberOfPlays = 0 + root.getAttribute("total").getValue();
  var totalNumberOfPages = Math.ceil(totalNumberOfPlays / 100);

  writeHeaderRow(
    sheet,
    includePlaytime,
    includeGameYear,
    includeGameRank,
    includeInCollection,
    includeNewGame,
    includePlayers,
    includePlayerScoresAndRatings,
    includeComments
  );

  var entries, item, players;
  var date,
    game,
    location,
    timesPlayed,
    playtime,
    isInCollection,
    isNew,
    numberOfPlayers,
    player1,
    player1Score,
    player1Rating,
    player2,
    player2Score,
    player2Rating,
    player3,
    player3Score,
    player3Rating,
    player4,
    player4Score,
    player4Rating,
    player5,
    player5Score,
    player5Rating,
    player6,
    player6Score,
    player6Rating,
    player7,
    player7Score,
    player7Rating,
    player8,
    player8Score,
    player8Rating,
    winner,
    comments;

  // Loop through the pages
  for (var h = 1; h <= totalNumberOfPages; h++) {
    url =
      "http://www.boardgamegeek.com/xmlapi2/plays?username=" +
      username +
      "&mindate=" +
      minDate +
      "&maxdate=" +
      maxDate +
      "&page=" +
      h;
    xml = UrlFetchApp.fetch(url, opt).getContentText();
    //Logger.log(xml);
    document = XmlService.parse(xml);
    root = document.getRootElement();
    entries = root.getChildren();

    // Loop through each play in the file
    for (var i = 0; i < entries.length; i++) {
      date = entries[i].getAttribute("date").getValue();
      location = entries[i].getAttribute("location").getValue();
      timesPlayed = entries[i].getAttribute("quantity").getValue();
      item = entries[i].getChildren("item")[0];
      game = item.getAttribute("name").getValue();
      var thisRow = [date, game, location, timesPlayed];

      if (includePlaytime) {
        playtime = entries[i].getAttribute('length').getValue();
        playtime = playtime != "0"
                    ? playtime
                    : ""
        thisRow.push(playtime);
      }

      var gameId = item.getAttribute("objectid").getValue();
      if (includeGameYear || includeGameRank) {
        var gameYearAndRank = getGameYearAndRank(gameId, opt);
        if (includeGameYear) {
          thisRow.push(gameYearAndRank[0]);
        }
        if (includeGameRank) {
          thisRow.push(gameYearAndRank[1]);
        }
      }

      if (includeInCollection) {
        isInCollection = getIsInCollection(collectionItems, gameId);
        thisRow.push(isInCollection);
      }

      isNew = "No";
      if (entries[i].getChild("players") != null) {
        players = entries[i].getChild("players").getChildren();

        if (includeNewGame) {
          isNew = getIsNew(players, username);
          thisRow.push(isNew);
        }

        numberOfPlayers = players.length;
        thisRow.push(numberOfPlayers);
        if (includePlayers) {
          player1 =
            players[0] != null
              ? players[0].getAttribute("name").getValue()
              : "";
          player2 =
            players[1] != null
              ? players[1].getAttribute("name").getValue()
              : "";
          player3 =
            players[2] != null
              ? players[2].getAttribute("name").getValue()
              : "";
          player4 =
            players[3] != null
              ? players[3].getAttribute("name").getValue()
              : "";
          player5 =
            players[4] != null
              ? players[4].getAttribute("name").getValue()
              : "";
          player6 =
            players[5] != null
              ? players[5].getAttribute("name").getValue()
              : "";
          player7 =
            players[6] != null
              ? players[6].getAttribute("name").getValue()
              : "";
          player8 =
            players[7] != null
              ? players[7].getAttribute("name").getValue()
              : "";

          if (includePlayerScoresAndRatings) {
            player1Score =
              players[0] != null
                ? players[0].getAttribute("score").getValue()
                : "";
            player1Rating =
              players[0] != null
                ? players[0].getAttribute("rating").getValue() != "0"
                  ? players[0].getAttribute("rating").getValue()
                  : ""
                : "";
            player2Score =
              players[1] != null
                ? players[1].getAttribute("score").getValue()
                : "";
            player2Rating =
              players[1] != null
                ? players[1].getAttribute("rating").getValue() != "0"
                  ? players[1].getAttribute("rating").getValue()
                  : ""
                : "";
            player3Score =
              players[2] != null
                ? players[2].getAttribute("score").getValue()
                : "";
            player3Rating =
              players[2] != null
                ? players[2].getAttribute("rating").getValue() != "0"
                  ? players[2].getAttribute("rating").getValue()
                  : ""
                : "";
            player4Score =
              players[3] != null
                ? players[3].getAttribute("score").getValue()
                : "";
            player4Rating =
              players[3] != null
                ? players[3].getAttribute("rating").getValue() != "0"
                  ? players[3].getAttribute("rating").getValue()
                  : ""
                : "";
            player5Score =
              players[4] != null
                ? players[4].getAttribute("score").getValue()
                : "";
            player5Rating =
              players[4] != null
                ? players[4].getAttribute("rating").getValue() != "0"
                  ? players[4].getAttribute("rating").getValue()
                  : ""
                : "";
            player6Score =
              players[5] != null
                ? players[5].getAttribute("score").getValue()
                : "";
            player6Rating =
              players[5] != null
                ? players[5].getAttribute("rating").getValue() != "0"
                  ? players[5].getAttribute("rating").getValue()
                  : ""
                : "";
            player7Score =
              players[6] != null
                ? players[6].getAttribute("score").getValue()
                : "";
            player7Rating =
              players[6] != null
                ? players[6].getAttribute("rating").getValue() != "0"
                  ? players[6].getAttribute("rating").getValue()
                  : ""
                : "";
            player8Score =
              players[7] != null
                ? players[7].getAttribute("score").getValue()
                : "";
            player8Rating =
              players[7] != null
                ? players[7].getAttribute("rating").getValue() != "0"
                  ? players[7].getAttribute("rating").getValue()
                  : ""
                : "";
          } else {
            player1Score = "";
            player1Rating = "";
            player2Score = "";
            player2Rating = "";
            player3Score = "";
            player3Rating = "";
            player4Score = "";
            player4Rating = "";
            player5Score = "";
            player5Rating = "";
            player6Score = "";
            player6Rating = "";
            player7Score = "";
            player7Rating = "";
            player8Score = "";
            player8Rating = "";
          }

          winner = getWinner(players);
        } else {
          player1 = "";
          player1Score = "";
          player1Rating = "";
          player2 = "";
          player2Score = "";
          player2Rating = "";
          player3 = "";
          player3Score = "";
          player3Rating = "";
          player4 = "";
          player4Score = "";
          player4Rating = "";
          player5 = "";
          player5Score = "";
          player5Rating = "";
          player6 = "";
          player6Score = "";
          player6Rating = "";
          player7 = "";
          player7Score = "";
          player7Rating = "";
          player8 = "";
          player8Score = "";
          player8Rating = "";
          winner = "";
        }
        if (includePlayerScoresAndRatings) {
          thisRow.push(
            player1,
            player1Score,
            player1Rating,
            player2,
            player2Score,
            player2Rating,
            player3,
            player3Score,
            player3Rating,
            player4,
            player4Score,
            player4Rating
          );
          thisRow.push(
            player5,
            player5Score,
            player5Rating,
            player6,
            player6Score,
            player6Rating,
            player7,
            player7Score,
            player7Rating,
            player8,
            player8Score,
            player8Rating,
            winner
          );
        } else if (includePlayers) {
          thisRow.push(
            player1,
            player2,
            player3,
            player4,
            player5,
            player6,
            player7,
            player8,
            winner
          );
        } else {
          // do nothing
        }
      } else {
        thisRow.push(isNew);
      }

      if (includeComments) {
        if (entries[i].getChildren("comments").length > 0) {
          comments = entries[i].getChildren("comments")[0].getText();
        } else {
          comments = "";
        }
        thisRow.push(comments);
      }

      sheet.appendRow(thisRow);
    }
  }
  sheet.autoResizeColumns(1, 10);
}

function writeHeaderRow(
  sheet,
  includePlaytime,
  includeGameYear,
  includeGameRank,
  includeInCollection,
  includeNewGame,
  includePlayers,
  includePlayerScoresAndRatings,
  includeComments
) {
  var row = ["Date", "Game", "Location", "Times Played"];
  if (includePlaytime) {
    row.push("Playtime");
  }
  if (includeGameYear) {
    row.push("Year Published");
  }
  if (includeGameRank) {
    row.push("Current BGG Rank");
  }
  if (includeInCollection) {
    row.push("Is In Collection");
  }
  if (includeNewGame) {
    row.push("Is New Game");
  }
  row.push("Number of Players");
  if (includePlayers) {
    if (includePlayerScoresAndRatings) {
      row.push(
        "Player 1",
        "Player 1 Score",
        "Player 1 Rating",
        "Player 2",
        "Player 2 Score",
        "Player 2 Rating",
        "Player 3",
        "Player 3 Score",
        "Player 3 Rating"
      );
      row.push(
        "Player 4",
        "Player 4 Score",
        "Player 4 Rating",
        "Player 5",
        "Player 5 Score",
        "Player 5 Rating",
        "Player 6",
        "Player 6 Score",
        "Player 6 Rating"
      );
      row.push(
        "Player 7",
        "Player 7 Score",
        "Player 7 Rating",
        "Player 8",
        "Player 8 Score",
        "Player 8 Rating",
        "Winner"
      );
    } else {
      row.push(
        "Player 1",
        "Player 2",
        "Player 3",
        "Player 4",
        "Player 5",
        "Player 6",
        "Player 7",
        "Player 8",
        "Winner"
      );
    }
  }
  if (includeComments) {
    row.push("Comments");
  }

  sheet.appendRow(row);
}

function getWinner(players) {
  var winner = "";
  if (players[0] != null && players[0].getAttribute("win").getValue() == "1") {
    winner = players[0].getAttribute("name").getValue();
  } else if (
    players[1] != null &&
    players[1].getAttribute("win").getValue() == "1"
  ) {
    winner = players[1].getAttribute("name").getValue();
  } else if (
    players[2] != null &&
    players[2].getAttribute("win").getValue() == "1"
  ) {
    winner = players[2].getAttribute("name").getValue();
  } else if (
    players[3] != null &&
    players[3].getAttribute("win").getValue() == "1"
  ) {
    winner = players[3].getAttribute("name").getValue();
  } else if (
    players[4] != null &&
    players[4].getAttribute("win").getValue() == "1"
  ) {
    winner = players[4].getAttribute("name").getValue();
  } else if (
    players[5] != null &&
    players[5].getAttribute("win").getValue() == "1"
  ) {
    winner = players[5].getAttribute("name").getValue();
  } else if (
    players[6] != null &&
    players[6].getAttribute("win").getValue() == "1"
  ) {
    winner = players[6].getAttribute("name").getValue();
  } else if (
    players[7] != null &&
    players[7].getAttribute("win").getValue() == "1"
  ) {
    winner = players[7].getAttribute("name").getValue();
  }
  return winner;
}

function getIsNew(players, username) {
  var isNew = "No";
  for (var i = 0; i < players.length; i++) {
    if (players[i].getAttribute("username").getValue() == username) {
      if (players[i].getAttribute("new").getValue() == "1") {
        isNew = "Yes";
      }
    }
  }
  return isNew;
}

function getIsInCollection(collectionItems, gameId) {
  var isFound = false;
  var i = 0;
  var isInCollection = "No";
  while (!isFound && i < collectionItems.length) {
    if (collectionItems[i].getAttribute("objectid").getValue() == gameId) {
      isInCollection = "Yes";
      isFound = true;
    }
    i++;
  }
  return isInCollection;
}

function getGameYearAndRank(gameId, opt) {
  try {
    var url = "http://www.boardgamegeek.com/xmlapi2/thing?stats=1&id=" + gameId;
    var xml = UrlFetchApp.fetch(url, opt).getContentText();
    var document = XmlService.parse(xml);
    var root = document.getRootElement();
    var entry = root.getChild("item");

    var year = entry.getChild("yearpublished").getAttribute("value").getValue();

    var ranks = entry
      .getChild("statistics")
      .getChild("ratings")
      .getChild("ranks")
      .getChildren("rank");
    var rank = "";

    for (var i = 0; i < ranks.length; i++) {
      var thisRank = ranks[i];
      if (thisRank.getAttribute("id").getValue() == "1") {
        rank = thisRank.getAttribute("value").getValue();
      }
    }

    return [year, rank];
  } catch (e) {
    return ["0", "0"];
  }
}