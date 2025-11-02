function ImportBGGCollection(username, incExpansions) {
  SpreadsheetApp.getUi().alert("Collection Import starting");
  // Delete previous data from spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange("A2:A").clear();

  // Save the preferences
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperties({
    BGGUsername: username,
    BGGIncExpansions: incExpansions
  });

  var authorization = "Bearer [AUTH TOKEN]"; // Replace with your actual token
  var opt = {
    method: "get",
    headers: {
      authorization: authorization,
    },
  };

  if (incExpansions) {
    var collectionUrl = "https://www.boardgamegeek.com/xmlapi2/collection?own=1&username=" + username;
  } else {
    var collectionUrl = "https://www.boardgamegeek.com/xmlapi2/collection?own=1&excludesubtype=boardgameexpansion&username=" + username;
  }
  var collectionResponse = UrlFetchApp.fetch(collectionUrl, opt);
  var collectionXml = collectionResponse.getContentText();

  var collectionDoc = XmlService.parse(collectionXml);
  var collectionRoot = collectionDoc.getRootElement();

  while (collectionRoot.getName() != "items") {
    // Sleep to wait for BGG to produce the list
    Utilities.sleep(500);
    collectionResponse = UrlFetchApp.fetch(collectionUrl, opt);
    collectionXml = collectionResponse.getContentText();
    collectionDoc = XmlService.parse(collectionXml);
    collectionRoot = collectionDoc.getRootElement();
  }

  var collectionItems = collectionRoot.getChildren("item");

  var item, game;

  // Loop through each play in the file
  for (var i = 0; i < collectionItems.length; i++) {
    item = collectionItems[i];
    game = item.getChild("name").getValue().toString();

    sheet.getRange("A" + (i + 2)).setValue(game);
  }
  sheet.autoResizeColumns(1, 20);
}