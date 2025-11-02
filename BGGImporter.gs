function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("BGG Logged Play Import")
    .addItem("Run Play Importer", "OpenBGGImportDialog")
    .addItem("Run Collection Importer ", "OpenBGGCollectionDialog")
    .addToUi();
}

function OpenBGGCollectionDialog() {
  var html = HtmlService.createTemplateFromFile("BGGCollectionImporterDialog");

  // Read the preferences from the last import
  var userProperties = PropertiesService.getUserProperties();
  html.usernameVar =
    userProperties.getProperty("BGGUsername") != null
      ? userProperties.getProperty("BGGUsername")
      : "";
  html.incExpansionsVar =
    userProperties.getProperty("BGGIncExpansions") != null &&
    userProperties.getProperty("BGGIncExpansions") == "true"
      ? "checked"
      : "";

  var htmlOutput = html.evaluate().setWidth(410).setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "BGG Collection Import");
}

function OpenBGGImportDialog() {
  var html = HtmlService.createTemplateFromFile("BGGImporterDialog");

  // Read the preferences from the last import
  var userProperties = PropertiesService.getUserProperties();
  html.usernameVar =
    userProperties.getProperty("BGGUsername") != null
      ? userProperties.getProperty("BGGUsername")
      : "";
  html.minDateVar =
    userProperties.getProperty("BGGMinDate") != null
      ? userProperties.getProperty("BGGMinDate")
      : "2025-01-01";
  html.maxDateVar =
    userProperties.getProperty("BGGMaxDate") != null
      ? userProperties.getProperty("BGGMaxDate")
      : "2025-12-31";
  html.incPlaytimeVar =
    userProperties.getProperty("BGGIncPlaytime") != null &&
    userProperties.getProperty("BGGIncPlaytime") == "true"
      ? "checked"
      : "";
  html.incGameYearVar =
    userProperties.getProperty("BGGIncGameYear") != null &&
    userProperties.getProperty("BGGIncGameYear") == "true"
      ? "checked"
      : "";
  html.incGameRankVar =
    userProperties.getProperty("BGGIncGameRank") != null &&
    userProperties.getProperty("BGGIncGameRank") == "true"
      ? "checked"
      : "";
  html.incInCollectionVar =
    userProperties.getProperty("BGGIncInCollection") != null &&
    userProperties.getProperty("BGGIncInCollection") == "true"
      ? "checked"
      : "";
  html.incNewGameVar =
    userProperties.getProperty("BGGNewGame") != null &&
    userProperties.getProperty("BGGNewGame") == "true"
      ? "checked"
      : "";
  html.incPlayersVar =
    userProperties.getProperty("BGGIncPlayers") != null &&
    userProperties.getProperty("BGGIncPlayers") == "true"
      ? "checked"
      : "";
  html.incPlayerScoresAndRatingsVar =
    userProperties.getProperty("BGGIncPlayerScoresAndRatings") != null &&
    userProperties.getProperty("BGGIncPlayerScoresAndRatings") == "true"
      ? "checked"
      : "";
  html.incCommentsVar =
    userProperties.getProperty("BGGIncComments") != null &&
    userProperties.getProperty("BGGIncComments") == "true"
      ? "checked"
      : "";

  var htmlOutput = html.evaluate().setWidth(410).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "BGG Play Stats Import");
}

function TestPlayImporter() {
  var username = "[USERNAME]";
  ImportBGGPlays(
    username,
    "2025-01-01",
    "2025-12-31",
    true,
    true,
    true,
    true,
    true,
    true,
    true
  );
}

function TestCollectionImporter() {
  var username = "[USERNAME]";
  var incExpansions = true;
  ImportBGGCollection(username, incExpansions);
}
