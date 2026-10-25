// Uncomment the menu item for each script imported into this project.

function onOpen() {
  DocumentApp.getUi()
    .createMenu("Document Utilities")

    // .addItem("Count words by heading", "showHeadingWordCounts")
    // .addItem("Tidy APA statistical notation", "tidyApaStatNotation")
    // .addItem("Highlight references…", "showHighlightOptions")

    .addToUi();
}
