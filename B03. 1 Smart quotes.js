function replaceNonSmartWithSmartQuotes() {
  const ui = DocumentApp.getUi();
  try {
    const body = DocumentApp.getActiveDocument().getBody();

    if (body.getText().search(`"|'`) != -1) {
      body.replaceText(' "', ' “');
      body.replaceText('" ', '” ');
      body.replaceText(" '", " ‘");
      body.replaceText("' ", "’ ");

      if (body.getText().search(`"|'`) != -1) {

        const paragraphs = body.getParagraphs();
        if (body.getText().search('"') != -1) {
          for (let i in paragraphs) {
            paragraphs[i].replaceText('^"', '“');
            paragraphs[i].replaceText('"$', '”');
          }
        }

        if (body.getText().search("'") != -1) {
          for (let i in paragraphs) {
            paragraphs[i].replaceText("^'", "‘");
            paragraphs[i].replaceText("'$", "’");
          }
        }

        if (body.getText().search('\\("') != -1) {
          body.replaceText('\\("', '(“');
        }

        if (body.getText().search("\\('") != -1) {
          body.replaceText("\\('", "(‘");
        }

        if (body.getText().search('"') != -1) {
          body.replaceText('"', '”');
        }

        if (body.getText().search("'") != -1) {
          body.replaceText("'", "’");
        }
      }
    }
  }
  catch (error) {
    ui.alert('Error in replaceNonSmartWithSmartQuotes: ' + error);
    return 0;
  }
}