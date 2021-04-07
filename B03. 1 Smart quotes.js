function replaceNonSmartWithSmartQuotes() {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    Logger.log(1);
    body.replaceText(' "', ' “');
    body.replaceText('" ', '” ');
    body.replaceText(" '", " ‘");
    body.replaceText("' ", "’ ");
    Logger.log(2);
    if (body.getText().search('"') != -1) {
      const paragraphs = body.getParagraphs();
      for (let i in paragraphs) {
        paragraphs[i].replaceText('^"', '“');
        paragraphs[i].replaceText('"$', '”');
      }
    }
    Logger.log(3);
    if (body.getText().search("'") != -1) {
      const paragraphs = body.getParagraphs();
      for (let i in paragraphs) {
        paragraphs[i].replaceText("^'", "‘");
        paragraphs[i].replaceText("'$", "’");
      }
    }
    Logger.log(4);
    if (body.getText().search('"') != -1) {
      replaceQuoteMark(body, /\W"/, '"', '“', true);
      replaceQuoteMark(body, /"\W/, '"', '”', false);
    }
    Logger.log(5);
    if (body.getText().search("'") != -1) {
      replaceQuoteMark(body, /\W'/, "'", "‘", true);
      replaceQuoteMark(body, /'\W/, "'", "’", false);
    }
  }
  catch (error) {
    Logger.log(error);
  }
}

function replaceQuoteMark(body, pattern, wrongQuoteMark, correctQuoteMark, beforeWord) {
  let symbols = ['(', ')'];
  let quoteMarkW, realW, replacement;
  quoteMarkW = pattern.exec(body.getText());
  while (quoteMarkW != null) {
    Logger.log(JSON.stringify(quoteMarkW));
    realW = quoteMarkW[0].replace(wrongQuoteMark, '');
    Logger.log(JSON.stringify(realW));
    Logger.log(realW);
    if (symbols.indexOf(realW) != -1) {
      quoteMarkW[0] = quoteMarkW[0].replace(realW, '\\' + realW);
    }
    if (beforeWord) {
      replacement = realW + correctQuoteMark;
    } else {
      replacement = correctQuoteMark + realW;
    }
    Logger.log('new');
    body.replaceText(quoteMarkW[0], replacement);
    quoteMarkW = pattern.exec(body.getText());
  }
}