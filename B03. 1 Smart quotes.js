function replaceNonSmartWithSmartQuotes() {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    body.replaceText(' "', ' “');
    body.replaceText('" ', '” ');
    body.replaceText(" '", " ‘");
    body.replaceText("' ", "’ ");
 
    if (body.getText().search('"') != -1) {
      const paragraphs = body.getParagraphs();
      for (let i in paragraphs) {
        paragraphs[i].replaceText('^"', '“');
        paragraphs[i].replaceText('"$', '”');
      }
    }
  
    if (body.getText().search("'") != -1) {
      const paragraphs = body.getParagraphs();
      for (let i in paragraphs) {
        paragraphs[i].replaceText("^'", "‘");
        paragraphs[i].replaceText("'$", "’");
      }
    }
    
    if (body.getText().search('"') != -1) {
      replaceQuoteMark(body, /\W"/, '"', '“', true);
      replaceQuoteMark(body, /"\W/, '"', '”', false);
    }
    
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
  const symbols = ['(', ')'];
  let quoteMarkW, realW, replacement;
  quoteMarkW = pattern.exec(body.getText());
  while (quoteMarkW != null) {
    realW = quoteMarkW[0].replace(wrongQuoteMark, '');

    if (symbols.indexOf(realW) != -1) {
      quoteMarkW[0] = quoteMarkW[0].replace(realW, '\\' + realW);
    }
    if (beforeWord) {
      replacement = realW + correctQuoteMark;
    } else {
      replacement = correctQuoteMark + realW;
    }
    body.replaceText(quoteMarkW[0], replacement);
    quoteMarkW = pattern.exec(body.getText());
  }
}