






function helpPopupUndo() {
  const ui = DocumentApp.getUi();
  ui.alert(`You may have noticed that BFormat actions cannot be undone. The reason for this is that we are using the Docs API. The reason for using the Docs API is that there are certain operations that are only possible with the Doc API. 
However, it does have the disadvantage that you cannot undo those functions. Use BFormat functions with care. If necessary you can revert to an earlier version of the document using the File > Version History.`);
}

function onOpen(e) {

  const thisDocStyle = getDefaultStyle();

  const subMenu = DocumentApp.getUi().createMenu('More styles');

  let selectedStyleMarker = '';
  for (let styleName in styles) {
    subMenu.addItem(styles[styleName]['name'] + selectedStyleMarker, styleName);
  }

  DocumentApp.getUi().createMenu('BFormat')
    .addItem('Why can I not use undo?', 'helpPopupUndo')
    .addItem('Help for setting default styles', 'setDefaultStylesManually')
    .addSeparator()
    .addItem('Use default: ' + styles[thisDocStyle]['name'], 'defaultStyleReport')
    .addSubMenu(subMenu)
    .addItem('Format text like Heading 1', 'formatTextLikeH1')
    .addItem('Reformat headings for tables, figures, boxes', 'reformatHeadings5and6') //TODO
    .addSeparator()

    .addItem('Insert box', 'insertBox')  //TODO
    .addItem('Format this box', 'formatBox') //TODO
    .addItem('Add right-border to paragraph', 'leftBorderParagraph')
    .addSeparator()

    .addItem('Insert table 2x2', 'insertTable2x2')
    .addItem('Insert table 3x3', 'insertTable3x3')
    .addItem('Insert table 4x4', 'insertTable4x4')
    // Elena: do not change bold:
    .addItem('Format this table', 'formatTableNoBold')
    // Function as before: everything bold.
    .addItem('Format this table (all bold)', 'formatTable')
    .addItem('Format this table (basic)', 'formatTableBasic')  //TODO
    .addSeparator()

    .addItem('Insert figure/image (style 1)', 'insertFigure1')
    .addItem('Insert figure/image (style 2)', 'insertFigure2')
    .addSeparator()

    .addItem('Insert pull quote', 'insertPullQuote')
    //.addItem('Format as pull quote', 'formatAsPullQuote') //TODO
    .addItem('Insert extracted quote', 'insertExtractedQuote') //TODO
    //.addItem('Format as extracted quote', 'formatAsExtractedQuote') //TODO
    .addSeparator()

    .addItem('Format lists', 'formatListsPart1')
    .addItem('Remove underline from hyperlinks', 'removeUnderlineFromHyperlinks')
    .addItem('Replace non-smart quotes with smart quotes', 'replaceNonSmartWithSmartQuotes') //TODO
    .addSeparator()

    .addItem('Format header', 'formatHeader')
    .addItem('Update footer', 'formatFooter')
    .addItem('Use default margins (Report)', 'defaultMargins')
    .addToUi();
}