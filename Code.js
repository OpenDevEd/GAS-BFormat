function onOpen(e) {
  DocumentApp.getUi().createMenu('BFormat')
     //.createAddOnMenu('BFormat')
    .addItem('Format text like Heading 1', 'formatTextLikeH1')
    .addSeparator()

    .addItem('Insert table 2x2', 'insertTable2x2')
    .addItem('Insert table 3x3', 'insertTable3x3')
    .addItem('Insert table4x4', 'insertTable4x4')
    .addItem('Format table', 'formatTable')
    .addSeparator()

    .addItem('Insert figure (style 1)', 'insertFigure1')
    .addItem('Insert figure (style 2)', 'insertFigure2')
    .addSeparator()

    .addItem('Insert pull quote', 'insertPullQuote')
    .addSeparator()
  
    .addItem('Format lists', 'formatListsPart1')
    .addItem('Remove underline from hyperlinks', 'removeUnderlineFromHyperlinks')
    .addSeparator()
  
    .addSubMenu(DocumentApp.getUi().createMenu('Advanced functions')
      .addItem('Set default styles (manually)', 'setDefaultStylesManually')
      .addItem('Use default styles (Report)', 'defaultStyleReport')
      .addSeparator()
      .addItem('Format header', 'formatHeader')
      .addItem('Update footer', 'formatFooter')
      .addItem('Use default margins (Report)', 'defaultMargins')
    )
    .addToUi();
}