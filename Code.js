function onOpen(e) {
  DocumentApp.getUi().createMenu('BFormat')
    .addItem('Help for setting default styles', 'setDefaultStylesManually')
    .addItem('Use default styles (for Report)', 'defaultStyleReport')
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
    .addItem('Format this table', 'formatTable')
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

function reformatHeadings5and6Example() {
  stylesetTable = {
    run_in_regexp: /^Table (\d+)\./,
    run_in_style: {
// Bold, orange
    },
    follow_on_style: {
// italics, orange
    }
  }
  stylesetFigure = {
    run_in_regexp: /^(Figure|Box) (\d+)\./,
    run_in_style: {
// Bold black
    },
    follow_on_style: {
// italics, black      
    }
  }
  reformatHeadingWithRunInStyle('Heading 5', stylesetTable)
  reformatHeadingWithRunInStyle('Heading 6', stylesetFigure)

}