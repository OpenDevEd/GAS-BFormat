// Sets 20 pt space after tables
function setSpaceAfterTables20pt(body = DocumentApp.getActiveDocument().getBody()) {
  let childIndexTable, parAfterTable, parText;
  const tables = body.getTables();
  for (let i = 1; i < tables.length; i++) {
    const parentType = tables[i].getParent().getType();
    if (parentType == 'BODY_SECTION') {
      //Logger.log('tables[i].getText()= %s', tables[i].getText());
      childIndexTable = body.getChildIndex(tables[i]);
      //Logger.log('childIndexTable= %s', childIndexTable);
      parAfterTable = body.getChild(childIndexTable + 1);
      if (parAfterTable.getType() == DocumentApp.ElementType.PARAGRAPH) {
        parText = parAfterTable.asText().getText();
        if (parText != '') {
          body.insertParagraph(childIndexTable + 1, '').setHeading(DocumentApp.ParagraphHeading.NORMAL).setSpacingAfter(8);
        } else {
          parAfterTable.asParagraph().setHeading(DocumentApp.ParagraphHeading.NORMAL).setSpacingAfter(8);
        }
      }
    }
  }
}


function setSpaceBeforeAfterInTable(body = DocumentApp.getActiveDocument().getBody()) {
  const tables = body.getTables();
  for (var i = 0; i < tables.length; i++) {
    var table = tables[i];
    numRows = table.getNumRows();
    for (var j = 0; j < numRows; j++) {
      row = table.getRow(j);
      numCells = row.getNumCells();
      for (var k = 0; k < numCells; k++) {
        var cell = row.getCell(k);
        var psn = cell.getNumChildren();
        for (var l = 0; l < psn; l++) {
          //paras.push(cell.getChild(i));
          childElement = cell.getChild(l);
          if (childElement.getType() == DocumentApp.ElementType.PARAGRAPH) {
            childElement.asParagraph().setSpacingBefore(5);
            childElement.asParagraph().setSpacingAfter(5);
          }

          // body.insertParagraph(index, cell.getChild(i).copy())
        };
      }
    }
  }
}
