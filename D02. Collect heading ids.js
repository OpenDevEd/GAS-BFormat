// Constants
const BROKEN_INTERNAL_LINK_MARKER = "<BROKEN_INTERNAL_LINK>";
const LINK_MARK_STYLE_NEW = new Object();
LINK_MARK_STYLE_NEW[DocumentApp.Attribute.FOREGROUND_COLOR] = "#ff0000";
LINK_MARK_STYLE_NEW[DocumentApp.Attribute.BACKGROUND_COLOR] = "#ffffff";
LINK_MARK_STYLE_NEW[DocumentApp.Attribute.BOLD] = true;

// UPDATED - Returns allLinks as well
function collectHeadingTexts(doc) {
  try {
    if (!doc) {
      doc = DocumentApp.getActiveDocument();
    }

    const documentId = doc.getId();
    const document = Docs.Documents.get(documentId);

    const bodyElements = document.body.content;

    const allLinks = [];
    const allHeadingParagraphs = new Object();
    const allHeadings = [];
    requests = [];
    for (let i in bodyElements) {
      // If body element contains table
      if (bodyElements[i].table) {
        if (bodyElements[i].table.tableRows) {
          for (let j in bodyElements[i].table.tableRows) {
            if (bodyElements[i].table.tableRows[j].tableCells) {
              for (let k in bodyElements[i].table.tableRows[j].tableCells) {
                if (bodyElements[i].table.tableRows[j].tableCells[k].content) {
                  for (let l in bodyElements[i].table.tableRows[j].tableCells[k].content) {
                    if (bodyElements[i].table.tableRows[j].tableCells[k].content[l].paragraph) {
                      collectHeadingTextsHelper(bodyElements[i].table.tableRows[j].tableCells[k].content[l].paragraph, allHeadingParagraphs, allHeadings, allLinks);
                    }
                  }
                }
              }
            }
          }
        }
      }
      // End. If body element contains table

      // If body element contains paragraph
      if (bodyElements[i].paragraph) {
        collectHeadingTextsHelper(bodyElements[i].paragraph, allHeadingParagraphs, allHeadings, allLinks);
      }
      // End. If body element contains paragraph
    }
    // Logger.log(allLinks);
    // Logger.log(allHeadingParagraphs);
    // Logger.log(allHeadings);
    return { status: 'ok', allHeadingsObj: allHeadingParagraphs, allHeadingsArray: allHeadings, allLinks: allLinks };
  }
  catch (error) {
    if (error.toString().includes('API call to docs.documents.get failed with error: Internal error encountered.')) {
      return { status: 'error', message: 'Sorry, bNumbers has encountered an error. It\'s a known error we haven\'t been able to trace yet. However, you can fix it by copying the contents of your document to a new document and running bNumbers in the new document. (Error in collectHeadingTexts: Error in collectHeadingTexts: ' + error + ')' };
    } else {
      return { status: 'error', message: 'Error in collectHeadingTexts: ' + error };
    }
  }
}

// Checks paragraphs in body and paragraphs in tables
function collectHeadingTextsHelper(paragraph, allHeadingParagraphs, allHeadings, allLinks) {
  let headingParagraph = false;
  let headingText, headingId;
  if (paragraph.paragraphStyle) {
    if (paragraph.paragraphStyle.headingId) {
      headingId = paragraph.paragraphStyle.headingId;
      allHeadingParagraphs[headingId] = '';
      allHeadings.push({ headingId: headingId, text: '' });
      allHeadingsLastEl = allHeadings.length - 1;
      headingParagraph = true;
    }
  }

  if (paragraph.elements) {
    paragraph.elements.forEach(function (item) {
      if (item.textRun) {
        // Collects links
        if (item.textRun.textStyle) {
          if (item.textRun.textStyle.link) {
            if (item.textRun.textStyle.link.headingId) {
              allLinks.push({ linkHeadingId: item.textRun.textStyle.link.headingId, content: item.textRun.content, startIndex: item.startIndex, endIndex: item.endIndex });
            }
          }
        }
        // End. Collects links

        if (item.textRun.content && headingParagraph) {
          headingText = item.textRun.content.replace('\n', '');
          allHeadingParagraphs[headingId] += headingText;
          allHeadings[allHeadingsLastEl]['text'] += headingText;
        }
      }
    });
  }
  // When you press enter at the start of an existing heading, headingId still exists, but the paragraph doesn't have content, Doc shows "Heading no longer exists"
  // "Heading no longer exists" case
  if (allHeadingParagraphs[headingId] == '' && headingParagraph) {
    delete allHeadingParagraphs[headingId];
    allHeadings.pop();
    allHeadingsLastEl--;
  }
  // End.  "Heading no longer exists" case
}