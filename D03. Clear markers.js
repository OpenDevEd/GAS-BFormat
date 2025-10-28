function clearInternalLinkMarkers() {
  const ui = DocumentApp.getUi();
  try {
    const doc = DocumentApp.getActiveDocument();
    doc.replaceText(BROKEN_INTERNAL_LINK_MARKER, '');

    const footnotes = doc.getFootnotes();
    let footnote;
    for (let i in footnotes) {
      footnote = footnotes[i].getFootnoteContents();
      if (footnote == null) continue;
      footnote.replaceText(BROKEN_INTERNAL_LINK_MARKER, '');
    }
  }
  catch (error) {
    ui.alert('Error in clearInternalLinkMarkers: ' + error);
  }
}