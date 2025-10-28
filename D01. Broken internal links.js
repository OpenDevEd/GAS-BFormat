// New optimised version
function findAndMarkBrokenLinksOptimised() {
  const doc = DocumentApp.getActiveDocument();
  const ui = DocumentApp.getUi();

  try {
    // STEP 1: Collect all internal links first
    const allInternalLinks = {
      headingLinks: [],
      bookmarkLinks: [],
      tabLinks: []
    };

    // Scan body
    collectInternalLinks(doc.getBody(), allInternalLinks);

    // Scan footnotes
    const footnotes = doc.getFootnotes();
    for (let i = 0; i < footnotes.length; i++) {
      const footnoteContents = footnotes[i].getFootnoteContents();
      if (footnoteContents == null) {
        continue;
      }
      const numChildren = footnoteContents.getNumChildren();
      for (let j = 0; j < numChildren; j++) {
        collectInternalLinks(footnoteContents.getChild(j), allInternalLinks);
      }
    }

    const totalInternalLinks = allInternalLinks.headingLinks.length +
      allInternalLinks.bookmarkLinks.length +
      allInternalLinks.tabLinks.length;

    // Logger.log('Found internal links:');
    // Logger.log('  Heading links: ' + allInternalLinks.headingLinks.length);
    // Logger.log('  Bookmark links: ' + allInternalLinks.bookmarkLinks.length);
    // Logger.log('  Tab links: ' + allInternalLinks.tabLinks.length);

    // If no internal links found, exit early
    if (totalInternalLinks === 0) {
      ui.alert('No internal links found in the document.');
      return;
    }

    // STEP 2: Validate only the types of links that exist
    let validHeadingIds = [];
    let validBookmarkIds = [];
    let validTabIds = [];

    // Collect valid heading IDs only if heading links exist
    if (allInternalLinks.headingLinks.length > 0) {
      // Logger.log('Collecting valid heading IDs...');
      const headingData = collectHeadingTexts(doc);

      if (headingData.status === 'error') {
        ui.alert('Error: ' + headingData.message);
        return;
      }

      validHeadingIds = Object.keys(headingData.allHeadingsObj);
      // Logger.log('Valid headings found: ' + validHeadingIds.length);
    }

    // Collect valid bookmark IDs only if bookmark links exist
    if (allInternalLinks.bookmarkLinks.length > 0) {
      // Logger.log('Collecting valid bookmark IDs...');
      const bookmarks = doc.getBookmarks();
      validBookmarkIds = bookmarks.map(function (bookmark) {
        return bookmark.getId();
      });
      // Logger.log('Valid bookmarks found: ' + validBookmarkIds.length);
    }

    // Collect valid tab IDs only if tab links exist
    if (allInternalLinks.tabLinks.length > 0) {
      // Logger.log('Collecting valid tab IDs...');
      const tabs = doc.getTabs();
      validTabIds = tabs.map(function (tab) {
        return tab.getId();
      });
      // Logger.log('Valid tabs found: ' + validTabIds.length);
    }

    // STEP 3: Find broken links
    const brokenLinks = [];

    // Check heading links
    allInternalLinks.headingLinks.forEach(function (link) {
      if (validHeadingIds.indexOf(link.id) === -1) {
        link.type = 'heading';
        brokenLinks.push(link);
      }
    });

    // Check bookmark links
    allInternalLinks.bookmarkLinks.forEach(function (link) {
      if (validBookmarkIds.indexOf(link.id) === -1) {
        link.type = 'bookmark';
        brokenLinks.push(link);
      }
    });

    // Check tab links
    allInternalLinks.tabLinks.forEach(function (link) {
      if (validTabIds.indexOf(link.id) === -1) {
        link.type = 'tab';
        brokenLinks.push(link);
      }
    });

    // Logger.log('Total broken links found: ' + brokenLinks.length);

    // STEP 4: Mark broken links
    if (brokenLinks.length > 0) {
      // Sort in descending order to avoid index shifting
      brokenLinks.sort(function (a, b) {
        return b.offset - a.offset;
      });

      brokenLinks.forEach(function (link) {
        const element = link.element;
        const offset = link.offset;

        element.insertText(offset, BROKEN_INTERNAL_LINK_MARKER)
          .setLinkUrl(offset, offset + BROKEN_INTERNAL_LINK_MARKER.length - 1, null)
          .setAttributes(offset, offset + BROKEN_INTERNAL_LINK_MARKER.length - 1, LINK_MARK_STYLE_NEW);
      });
    }

    // STEP 5: Build and show alert
    let infoLinks = '';
    const brokenByType = {
      heading: [],
      bookmark: [],
      tab: []
    };

    brokenLinks.forEach(function (link) {
      brokenByType[link.type].push(link);
    });

    if (brokenByType.heading.length > 0) {
      infoLinks += 'Broken heading links: ' + brokenByType.heading.length + '\n';
      brokenByType.heading.forEach(function (link, index) {
        infoLinks += '  ' + (index + 1) + '. "' + link.text + '" (ID: ' + link.id + ')\n';
      });
      infoLinks += '\n';
    }

    if (brokenByType.bookmark.length > 0) {
      infoLinks += 'Broken bookmark links: ' + brokenByType.bookmark.length + '\n';
      brokenByType.bookmark.forEach(function (link, index) {
        infoLinks += '  ' + (index + 1) + '. "' + link.text + '" (ID: ' + link.id + ')\n';
      });
      infoLinks += '\n';
    }

    if (brokenByType.tab.length > 0) {
      infoLinks += 'Broken tab links: ' + brokenByType.tab.length + '\n';
      brokenByType.tab.forEach(function (link, index) {
        infoLinks += '  ' + (index + 1) + '. "' + link.text + '" (ID: ' + link.id + ')\n';
      });
      infoLinks += '\n';
    }

    if (brokenLinks.length > 0) {
      infoLinks = 'There were broken links. Please search for ' + BROKEN_INTERNAL_LINK_MARKER + ' and fix the links. Note that broken links can occur when you press enter at the start of an existing heading.\n\n' + infoLinks;
      ui.alert(infoLinks);
      return { status: 'error' };
    } else {
      return { status: 'ok' };
    }
  } catch (error) {
    ui.alert('Error in findAndMarkBrokenLinksOptimised: ' + error);
    // Logger.log('Error: ' + error);
    return { status: 'error' };
  }
}

// Collect all internal links from an element
function collectInternalLinks(element, allInternalLinks) {
  const elementType = String(element.getType());

  if (elementType === 'TEXT') {
    const textElement = element.asText();
    const text = textElement.getText();
    const indices = textElement.getTextAttributeIndices();

    for (let i = 0; i < indices.length; i++) {
      const index = indices[i];
      const partAttributes = textElement.getAttributes(index);

      if (partAttributes.LINK_URL) {
        const url = partAttributes.LINK_URL;
        const endIndex = (i < indices.length - 1) ? indices[i + 1] : text.length;
        const linkText = text.substring(index, endIndex);

        // Check for heading link
        const headingMatch = url.match(/#heading=([^&]+)/);
        if (headingMatch) {
          allInternalLinks.headingLinks.push({
            text: linkText,
            url: url,
            id: headingMatch[1],
            element: textElement,
            offset: index
          });
          continue;
        }

        // Check for bookmark link
        const bookmarkMatch = url.match(/#bookmark=([^&]+)/);
        if (bookmarkMatch) {
          allInternalLinks.bookmarkLinks.push({
            text: linkText,
            url: url,
            id: bookmarkMatch[1],
            element: textElement,
            offset: index
          });
          continue;
        }

        // Check for tab link (format: ?tab=t.xxx or &tab=t.xxx)
        const tabMatch = url.match(/[?&]tab=(t\.[^&#]+)/);
        if (tabMatch) {
          allInternalLinks.tabLinks.push({
            text: linkText,
            url: url,
            id: tabMatch[1],
            element: textElement,
            offset: index
          });
        }
      }
    }
  } else {
    // Recursively process child elements
    const arrayTypes = ['BODY_SECTION', 'PARAGRAPH', 'LIST_ITEM', 'TABLE', 'TABLE_ROW', 'TABLE_CELL'];
    if (arrayTypes.indexOf(elementType) !== -1) {
      const numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        collectInternalLinks(element.getChild(i), allInternalLinks);
      }
    }
  }
}