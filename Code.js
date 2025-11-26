// When the add-on is open, the JSON file calls this function, which creates the menu/user interface.
function onHomePage() {
  
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty('numClicksForbiddenSymbolsBody', '0');
  properties.setProperty('numClicksFormatBody', '1')
  properties.setProperty('numClicksFormatFootnotes', '0')
  
  // Create an OpenLink object for the user guide
  const userGuideLink = CardService.newOpenLink()
    .setUrl('https://docs.google.com/document/d/13m9SrdctCp64Gv8I2NrBVf0TXYfiKUwMNDmQLPftdEw/edit')
    .setOpenAs(CardService.OpenAs.FULL_SIZE);

  // Create a button that links to the user guide
  const userGuideButton = CardService.newTextButton()
    .setText('Click here')
    .setOpenLink(userGuideLink);

  var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Convert your doc to HTML!"))
    
      // User guide
      .addSection(CardService.newCardSection()
          .setHeader("User guide")
          .addWidget(userGuideButton))
      // 1. Create a backup
      .addSection(CardService.newCardSection()
          .setHeader("1. Make a copy of your doc")
          .addWidget(CardService.newTextParagraph().setText("File \u2192 Make a copy \u2192 Copy comments and suggestions")))
          // the stuff below is deprecated, we no longer use the createBackup() function.
          //.addWidget(CardService.newTextButton()
              //.setText("Create backup")
              //.setOnClickAction(CardService.newAction()
                  //.setFunctionName("createBackup")))
      // 2. Replace forbidden symbols
      .addSection(CardService.newCardSection()
          .setHeader("2. Replace forbidden symbols")
          .addWidget(CardService.newTextParagraph().setText("Use Find and Replace to substitute <, ≤, >, and ≥ with their corresponding HTML tags.")))
      // 3. Complete manual edits
      .addSection(CardService.newCardSection()
          .setHeader("3. Complete manual edits")
          .addWidget(CardService.newTextParagraph().setText("Complete all HTML coding that is not handled by the app. This includes summary boxes, hover-overs, source table, table of contents, and links to non-headers (if applicable).")))
      // 4. Modify document body
      .addSection(CardService.newCardSection()
          .setHeader("4. Modify document body")
          .addWidget(CardService.newTextButton()
              .setText("Get body data")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("bodyData")))
          // DEPRECATED, EASIER TO DO THIS MANUALLY
          //.addWidget(CardService.newTextButton()
              //.setText("Replace forbidden symbols")
              //.setOnClickAction(CardService.newAction()
                  //.setFunctionName("replaceForbiddenSymbolsInBody")))
          .addWidget(CardService.newTextButton()
              .setText("Format rich text")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("formatBody"))))
      // 5. Modify and move footnotes
      .addSection(CardService.newCardSection()
          .setHeader("5. Modify and move footnotes")
          .addWidget(CardService.newTextButton()
              .setText("Get footnote data")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("footnoteData")))
          // DEPRECATED, EASIER TO DO THIS MANUALLY
          //.addWidget(CardService.newTextButton()
              //.setText("Replace forbidden symbols")
              //.setOnClickAction(CardService.newAction()
                  //.setFunctionName("replaceForbiddenSymbolsInFootnotes")))
          .addWidget(CardService.newTextButton()
              .setText("Format rich text")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("formatFootnotes")))
          .addWidget(CardService.newTextButton()
              .setText("Convert bulleted lists")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("formatFootnoteBulletedLists")))
          .addWidget(CardService.newTextButton()
              .setText("Convert numbered lists")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("formatFootnoteNumberedLists")))
          .addWidget(CardService.newTextButton()
              .setText("Move footnotes to body")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("moveFootnotes"))))
      // 6. Lists
      .addSection(CardService.newCardSection()
          .setHeader("6. Convert bulleted and numbered lists")
          .addWidget(CardService.newTextButton()
              .setText("Convert bulleted lists")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("formatBulletedLists")))
          .addWidget(CardService.newTextButton()
              .setText("Convert numbered lists")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("formatNumberedLists"))))
      
      // 7. Headers
      .addSection(CardService.newCardSection()
          .setHeader("7. Format headers")
          .addWidget(CardService.newTextButton()
              .setText("Change headers to HTML")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("convertHeadersToHTML"))))
    
      // 8. Tables
      .addSection(CardService.newCardSection()
          .setHeader("8. Convert tables")
          .addWidget(CardService.newTextButton()
              .setText("Format rich text")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("formatTables")))
          .addWidget(CardService.newTextButton()
              .setText("Format bulleted lists")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("formatTableBulletedLists")))
          .addWidget(CardService.newTextButton()
              .setText("Format numbered lists")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("formatTableNumberedLists")))
          .addWidget(CardService.newTextButton()
              .setText("Convert tables to HTML code")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("findAndReplaceTables"))))
      
      // 9. Manually edit table formatting
      .addSection(CardService.newCardSection()
          .setHeader("9. Manually edit table formatting")
          .addWidget(CardService.newTextParagraph().setText("Follow the instructions in the user guide to manually add gray-highlighted or header rows.")))
      // 10. Format GiveWell links
      .addSection(CardService.newCardSection()
        .setHeader("10. Format GiveWell links")
        .addWidget(CardService.newTextButton()
          .setText("Remove 'https://www.givewell.org'")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("formatGiveWellLinks"))))
      // Version log
      .addSection(CardService.newCardSection()
          .setHeader("Version info")
          .addWidget(CardService.newTextParagraph().setText("Version 2.3, deployment 226, 1/31/24")))

      .build();
  return [card];

}

// Function that tells the user how many times they'll need to process body features.
function bodyData() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var elementCount = body.getNumChildren();

  // Use Ui service to display a dialog box
  var ui = DocumentApp.getUi();
  ui.alert("Element Count", "The body of the document contains " + elementCount + " elements. The code processes 25 elements at a time, so you'll need to use the 'Format rich text' function " + Math.ceil(elementCount/25) + " times.", ui.ButtonSet.OK);
}

// Function that tells the user how many times they'll need to process footnotes.
function footnoteData() {
  var doc = DocumentApp.getActiveDocument();
  var footnotes = doc.getFootnotes();
  var footnoteCount = footnotes.length;

  // Use Ui service to display a dialog box
  var ui = DocumentApp.getUi();
  ui.alert("Footnote Count", "The document contains " + footnoteCount + " footnotes. The code processes 10 footnotes at a time, so you'll need to use the 'Format rich text' function " + Math.ceil(footnoteCount/10) + " times.", ui.ButtonSet.OK);
}

// DEPRECATED Function to create a backup copy of a Google Doc
function createBackup() {
  var doc = DocumentApp.getActiveDocument();
  var docId = doc.getId();
  
  // Set the title for the copied file.
  var copiedFile = Drive.Files.copy({title: doc.getName() + " - Backup"}, docId);
}

// DEPRECATED function that replaces <, >, ≤, and ≥ in the body. Turns out that Find and Replace is easier.
function replaceForbiddenSymbolsInBody() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();  // Fetch the document body
  
  // tracking number of clicks
  var properties = PropertiesService.getScriptProperties();
  var currentCount = parseInt(properties.getProperty('numClicksForbiddenSymbolsBody'), 10);
  currentCount +=1;
  properties.setProperty('numClicksForbiddenSymbolsBody',currentCount.toString());

  // Define the range of elements to process
  var startElement = (currentCount-1)*25;
  var endElement = currentCount * 25;

  //Ensure endElement and startElement don't exceed the number of elements in the body
  var numElements = body.getNumChildren();
  
  if (startElement >= numElements) {
    return
  }
  
  if (endElement > numElements) {
    endElement = numElements;
  }

  for (var i = startElement; i < endElement; i++) {
    var child = body.getChild(i);
    
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH ||
        child.getType() === DocumentApp.ElementType.LIST_ITEM) { // Handle both paragraphs and list items
      var text;
      if (child.getNumChildren() > 0) {
        text = child.getChild(0).asText();
      } else {
        continue; // skip the current iteration and proceed to the next one
      }

      var len = text.getText().length;  // Initial length
      
      for (var j = 0; j < len; j++) {
        var char = text.getText()[j];  // Get character
        var isBold = text.isBold(j);
        var isItalic = text.isItalic(j);
        var isUnderline = text.isUnderline(j);
        var linkUrl = text.getLinkUrl(j);  // Get hyperlink URL if it exists
        
        // Replace forbidden characters and retain formatting
        // Repeated block for each forbidden character
        ['<', '>', '≤', '≥'].forEach(function(forbiddenChar) {
          var replacement;
          switch (forbiddenChar) {
            case '<':
              replacement = '&lt;';
              break;
            case '>':
              replacement = '&gt;';
              break;
            case '≤':
              replacement = '&le;';
              break;
            case '≥':
              replacement = '&ge;';
              break;
          }
          if (char === forbiddenChar) {
            text.deleteText(j, j);
            text.insertText(j, replacement);
            text.setBold(j, j + 3, isBold);
            text.setItalic(j, j + 3, isItalic);
            text.setUnderline(j, j + 3, isUnderline);
            if (linkUrl) {
              text.setLinkUrl(j, j + 3, linkUrl);  // Reapply hyperlink
            }
            len += 3;  // Adjust length
            j += 3;  // Adjust index
          }
        });
      }
    }
  }
}

// formats rich text in the body, 25 elements at a time.
function formatBody() {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    
    // Tracking number of clicks
    var properties = PropertiesService.getScriptProperties();
    var currentCount = parseInt(properties.getProperty('numClicksFormatBody'), 10) || 0;

    // Define the range of elements to process
    var startElement = (currentCount - 1) * 25;
    var endElement = currentCount * 25;
    var numElements = body.getNumChildren();
    
    if (startElement >= numElements) {
        return;
    }
    
    if (endElement > numElements) {
        endElement = numElements;
    }

    for (var i = startElement; i < endElement; i++) {
        var child = body.getChild(i);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH || child.getType() === DocumentApp.ElementType.LIST_ITEM) {
            removeUnderlinesFromLinks(child);
            processLinks(child);
            processFormatting(child);
        }
    }
  currentCount += 1;
  properties.setProperty('numClicksFormatBody', currentCount.toString());  
}

// helper function to formatBody(), removes underlines from hyperlinks to avoid getting confused about underlined vs. hyperlinked text
function removeUnderlinesFromLinks(child) {
    var text = child.asText();
    var len = text.getText().length;

    for (var m = 0; m < len; m++) {
        if (text.getLinkUrl(m)) {
            text.setUnderline(m, m, false);
        }
    }
}

// helper function to formatBody(), formats hyperlinks
function processLinks(child) {
    var text = child.asText();
    var len = text.getText().length;
    var m = 0;
    var currentLinkUrl = "";
    var modifications = [];

    while (m < len) {
        var charLink = text.getLinkUrl(m);
        if (charLink && currentLinkUrl === "") {
            currentLinkUrl = charLink;
            modifications.push({ index: m, insertText: `<a href="${currentLinkUrl}">` });
            m += 1;
        } else if (charLink && charLink === currentLinkUrl) {
            m++;
        } else if (currentLinkUrl !== "") {
            modifications.push({ index: m, insertText: "</a>", clearLink: true });
            currentLinkUrl = "";
        } else {
            m++;
        }
    }

    // If a link was left open, close it
    if (currentLinkUrl !== "") {
        modifications.push({ index: m, insertText: "</a>", clearLink: true });
    }

    // Apply modifications in reverse order
    for (var i = modifications.length - 1; i >= 0; i--) {
        var mod = modifications[i];
        text.insertText(mod.index, mod.insertText);
        if (mod.clearLink) {
            text.setLinkUrl(mod.index, mod.index, null);
        }
    }
}

// helper function to formatBody(), formats all types of rich text besides hyperlinks
function processFormatting(child) {
    var text = child.asText();
    var len = text.getText().length;
    var formats = {
        Bold: { startTag: "<strong>", endTag: "</strong>", remover: "setBold" },
        Italic: { startTag: "<em>", endTag: "</em>", remover: "setItalic" },
        Underline: { startTag: "<u>", endTag: "</u>", remover: "setUnderline" }
    };

    for (var format in formats) {
        var j = 0;
        while (j < len) {
            if (text[`is${format}`](j)) {
                var formatStart = j;
                while (j < len && text[`is${format}`](j)) j++;
                var formatEnd = j - 1;
                text.insertText(formatEnd + 1, formats[format].endTag);
                text.insertText(formatStart, formats[format].startTag);
                text[formats[format].remover](formatStart, formatEnd + formats[format].startTag.length + formats[format].endTag.length, false); // Remove the format
                j += formats[format].startTag.length + formats[format].endTag.length; // Jump past the inserted tags
                len += formats[format].startTag.length + formats[format].endTag.length; // Adjust length for added tags
            } else {
                j++;
            }
        }
    }
}

// DEPRECATED function that replaces <, >, ≤, and ≥ in the footnotes. Turns out that Find and Replace is easier.
function replaceForbiddenSymbolsInFootnotes() {
  var doc = DocumentApp.getActiveDocument();
  var footnotes = doc.getFootnotes();  // Fetch all footnotes in the document
  var firstTenFootnotes = footnotes.slice(0, 10); // get first 10 footnotes

  for (var i = 0; i < firstTenFootnotes.length; i++) {
    var content = firstTenFootnotes[i].getFootnoteContents();  // Get content

    for (var k = 0; k < content.getNumChildren(); k++) {
      var child = content.getChild(k);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH ||
          child.getType() === DocumentApp.ElementType.LIST_ITEM) { // handle both paragraphs and list items
        
        var text;
        if (child.getNumChildren() > 0) {
          text = child.getChild(0).asText();
        } else {
          continue; // skip the current iteration and proceed to the next one
        }

        var len = text.getText().length;  // Initial length
        
        for (var j = 0; j < len; j++) {
          var char = text.getText()[j];  // Get character
          var isBold = text.isBold(j);
          var isItalic = text.isItalic(j);
          var isUnderline = text.isUnderline(j);
          var linkUrl = text.getLinkUrl(j);  // Get hyperlink URL if it exists
          
          if (char === '<') {  // Check if forbidden symbol
            text.deleteText(j, j);
            text.insertText(j, '&lt;');
            text.setBold(j, j + 3, isBold);
            text.setItalic(j, j + 3, isItalic);
            text.setUnderline(j, j + 3, isUnderline);
            if (linkUrl) {
              text.setLinkUrl(j, j + 3, linkUrl);  // Reapply hyperlink
            }
            len += 3;  // Adjust length
            j += 3;  // Adjust index
          }
          if (char === '>') {  // Check if forbidden symbol
            text.deleteText(j, j);
            text.insertText(j, '&gt;');
            text.setBold(j, j + 3, isBold);
            text.setItalic(j, j + 3, isItalic);
            text.setUnderline(j, j + 3, isUnderline);
            if (linkUrl) {
              text.setLinkUrl(j, j + 3, linkUrl);  // Reapply hyperlink
            }
            len += 3;  // Adjust length
            j += 3;  // Adjust index
          }
          if (char === '≤') {  // Check if forbidden symbol
            text.deleteText(j, j);
            text.insertText(j, '&le;');
            text.setBold(j, j + 3, isBold);
            text.setItalic(j, j + 3, isItalic);
            text.setUnderline(j, j + 3, isUnderline);
            if (linkUrl) {
              text.setLinkUrl(j, j + 3, linkUrl);  // Reapply hyperlink
            }
            len += 3;  // Adjust length
            j += 3;  // Adjust index
          }
          if (char === '≥') {  // Check if forbidden symbol
            text.deleteText(j, j);
            text.insertText(j, '&ge;');
            text.setBold(j, j + 3, isBold);
            text.setItalic(j, j + 3, isItalic);
            text.setUnderline(j, j + 3, isUnderline);
            if (linkUrl) {
              text.setLinkUrl(j, j + 3, linkUrl);  // Reapply hyperlink
            }
            len += 3;  // Adjust length
            j += 3;  // Adjust index
          }
        }
      }
    }
  }
}

// formats rich text in footnotes
function formatFootnotes() {
  var doc = DocumentApp.getActiveDocument();
  var footnotes = doc.getFootnotes();  // Fetch all footnotes in the document
  var firstTenFootnotes = footnotes.slice(0, 10); // I'm pretty sure this is deprecated since we're tracking with button clicks now.
  
  // Tracking number of clicks
  var properties = PropertiesService.getScriptProperties();
  var currentCount = parseInt(properties.getProperty('numClicksFormatFootnotes'), 10) || 0;
  currentCount += 1;
  properties.setProperty('numClicksFormatFootnotes', currentCount.toString());

  // Define the range of elements to process
  var startFootnote = (currentCount - 1) * 10;
  var endFootnote = currentCount * 10;
  var numFootnotes = footnotes.length;
    
  if (startFootnote >= numFootnotes) {
      return;
  }
    
  if (endFootnote > numFootnotes) {
      endFootnote = numFootnotes;
  }

  for (var i = startFootnote; i < endFootnote; i++) {
    var content = footnotes[i].getFootnoteContents();  // Get content
    
    for (var k = 0; k < content.getNumChildren(); k++) {
      var child = content.getChild(k);
      
      if (child.getType() == DocumentApp.ElementType.PARAGRAPH ||
          child.getType() === DocumentApp.ElementType.LIST_ITEM) { // Handle both paragraphs and list items
        
        var text;
        if (child.getNumChildren() > 0) {
          text = child.getChild(0).asText();
        } else {
          continue; // skip the current iteration and proceed to the next one
        }


        var len = text.getText().length;  // Initial length
        var newText = "";
        var linkText = "";
        var linkUrl = "";
        var openTags = [];
        
        // Remove underline from linked text
        for (var m = 0; m < len; m++) {
          if (text.getLinkUrl(m)) {
            text.setUnderline(m, m, false);
          }
        }
        
        for (var j = 0; j < len; j++) {
          var char = text.getText()[j];  // Get character
          
          // Fetch formatting for the current character
          var isBold = text.isBold(j);
          var isItalic = text.isItalic(j);
          var isUnderline = text.isUnderline(j);
          linkUrl = text.getLinkUrl(j);  // Get hyperlink URL if it exists
          
          if (linkUrl) {
            if (linkText === "") {
              openTags.push(`<a href="${linkUrl}">`);
            }
            linkText += char;
          } else {
            if (linkText !== "") {
              newText += `${openTags.pop()}${linkText}</a>`;
              linkText = "";
            }
          }
          
          if (isBold && !openTags.includes("<strong>")) {
            newText += "<strong>";
            openTags.push("<strong>");
          } else if (!isBold && openTags.includes("<strong>")) {
            newText += "</strong>";
            openTags.pop();
          }
          
          if (isItalic && !openTags.includes("<em>")) {
            newText += "<em>";
            openTags.push("<em>");
          } else if (!isItalic && openTags.includes("<em>")) {
            newText += "</em>";
            openTags.pop();
          }
          
          if (isUnderline && !openTags.includes("<u>")) {
            newText += "<u>";
            openTags.push("<u>");
          } else if (!isUnderline && openTags.includes("<u>")) {
            newText += "</u>";
            openTags.pop();
          }
          
          if (!linkUrl) {
            newText += char;
          }
        }
        
        if (linkText !== "") {
          newText += `${openTags.pop()}${linkText}</a>`;
          linkText = "";
        }


        // Close any remaining tags
        for (var tag of openTags.reverse()) {
          newText += tag.replace("<", "</");
        }
        
        // Replace the original text with the newly formatted text
        text.setText(newText);

        if (newText.length > 0) {
        text.setBold(0, newText.length - 1, false);
        text.setItalic(0, newText.length - 1, false);
        text.setUnderline(0, newText.length - 1, false);
        }
      }
    }
  }
}

// function to convert footnotes into text objects with reference to parent paragraphs.
function moveFootnotes() {
  var doc = DocumentApp.getActiveDocument();
  var footnotes = doc.getFootnotes(); // get the footnotes
  var firstTenFootnotes = footnotes.slice(0, 10); // We now process all footnotes at the same time because we don't worry about rich text. If you'd like to revert this, just change later references to "footnotes" to "firstTenFootnotes". This function has one place to change.
  footnotes.forEach(function (note) {
    // Traverse the child elements to get to the `Text` object
    // and make a deep copy

    var paragraph = note.getParent(); // get the paragraph
    var noteIndex = paragraph.getChildIndex(note); // get the footnote's "child index"
    insertFootnote(note.getFootnoteContents(),true, paragraph, noteIndex);
    note.removeFromParent();
  })
} 

// function to insert footnotes after they've been converted. Works.
function insertFootnote(note, recurse, paragraph, noteIndex){
  var numC = note.getNumChildren(); //find the # of children
  paragraph.insertText(noteIndex,"<fn>\n\n");
  noteIndex++;
  for (var i=0; i<numC; i++){
    var C;
    if (note.getChild(i).getNumChildren() > 0) {
      C = note.getChild(i).getChild(0).copy();
    } else {
      continue; // skip the current iteration and proceed to the next one
    }

    if (i == 0 && C.getText().trim() === "") {
      continue;
    }

    if (i==0){
      var temp = C.getText();
      if (temp[0] ===" "){
        C = C.deleteText(0,0);
      }
    }

    if (i>0){
      paragraph.insertText(noteIndex,"\n");
      noteIndex++;
    }
    paragraph.insertText(noteIndex,C);
    noteIndex++;

  } //end of looping through children
  paragraph.insertText(noteIndex,"\n\n</fn>");
}

// Function to convert headers in a Google Doc to HTML format
function convertHeadersToHTML() {
  var doc = DocumentApp.getActiveDocument();
  // Access the body of the document
  var body = doc.getBody();

  // Find all the paragraphs
  var paragraphs = body.getParagraphs();

  for (var i = 0; i < paragraphs.length; i++) {
    var headerText, htmlHeaderText;

    if (paragraphs[i].getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
      headerText = paragraphs[i].getText();
      htmlHeaderText = '<h2>' + headerText + '</h2>';
      paragraphs[i].setText(htmlHeaderText);
      paragraphs[i].setHeading(DocumentApp.ParagraphHeading.NORMAL);
    }

    if (paragraphs[i].getHeading() === DocumentApp.ParagraphHeading.HEADING2) {
      headerText = paragraphs[i].getText();
      htmlHeaderText = '<h3>' + headerText + '</h3>';
      paragraphs[i].setText(htmlHeaderText);
      paragraphs[i].setHeading(DocumentApp.ParagraphHeading.NORMAL);
    }

    if (paragraphs[i].getHeading() === DocumentApp.ParagraphHeading.HEADING3) {
      headerText = paragraphs[i].getText();
      htmlHeaderText = '<h4>' + headerText + '</h4>';
      paragraphs[i].setText(htmlHeaderText);
      paragraphs[i].setHeading(DocumentApp.ParagraphHeading.NORMAL);
    }

    if (paragraphs[i].getHeading() === DocumentApp.ParagraphHeading.HEADING4) {
      headerText = paragraphs[i].getText();
      htmlHeaderText = '<h5>' + headerText + '</h5>';
      paragraphs[i].setText(htmlHeaderText);
      paragraphs[i].setHeading(DocumentApp.ParagraphHeading.NORMAL);
    }
  }
}

// Function to convert unordered lists to HTML formatting
function formatBulletedLists() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var numElements = body.getNumChildren();
  var currentNestLevel = -1;
  var output = "";
  var listStartIndex = null;
  var replacements = [];

  for (var i = 0; i < numElements; i++) {
    var para = body.getChild(i);
    
    if (para.getType() === DocumentApp.ElementType.LIST_ITEM) {
      var listItem = para.asListItem();
      var glyphType = listItem.getGlyphType();
      
      if (listStartIndex === null) {
        listStartIndex = i;
      }
      
      if (glyphType === DocumentApp.GlyphType.BULLET ||
          glyphType === DocumentApp.GlyphType.HOLLOW_BULLET ||
          glyphType === DocumentApp.GlyphType.SQUARE_BULLET) {
        
        var nestLevel = listItem.getNestingLevel();
        while (nestLevel > currentNestLevel) {
          output += "<ul>\n";
          currentNestLevel++;
        }
        
        while (nestLevel < currentNestLevel) {
          output += "</ul>\n";
          currentNestLevel--;
        }
        
        output += "<li>" + para.getText() + "\n";
      } else {
        listStartIndex = null;
        output = "";
        currentNestLevel = -1;
      }
    } else {
      if (listStartIndex !== null && output.includes("<li>")) {
        if (currentNestLevel == 0) {
          output += "</ul>\n";
        }
        if (currentNestLevel > 0) {
          output += "</ul>\n</ul>\n";
        }
        
        replacements.push({start: listStartIndex, end: i - 1, text: output});
        listStartIndex = null;
        //console.log(output);
        output = "";
        currentNestLevel = -1;
      }
    }
  }
  
  if (listStartIndex !== null && output.includes("<li>")) { // I think this has something to do with the fact that output is empty at this point?
    if (currentNestLevel == 0) {
        output += "</ul>";
      }
    if (currentNestLevel > 0) {
        output += "</ul>\n</ul>\n";
      }
    
    replacements.push({start: listStartIndex, end: numElements - 1, text: output});
  }
  
  // Process the replacements in reverse to avoid index shifting issues
  for (var i = replacements.length - 1; i >= 0; i--) {
    var rep = replacements[i];
    
    for (var j = rep.end; j >= rep.start; j--) {
      
      // Additional logging to check the type and content of the element being removed
      var elementToRemove = body.getChild(j);
      
      body.removeChild(elementToRemove);
    }

    body.insertParagraph(rep.start, rep.text);
  }
}

// Function to convert bulleted lists in footnotes to HTML formatting
function formatFootnoteBulletedLists() {
  var doc = DocumentApp.getActiveDocument();
  var footnotes = doc.getFootnotes();
  var firstTenFootnotes = footnotes.slice(0, 10); // We now process all footnotes at the same time because we don't worry about rich text. If you'd like to revert this, just change later references to "footnotes" to "firstTenFootnotes". This function has three places to change.
  var replacements = [];

  for (var i = 0; i < footnotes.length; i++) {
    var footnote = footnotes[i];
    var footnoteContent = footnote.getFootnoteContents();
    var children = footnoteContent.getNumChildren();

    var currentNestLevel = -1;
    var output = "";
    var listStartIndex = null;

    for (var j = 0; j < children; j++) {
      var child = footnoteContent.getChild(j);

      if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
        var listItem = child.asListItem();
        var glyphType = listItem.getGlyphType();
      
        if (listStartIndex === null) {
          listStartIndex = j;
        }
        
        if (glyphType === DocumentApp.GlyphType.BULLET ||
            glyphType === DocumentApp.GlyphType.HOLLOW_BULLET ||
            glyphType === DocumentApp.GlyphType.SQUARE_BULLET) {
          
          var nestLevel = listItem.getNestingLevel();
          while (nestLevel > currentNestLevel) {
            output += "<ul>\n";
            currentNestLevel++;
          }
          
          while (nestLevel < currentNestLevel) {
            output += "</ul>\n";
            currentNestLevel--;
          }
          
          output += "<li>" + child.asText().getText() + "\n";
        } else {
          listStartIndex = null;
          output = "";
          currentNestLevel = -1;
        }
      } else {
        if (listStartIndex !== null && output.includes("<li>")) {
          if (currentNestLevel == 0) {
            output += "</ul>\n";
          }
          if (currentNestLevel > 0) {
            output += "</ul>\n</ul>\n";
          }
          //console.log(output);
          replacements.push({footnoteIndex: i, start: listStartIndex, end: j - 1, text: output});
          
          listStartIndex = null;
          output = "";
          currentNestLevel = -1;
        }
      }
    }
    
    if (listStartIndex !== null && output.includes("<li>")) {
      if (currentNestLevel == 0) {
        output += "</ul>";
      }
      if (currentNestLevel > 0) {
        output += "</ul>\n</ul>\n";
      }
  
      var replacementObj = {footnoteIndex: i, start: listStartIndex, end: children - 1, text: output};
      replacements.push(replacementObj);

      if (replacementObj.end === (footnoteContent.getNumChildren() - 1)) {
      footnoteContent.appendParagraph('PLACEHOLDER');
      }
    }
  }

  for (var i = replacements.length - 1; i >= 0; i--) {
    var rep = replacements[i];
    var footnote = footnotes[rep.footnoteIndex];
    var footnoteContent = footnote.getFootnoteContents();

    for (var j = rep.start; j <= rep.end; j++) {
      footnoteContent.removeChild(footnoteContent.getChild(rep.start));
    }

    footnoteContent.insertParagraph(rep.start, rep.text);

    // Remove the placeholder paragraph if it exists
    var lastChildIndex = footnoteContent.getNumChildren() - 1;
    if (lastChildIndex >= 0 && footnoteContent.getChild(lastChildIndex).getText() === 'PLACEHOLDER') {
      footnoteContent.getChild(lastChildIndex).asText().setText(""); //change the placeholder text to empty string
    }
  }
}

function formatNumberedLists() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var numElements = body.getNumChildren();
  var output = "";
  var listStartIndex = null;
  var replacements = [];

  for (var i = 0; i < numElements; i++) {
    var para = body.getChild(i);
    
    if (para.getType() === DocumentApp.ElementType.LIST_ITEM) {
      var listItem = para.asListItem();
      var glyphType = listItem.getGlyphType();
      
      if (glyphType === DocumentApp.GlyphType.NUMBER) {
        
        // Start the numbered list if it hasn't started yet
        if (listStartIndex === null) {
          listStartIndex = i;
          output = "<ol>\n";  // Start the HTML for the ordered list
        }
        
        output += "<li>" + para.getText() + "\n";
      } else {
        // Reset variables if a non-numbered list item is found
        listStartIndex = null;
        output = "";
      }
    } else {
      // This is a non-list item paragraph
      if (listStartIndex !== null && output.includes("<li>")) {
        output += "</ol>\n";
        replacements.push({start: listStartIndex, end: i - 1, text: output});
        listStartIndex = null;
        output = "";
      }
    }
  }
  
  // Handle the case where the document ends with a list
  if (listStartIndex !== null && output.includes("<li>")) {
    output += "</ol>\n";
    replacements.push({start: listStartIndex, end: numElements - 1, text: output});
  }
  
  // Process the replacements in reverse to avoid index shifting issues
  for (var i = replacements.length - 1; i >= 0; i--) {
    var rep = replacements[i];
    
    for (var j = rep.end; j >= rep.start; j--) {
      body.removeChild(body.getChild(j));
    }

    body.insertParagraph(rep.start, rep.text);
  }
}

function formatFootnoteNumberedLists() {
  var doc = DocumentApp.getActiveDocument();
  var footnotes = doc.getFootnotes();
  var firstTenFootnotes = footnotes.slice(0, 10); // We now process all footnotes at the same time because we don't worry about rich text. If you'd like to revert this, just change later references to "footnotes" to "firstTenFootnotes". This function has three places to change.
  var replacements = [];

  for (var i = 0; i < footnotes.length; i++) {
    var footnote = footnotes[i];
    var footnoteContent = footnote.getFootnoteContents();
    var children = footnoteContent.getNumChildren();

    var currentNestLevel = -1;
    var output = "";
    var listStartIndex = null;

    for (var j = 0; j < children; j++) {
      var child = footnoteContent.getChild(j);

      if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
        var listItem = child.asListItem();
        var glyphType = listItem.getGlyphType();
      
        if (listStartIndex === null) {
          listStartIndex = j;
        }
        
        if (glyphType === DocumentApp.GlyphType.NUMBER) { // Modified this line
          
          var nestLevel = listItem.getNestingLevel();
          while (nestLevel > currentNestLevel) {
            output += "<ol>\n"; // Modified this line
            currentNestLevel++;
          }
          
          while (nestLevel < currentNestLevel) {
            output += "</ol>\n"; // Modified this line
            currentNestLevel--;
          }
          
          output += "<li>" + child.asText().getText() + "\n";
        } else {
          listStartIndex = null;
          output = "";
          currentNestLevel = -1;
        }
      } else {
        if (listStartIndex !== null && output.includes("<li>")) {
          if (currentNestLevel !== -1) {
            output += "</ol>\n"; // Modified this line
          }
          
          replacements.push({footnoteIndex: i, start: listStartIndex, end: j - 1, text: output});
          
          listStartIndex = null;
          output = "";
          currentNestLevel = -1;
        }
      }
    }
    
    if (listStartIndex !== null && output.includes("<li>")) {
      if (currentNestLevel !== -1) {
        output += "</ol>"; // Modified this line
      }
  
      var replacementObj = {footnoteIndex: i, start: listStartIndex, end: children - 1, text: output};
      replacements.push(replacementObj);

      if (replacementObj.end === (footnoteContent.getNumChildren() - 1)) {
      footnoteContent.appendParagraph('PLACEHOLDER');
      }
    }
  }

  for (var i = replacements.length - 1; i >= 0; i--) {
    var rep = replacements[i];
    var footnote = footnotes[rep.footnoteIndex];
    var footnoteContent = footnote.getFootnoteContents();

    for (var j = rep.start; j <= rep.end; j++) {
      footnoteContent.removeChild(footnoteContent.getChild(rep.start));
    }

    footnoteContent.insertParagraph(rep.start, rep.text);

    // Remove the placeholder paragraph if it exists
    var lastChildIndex = footnoteContent.getNumChildren() - 1;
    if (lastChildIndex >= 0 && footnoteContent.getChild(lastChildIndex).getText() === 'PLACEHOLDER') {
      footnoteContent.getChild(lastChildIndex).asText().setText("");
    }
  }
}

// Function that finds and replaces tables
function findAndReplaceTables() {
    const Docs = DocumentApp.getActiveDocument();
    const body = Docs.getBody();
    const tables = body.getTables();

    // Go in reverse order to ensure that we don't change indices of subsequent tables
    for (let i = tables.length - 1; i >= 0; i--) {
        const table = tables[i];
        const html = convertTableToHTML(table);

        const index = body.getChildIndex(table);
        
        body.insertParagraph(index, html);  // Use insertParagraph here
        table.removeFromParent();
    }
}

// Helper function to actually perform the HTML conversion
function convertTableToHTML(table) {
    let html = '<table>\n';

    for (let rowIndex = 0; rowIndex < table.getNumRows(); rowIndex++) {
        let row = table.getRow(rowIndex);
        html += '<tr>';

        for (let cellIndex = 0; cellIndex < row.getNumCells(); cellIndex++) {
            let cell = row.getCell(cellIndex);
            let cellContent = "";

            for (let elementIndex = 0; elementIndex < cell.getNumChildren(); elementIndex++) {
                let element = cell.getChild(elementIndex);
                cellContent += getContentFromElement(element);
            }

            if (rowIndex == 0) {  // Header row
                html += `<th><strong>${cellContent}</strong></th>`;
            } else {
                html += `<td>${cellContent}</td>`;
            }
        }
        if (rowIndex == 0) {
            html += '</tr>\n';  // Add a newline after each row, but no extra <tr> after the header row>
        } else {
            html += '</tr>\n<tr>\n';  // Add a newline after each row, plus an extra <tr> tag
        }
    }

    html += '</table>';
    return html;
}

//Helper function to get the content from each table cell.
function getContentFromElement(element) {
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
        return element.getText().trim();
    } else if (element.getType() === DocumentApp.ElementType.LIST_ITEM) {
        return element.getText().trim();
    }
    return "";  // Default case
}

// This function finds PARAGRAPH and LIST_ITEM elements in tables and then calls the existing functions to format rich text.
function formatTables() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const tables = body.getTables();

    tables.forEach(table => {
        const numRows = table.getNumRows();
        for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
            const row = table.getRow(rowIndex);
            const numCells = row.getNumCells();

            for (let cellIndex = 0; cellIndex < numCells; cellIndex++) {
                const cell = row.getCell(cellIndex);
                const numChildren = cell.getNumChildren();

                for (let childIndex = 0; childIndex < numChildren; childIndex++) {
                    const child = cell.getChild(childIndex);
                    if (child.getType() === DocumentApp.ElementType.PARAGRAPH || 
                        child.getType() === DocumentApp.ElementType.LIST_ITEM) {
                        removeUnderlinesFromLinks(child);
                        processLinks(child);
                        processFormatting(child);
                    }
                }
            }
        }
    });
}

// this is the main function that formats bulleted lists in tables.
function formatTableBulletedLists() {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var tables = body.getTables();
    
    tables.forEach(table => {
        const numRows = table.getNumRows();
        for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
            const row = table.getRow(rowIndex);
            const numCells = row.getNumCells();

            for (let cellIndex = 0; cellIndex < numCells; cellIndex++) {
                const cell = row.getCell(cellIndex);
                
                formatCellBulletedLists(cell);
            }
        }
    });
}

// this is the helper function that replaces lists on a cell level.
function formatCellBulletedLists(cell) {
    var numElements = cell.getNumChildren();
    var currentNestLevel = -1;
    var output = "";
    var listStartIndex = null;
    var replacements = [];

    for (var i = 0; i < numElements; i++) {
        var para = cell.getChild(i);
        
        if (para.getType() === DocumentApp.ElementType.LIST_ITEM) {
            var listItem = para.asListItem();
            var glyphType = listItem.getGlyphType();
            
            if (listStartIndex === null) {
                listStartIndex = i;
            }
            
            if (glyphType === DocumentApp.GlyphType.BULLET ||
                glyphType === DocumentApp.GlyphType.HOLLOW_BULLET ||
                glyphType === DocumentApp.GlyphType.SQUARE_BULLET) {

                var nestLevel = listItem.getNestingLevel();

                while (nestLevel > currentNestLevel) {
                    output += "<ul>\n";
                    currentNestLevel++;
                }
                
                while (nestLevel < currentNestLevel) {
                    output += "</ul>\n";
                    currentNestLevel--;
                }
                
                output += "<li>" + para.getText() + "\n";
            } else {
                listStartIndex = null;
                output = "";
                currentNestLevel = -1;
            }
        } else {
            if (listStartIndex !== null && output.includes("<li>")) {
                if (currentNestLevel !== -1) {
                    output += "</ul>\n";
                }
                
                replacements.push({start: listStartIndex, end: i - 1, text: output});
                listStartIndex = null;
                output = "";
                currentNestLevel = -1;
            }
        }
    }
    
    if (listStartIndex !== null && output.includes("<li>")) {
        if (currentNestLevel !== -1) {
            output += "</ul>";
        }
        replacements.push({start: listStartIndex, end: numElements - 1, text: output});
    }
    
    // Process the replacements in reverse to avoid index shifting issues
    for (var i = replacements.length - 1; i >= 0; i--) {
        var rep = replacements[i];
        
        for (var j = rep.end; j >= rep.start; j--) {
            var elementToRemove = cell.getChild(j);
            cell.removeChild(elementToRemove);
        }

        cell.insertParagraph(rep.start, rep.text);
    }
}

function formatTableNumberedLists() {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var tables = body.getTables();
    
    tables.forEach(table => {
        const numRows = table.getNumRows();
        for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
            const row = table.getRow(rowIndex);
            const numCells = row.getNumCells();

            for (let cellIndex = 0; cellIndex < numCells; cellIndex++) {
                const cell = row.getCell(cellIndex);
                
                formatCellNumberedLists(cell);
            }
        }
    });
}

function formatCellNumberedLists(cell) {
    var numElements = cell.getNumChildren();
    var output = "";
    var listStartIndex = null;
    var replacements = [];

    for (var i = 0; i < numElements; i++) {
        var para = cell.getChild(i);
        
        if (para.getType() === DocumentApp.ElementType.LIST_ITEM) {
            var listItem = para.asListItem();
            var glyphType = listItem.getGlyphType();
            
            if (listStartIndex === null) {
                listStartIndex = i;
            }
            
            if (glyphType === DocumentApp.GlyphType.NUMBER) {
                if (listStartIndex === i) {
                    output += "<ol>\n"; // Since there's no nesting, we open the <ol> tag at the start of a list
                }

                output += "<li>" + para.getText() + "\n";
            } else {
                if (listStartIndex !== null && output.includes("<li>")) {
                    output += "</ol>\n";  // Close the <ol> tag when we encounter a non-numbered list item
                    replacements.push({start: listStartIndex, end: i - 1, text: output});
                }
                
                listStartIndex = null;
                output = "";
            }
        } else {
            if (listStartIndex !== null && output.includes("<li>")) {
                output += "</ol>\n"; // Close the <ol> tag when we encounter a non-list item
                replacements.push({start: listStartIndex, end: i - 1, text: output});
                listStartIndex = null;
                output = "";
            }
        }
    }
    
    if (listStartIndex !== null && output.includes("<li>")) {
        output += "</ol>\n"; // Ensure we close the <ol> tag if the list is at the end
        replacements.push({start: listStartIndex, end: numElements - 1, text: output});
    }
    
    // Process the replacements in reverse to avoid index shifting issues
    for (var i = replacements.length - 1; i >= 0; i--) {
        var rep = replacements[i];
        
        for (var j = rep.end; j >= rep.start; j--) {
            var elementToRemove = cell.getChild(j);
            cell.removeChild(elementToRemove);
        }

        cell.insertParagraph(rep.start, rep.text);
    }
}

function formatGiveWellLinks() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var searchText = "https://www.givewell.org";
  var searchResult;

  // Create an array to store the positions of the found URLs.
  var urlPositions = [];

  // Find all instances of the URL and store their positions.
  searchResult = body.findText(searchText);
  while (searchResult != null) {
    // Store the start and end positions of the found URL.
    var foundElement = searchResult.getElement();
    var startOffset = searchResult.getStartOffset();
    var endOffset = searchResult.getEndOffsetInclusive();
    urlPositions.push({element: foundElement, start: startOffset, end: endOffset});

    // Continue searching the document.
    searchResult = body.findText(searchText, searchResult);
  }

  // Remove the URLs in reverse order to maintain the correct indexing.
  for (var i = urlPositions.length - 1; i >= 0; i--) {
    var position = urlPositions[i];
    position.element.deleteText(position.start, position.end);
  }
}

function addParagraphTags() {
  // Open the active document
  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  
  // Get all paragraphs in the document, including list items
  var paragraphs = body.getParagraphs();
  
  for (var i = 0; i < paragraphs.length; i++) {
    var paragraph = paragraphs[i];
    var elementType = paragraph.getType();
    var parent = paragraph.getParent();
    var parentType = parent.getType();
    
    // Check if the paragraph is a regular paragraph (excluding list items and table cells)
    if (elementType === DocumentApp.ElementType.PARAGRAPH && parentType !== DocumentApp.ElementType.TABLE_CELL) {
      var text = paragraph.getText();
      
      // Check if the paragraph is not empty
      if (text.trim() !== '') {
        // Update the paragraph text with <p> at the start and </p>&nbsp; at the end
        paragraph.setText('<p>' + text + '</p>&nbsp;');
      }
    }
    // No explicit action is needed for list items and table cells as they are skipped by the condition
  }
}

