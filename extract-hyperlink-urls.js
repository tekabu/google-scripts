function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  for (var i = 1; i <= numRows; i++) {
    for (var j = 1; j <= numCols; j++) {
      var cell = range.getCell(i, j);
      extractHyperlinkUrls(sheet, cell);
    }
  }
}

function extractHyperlinkUrls(sheet, answerCell) {
  var oldText = answerCell.getValue();

  if (oldText.length == 0) {
    sheet.getRange(answerCell.getRow(), 3).setValue(null);
    return;
  }

  var newText = oldText;
  var richTextValue = answerCell.getRichTextValue();
  Logger.log(oldText);

  var runs = richTextValue.getRuns();
  var replacements = [];

  if (runs.length == 0) {
    sheet.getRange(answerCell.getRow(), 3).setValue(null);
    return;
  }

  runs.forEach(run => {
    if (run.getLinkUrl() == null) {
      return;
    }

    var url = String(run.getLinkUrl());
    var text = String(run.getText()); 
    var href = "<a href=\""+url+"\">"+text+"</a>";

    replacements.push({
      "index": replacements.length,
      "startIndex": run.getStartIndex(),
      "endIndex": run.getEndIndex(),
      "url": url,
      "href": href,
      "text": text,
      "rLength": href.length - text.length,
    });
  });

  replacements.forEach(item => {
    newText = newText.substring(0, item.startIndex) + item.href + newText.substring(item.endIndex, newText.length);

    // update next configs
    // adjust start and end based on previously replaced string

    for (var i = item.index + 1; i < replacements.length; i++) {
      replacements[i].startIndex += item.rLength;
      replacements[i].endIndex += item.rLength;
    }
  });

  Logger.log(newText);
  sheet.getRange(answerCell.getRow(), 3).setValue(newText);
}
