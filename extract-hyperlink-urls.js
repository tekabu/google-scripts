function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var a1 = sheet.getRange('A1');
  var oldText = a1.getValue();
  var newText = oldText;
  var formula = a1.getRichTextValue();
  Logger.log(a1.getValue());

  var runs = formula.getRuns();
  var replacements = [];

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
}

