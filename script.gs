function runScript() {
    var ui = DocumentApp.getUi();
    
    // Создаем диалоговое окно для ввода данных
    var html = HtmlService.createHtmlOutputFromFile('Dialog')
        .setWidth(400)
        .setHeight(300);
    
    ui.showDialog(html);
  }
  
  function processData(correctWord, partWord, prefix, suffix) {
    var spreadsheet = SpreadsheetApp.create('Результаты проверки');
    var sheet = spreadsheet.getActiveSheet();
    
    // Записываем заголовки
    sheet.appendRow(['Номер страницы', 'Номер позиции', 'Найденное слово', 'Результат']);
    
    // Получаем данные из активного документа
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    var results = [];
    
    var totalParagraphs = paragraphs.length;
    var paragraphsPerPage = 50; // Приблизительное количество параграфов на странице, можно настроить
  
    paragraphs.forEach(function(paragraph, index) {
      var paragraphText = paragraph.getText();
      var pageNumber = Math.ceil((index + 1) / paragraphsPerPage); // Приближенное вычисление номера страницы
      var words = extractWords(paragraphText);
      
      words.forEach(function(word) {
        if (word.includes(partWord)) {
          var cleanWord = removePrefixSuffix(word, prefix, suffix);
          var result = checkWord(cleanWord, correctWord);
          var position = getPosition(paragraphText, word);
          results.push([pageNumber, position, cleanWord, result]);
        }
      });
    });
    
    // Записываем результаты в таблицу
    results.forEach(function(result) {
      sheet.appendRow(result);
    });
    
    // Открываем Google Spreadsheet в новом окне
    var url = spreadsheet.getUrl();
    var html = HtmlService.createHtmlOutput(
        '<html><script>window.open("' + url + '", "_blank");google.script.host.close();</script></html>'
    ).setWidth(100).setHeight(100);
    
    DocumentApp.getUi().showModalDialog(html, 'Открытие Google Spreadsheet');
  }
  
  function extractWords(text) {
    var regex = /\b\w+\b/g; // Регулярное выражение для извлечения слов
    var matches = [];
    var match;
    
    while ((match = regex.exec(text)) !== null) {
      matches.push(match[0]);
    }
    
    return matches;
  }
  
  function removePrefixSuffix(word, prefix, suffix) {
    if (prefix && word.startsWith(prefix)) {
      word = word.substring(prefix.length);
    }
    if (suffix && word.endsWith(suffix)) {
      word = word.substring(0, word.length - suffix.length);
    }
    return word;
  }
  
  function checkWord(foundWord, correctWord) {
    if (foundWord === correctWord) {
      return 'ОК';
    }
    
    var similar = false;
    var correctPattern = correctWord.replace(/_/g, '.').replace(/\*/g, '.*');
    
    if (new RegExp(correctPattern, 'i').test(foundWord)) {
      similar = true;
    }
    
    if (similar) {
      return 'Обратить внимание';
    }
    
    return 'Ошибка';
  }
  
  function getPosition(text, word) {
    var index = text.indexOf(word);
    return index !== -1 ? index + 1 : 'позиция не определена';
  }
  
  
  
  
  