function doGet(e) {
var tmp = HtmlService.createTemplateFromFile('qr');
tmp.title = "Національний університет оборониУкраїни";
var idFromUrl = e.parameter.idFromTable;
tmp.text = handlerFunction(idFromUrl);         
return tmp.evaluate();
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function handlerFunction (idFromUrl) {
var tableUrl = '1cUf_UH-3JbclwiirKQhSUTkolkLJoMmbZGp0UD_CaSY'
var wsData = SpreadsheetApp.openById(tableUrl).getSheetByName('Data');
var ws = SpreadsheetApp.openById(tableUrl).getSheetByName('Ответы на форму');
var date = LanguageApp.translate(Utilities.formatDate(new Date(), "GMT+2", "dd MMMM yyyy"),"en","uk");
for (var i=2; i<=ws.getRange("A1").getDataRegion().getLastRow(); i++) {
  var number = wsData.getRange(i,7,1,1).getValue();
  if (number == idFromUrl) {
    var name = wsData.getRange(i,5,1,1).getValue();
    var klishe = '<img class="displayed" src= "http://drive.google.com/uc?export=view&id=1or-OkAojMsNhR5JVB6v70H0TL9VBQEgt" width="200" height="53" alt="sign">';
    var text = '<p align="center">дійсним засвідчується, що</p>' 
          +'<p align="center"><b>'+name+'</b></p>'
          +'<p align="center">отримав(ла) цей сертифікат учасника </p>'
          +'<p p align="center">№ '+idFromUrl+'</p><br>'
          +'<div class="pidpys">'
              +'<p>Голова організаційного комітету<p>'
              +'<p>'+klishe+'</p>'
              +'<p>Сергій МИКУСЬ</p><br>'
          +'</div>'
          +'<p align="center">28 квітня 2021 року <br>'
              
          +'м. КИЇВ</p><br>'
          +'<div class="date">'
            +'<p align="center">Відповідь на запит сформована</p>'
            +'<p align="center">'+date+'</p>'
          +'</div>';
    break;   
  }
  else {
    
    var text = '<div class="body">'+'<p align="center">сертифікат з номером:</p>'
          +'<p align="center">'+idFromUrl+' - не існує</p></div>'
          +'<div class="date">'
            +'<p align="center">Відповідь на запит сформована</p>'
            +'<p align="center">'+date+'</p>'
          +'</div>';
  }
}
return text
}


