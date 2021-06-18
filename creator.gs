function myMail() {
  //Utilities.sleep(1000);
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let emailAddress = sheet.getRange(sheet.getLastRow(), 2).getValue();
  let arr = sheet.getRange(sheet.getLastRow(),4,1,3).getValues();
  let file = DocumentApp.openById('1_CcMjc17hMxc7lVZ9MXQE4UcR7DPXj6y1J5IQsqzlLI').getBlob();//Додатки до файлів
  let file1 = DocumentApp.openById('1WX3X0KnsnVRMo_Uhc4EQDUAMo1_kqvIkKUZGDDkJ_c8').getBlob();//Додатки до файлів
  MailApp.sendEmail ({ 
      to: emailAddress,
      subject: "Вебінар. Забезпечення кібероборони держави", 
      //htmlBody: "Шановний, "+ arr[0][0] + ' ' + arr[0][1] + ' ' + arr[0][2] + ", дякуємо за реєстрацію для участі у вебінарі! <br>"
      htmlBody: " Шановний(на) "+ arr[0][1] + ' ' + arr[0][2] + ", дякуємо за реєстрацію для участі у вебінарі! <br>"
                + 'Вебінар відбудеться 28 квітня 2021 року.<br>'
                + 'Тези доповіді та презентації для виступу <br>'
                + 'просимо надсилати на електронну адресу: yu.prybiliev@edu.nuou.org.ua <br>'
                + 'В додатках до цього листа містяться:<br>'
                + ' - Зразок оформлення тез доповіді <br>'
                + ' - Інформація щодо приєднання до кімнати проведення вебінару.<br>'
                + 'Чекаємо Вас на вебінарі!<br><br>'

                + 'З повагою Юрій Прібилєв <br>'
                + 'Тел.: +380 96 903 1003 <br>'
                + 'e-mail: yu.prybiliev@edu.nuou.org.ua',             
      name: "Юрій Прібилєв",
      attachments: [file, file1]
      });
      
}

function disaForm() {
  var formURL = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  var form = FormApp.openByUrl(formURL);
  //var formID = form.getId();
  form.setAcceptingResponses(false);
}

function enaForm() {
  var formURL = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  var form = FormApp.openByUrl(formURL);
  //var formID = form.getId();
  form.setAcceptingResponses(true);
}

function userClicked1() {
  Utilities.sleep(6000);
  var templateID = "1eBR-hLiYdNnZmKZM5TJqDeoK_3sFjIxQEop7AxuEgYU";//ID шаблону - презентація google
  var tableUrl = "https://docs.google.com/spreadsheets/d/1cUf_UH-3JbclwiirKQhSUTkolkLJoMmbZGp0UD_CaSY/edit#gid=1256370406";//ID поточної таблиці
  var wsdata = SpreadsheetApp.openByUrl(tableUrl).getSheetByName('Data');
  var ws = SpreadsheetApp.openByUrl(tableUrl).getSheetByName('Ответы на форму');
  var userInfopib = wsdata.getRange(ws.getRange("A1").getDataRegion().getLastRow(),5,1,1).getValue();
  var userInfopi = wsdata.getRange(ws.getRange("A1").getDataRegion().getLastRow(),6,1,1).getValue();
  //Генерирование
  var tmpfile = DriveApp.getFileById(templateID);
  tmpfile.makeCopy(userInfopib)
  var certNEW = DriveApp.getFilesByName(userInfopib).next();
  var certNEWID = certNEW.getId();
  var idFromTable = wsdata.getRange(ws.getRange("A1").getDataRegion().getLastRow(),7,1,1).getValue();
  var qrLinkFromTable = wsdata.getRange(ws.getRange("A1").getDataRegion().getLastRow(),9,1,1).getValue();
  var sl = SlidesApp.openById(certNEWID).getSlides();
  var shapes = sl[0].getShapes();
  shapes[1].getText().replaceAllText('{{name}}', userInfopi);
  shapes[2].getText().replaceAllText('{{idFromTable}}',idFromTable);
  Logger.log(qrLinkFromTable);
  sl[0].insertImage(qrLinkFromTable).setLeft(590).setTop(260);
  SlidesApp.openById(certNEWID).saveAndClose();
  var blob = DriveApp.getFileById(certNEWID).getBlob();
  var pdf = DriveApp.createFile(blob);
  var pdfUrl = pdf.getUrl();
  var pdfID = pdf.getId();
  pdf.moveTo(DriveApp.getFolderById('1Ww3jCAkRN_YT_xNl5jo2Vm18F43sKzjJ'));//ID папки де будуть зберігатися pdf сертифікати
  DriveApp.getFileById(certNEWID).setTrashed(true);
  wsdata.getRange(ws.getRange("A1").getDataRegion().getLastRow(),10,1,1).setValue(pdfUrl);
  wsdata.getRange(ws.getRange("A1").getDataRegion().getLastRow(),12,1,1).setValue(pdfID);
  disaForm();
  //return pdfUrl;
}

function sendCert() {
  let addressSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let contentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  let dataArrayFile = contentSheet.getRange(2,3,addressSheet.getRange("B1").getDataRegion().getLastRow()-1,10).getValues();
  for (let index in dataArrayFile) {
    if (dataArrayFile[index][4] != "#VALUE!" && dataArrayFile[index][9] !== "") {
      let file = DriveApp.getFileById(dataArrayFile[index][9]).getBlob();
      let emailAddress = addressSheet.getRange(Number(index)+2,2).getValue();
      let iPb = dataArrayFile[index][0] + " " + dataArrayFile[index][1];
      spam(emailAddress,iPb,file); 
    }
  }
}

function spam(emailAddress,iPb,file) {
  MailApp.sendEmail ({ 
      to: emailAddress,
      subject: "Вебінар. Забезпечення кібероборони держави",
      htmlBody: " Шановний(на) "+ iPb + ", дякуємо за участь у вебінарі! <br>"
                + 'Надсилаємо Вам сертифікат учасника.<br>'
                + 'До нових зустрічей!<br><br>'

                + 'З повагою Юрій Прібилєв <br>'
                + 'Тел.: +380 96 903 1003 <br>',             
      name: "Юрій Прібилєв",
      attachments: [file]
      });
}

function sendRemindToUsers() {
  var addressSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataArray = addressSheet.getRange(2,2,addressSheet.getRange("B1").getDataRegion().getLastRow()-1,1).getValues();
  for (let index in dataArray) {
    
    //if (/gmail.com/.test(dataArray[index][0])==true || /edu.nuou.org.ua/.test(dataArray[index][0])==true) {
        let emailAddress = addressSheet.getRange(Number(index)+2,2).getValue();
        let iB = addressSheet.getRange(Number(index)+2,5).getValue() + " " + addressSheet.getRange(Number(index)+2,6).getValue();
        MailApp.sendEmail ({ 
      to: emailAddress,
      subject: "Вебінар. Забезпечення кібероборони держави",
      htmlBody: "Шановний(на) "+ iB + ", нагадуємо адресу та код доступу до кімнати <br>"
                + 'Адреса кімнати проведення: https://meet.mil.gov.ua/b/jwr-sej-hhj-rtg<br>'
                + 'Код доступу до кімнати: 646816.<br>'
                + 'Приєднання до кімнати буде відкрито через 30 хвилин.<br>'
                + 'До зустрічі на вебінарі!.<br><br>'
                
                + 'З повагою Юрій Прібилєв <br>'
                + 'Тел.: +380 96 903 1003 <br>',             
      name: "Юрій Прібилєв",
      //attachments: [file]
      });
    //}
  }
}




function createTimeDrivenTriggers() {
  //var b = Utilities.formatDate(new Date(2021, 03, 19, 01, 10, 00),"GMT+3", "EEE' 'MMM' 'dd' 'HH:mm:ss' 'z' 'yyyy");
  //var a = new Date('April 19, 2021 02:32:00 +0300');//, "GMT+3", "EEE' 'MMM' 'dd' 'HH:mm:ss' 'z' 'yyyy");
  //Logger.log(b);
  //Logger.log(a);
  
  // Trigger disaForm at 19.04.2021 00:00.
  /*ScriptApp.newTrigger('disaForm')
      .timeBased()
      .at(a)
      .create();
  var a = new Date('April 27, 2021 09:00:00 +0300');
  // Trigger sendRemindToUsers at 27.04.2021 09:00.
  ScriptApp.newTrigger('sendRemindToUsers')
      .timeBased()
      .at(a)
      .create();
  var a = new Date('April 28, 2021 08:30:00 +0300');
  // Trigger sendRemindToUsers at 28.04.2021 08:30.
  ScriptApp.newTrigger('sendRemindToUsers')
      .timeBased()
      .at(a)
      .create();*/
  var a = new Date('April 28, 2021 13:30:00 +0300');
  // Trigger sendCert at 28.04.2021 13:30.
  ScriptApp.newTrigger('sendCert')
      .timeBased()
      .at(a)
      .create();
  /*var a = new Date('April 19, 2021 06:55:00 +0300');
  // Trigger enaForm at 19.04.2021 06:55.
  ScriptApp.newTrigger('enaForm')
      .timeBased()
      .at(a)
      .create();*/
}
