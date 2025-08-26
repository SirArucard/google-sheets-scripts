function EnviarRelatorio() {
  var sheetName = "RELATÓRIO";
  var emailDestinado = "ti@maternoinfantilserra.org";
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var dataHoje = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");

  var spreadsheetId = spreadsheet.getId();
  var sheetId = sheet.getSheetId();  
  
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?'
          + 'format=pdf'
          + '&size=A4'
          + '&portrait=false'
          + '&fitw=true'
          + '&fzr=false'
          + '&gid=' + sheetId
          + '&range=A1:Q41';

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token
    }
  });

  var pdf = response.getBlob().setName('Relatório_de_Turno_' + dataHoje + '.pdf');


  var assunto = "Relatório Turno - " + dataHoje;
  var corpo = "Olá, o relatório diário de turno está anexado a este e-mail.";

  MailApp.sendEmail({
    to: emailDestinado,
    subject: assunto,
    body: corpo,
    attachments: [pdf]
  });
}