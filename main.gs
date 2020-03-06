var EMAIL_SENT = 'Sim';

var ssheet = SpreadsheetApp.getActiveSpreadsheet()
var configSheet = ssheet.getSheetByName('Config')

var logCell = configSheet.getRange('G2')
var logs = logCell.getValue();

var replyTo = "gestão@thunderatz.org";

function log(str) {
    var date = Utilities.formatDate(new Date(), "GMT-3", "[yyyy-MM-dd hh:mm:ss] ");

    logs += '\n' + date + str;
    logCell.setValue(logs);
}

// Enviar e-mails convidando para a entrevista
function send_interview() {
  log('Iniciando envio de e-mails de aprovação na entrevista')

  var from = configSheet.getRange('C3').getValue()
  var subject = configSheet.getRange('C4').getValue()
  var message = configSheet.getRange('C5').getValue()
  var sheetName = configSheet.getRange('C25').getValue()
  var emailsColumn = configSheet.getRange('C26').getValue()
  var confirmColumn = configSheet.getRange('C27').getValue()
  var startLine = configSheet.getRange('C28').getValue()

  log('De: ' + from);
  log('Assunto: ' + subject);
  log('Aba: ' + sheetName);
  log('Coluna de E-mails: ' + emailsColumn.toString());
  log('Coluna de Confirmação: ' + confirmColumn.toString());
  log('Linha Inicial: ' + startLine.toString());

  var sheet = ssheet.getSheetByName(sheetName);
  var rows = sheet.getLastRow();

  log('Linha Final: ' + rows);

  for (var i = startLine; i <= rows; i++) {
    var emailAddress = sheet.getRange(i, emailsColumn).getValue();
    var emailSent = sheet.getRange(i, confirmColumn).getValue();

    if (emailSent == EMAIL_SENT) {
      continue;
    }

    MailApp.sendEmail({
      name: from,
      to: emailAddress,
      subject: subject,
      htmlBody: message,
      replyTo: replyTo
    });

    log('E-mail enviado para ' + emailAddress);
    sheet.getRange(i, confirmColumn).setValue(EMAIL_SENT);
    SpreadsheetApp.flush();
  }
}

// Enviar e-mails convidando para a palestra geral
function send_palestra() {
  log('Iniciando envio de e-mails de convite para a palestra geral')

  var from = configSheet.getRange('C31').getValue()
  var subject = configSheet.getRange('C32').getValue()
  var message = configSheet.getRange('C33').getValue()
  var sheetName = configSheet.getRange('C53').getValue()
  var emailsColumn = configSheet.getRange('C54').getValue()
  var confirmColumn = configSheet.getRange('C55').getValue()
  var startLine = configSheet.getRange('C56').getValue()

  log('De: ' + from);
  log('Assunto: ' + subject);
  log('Aba: ' + sheetName);
  log('Coluna de E-mails: ' + emailsColumn.toString());
  log('Coluna de Confirmação: ' + confirmColumn.toString());
  log('Linha Inicial: ' + startLine.toString());

  var sheet = ssheet.getSheetByName(sheetName);
  var rows = sheet.getLastRow();

  log('Linha Final: ' + rows);

  for (var i = startLine; i <= rows; i++) {
    var emailAddress = sheet.getRange(i, emailsColumn).getValue();
    var emailSent = sheet.getRange(i, confirmColumn).getValue();

    if (emailSent == EMAIL_SENT) {
      continue;
    }

    MailApp.sendEmail({
      name: from,
      to: emailAddress,
      subject: subject,
      htmlBody: message,
      replyTo: replyTo
    });

    log('E-mail enviado para ' + emailAddress);
    sheet.getRange(i, confirmColumn).setValue(EMAIL_SENT);
    SpreadsheetApp.flush();
  }
}

function send_approved() {
  log('Iniciando envio de e-mails de aprovação no PS')

  var from = configSheet.getRange('C59').getValue()
  var subject = configSheet.getRange('C60').getValue()
  var message = configSheet.getRange('C61').getValue()
  var sheetName = configSheet.getRange('C81').getValue()
  var emailsColumn = configSheet.getRange('C82').getValue()
  var confirmColumn = configSheet.getRange('C83').getValue()
  var startLine = configSheet.getRange('C84').getValue()

  log('De: ' + from);
  log('Assunto: ' + subject);
  log('Aba: ' + sheetName);
  log('Coluna de E-mails: ' + emailsColumn.toString());
  log('Coluna de Confirmação: ' + confirmColumn.toString());
  log('Linha Inicial: ' + startLine.toString());

  var sheet = ssheet.getSheetByName(sheetName);
  var rows = sheet.getLastRow();

  log('Linha Final: ' + rows);

  for (var i = startLine; i <= rows; i++) {
    var emailAddress = sheet.getRange(i, emailsColumn).getValue();
    var emailSent = sheet.getRange(i, confirmColumn).getValue();

    if (emailSent == EMAIL_SENT) {
      continue;
    }

    MailApp.sendEmail({
      name: from,
      to: emailAddress,
      subject: subject,
      htmlBody: Utilities.formatString(message, emailAddress),
      replyTo: replyTo
    });

    log('E-mail enviado para ' + emailAddress);
    sheet.getRange(i, confirmColumn).setValue(EMAIL_SENT);
    SpreadsheetApp.flush();
  }
}

function send_approved() {
  log('Iniciando envio de e-mails de reprovação no PS')

  var from = configSheet.getRange('C87').getValue()
  var subject = configSheet.getRange('C88').getValue()
  var message = configSheet.getRange('C89').getValue()
  var sheetName = configSheet.getRange('C109').getValue()
  var emailsColumn = configSheet.getRange('C110').getValue()
  var confirmColumn = configSheet.getRange('C111').getValue()
  var startLine = configSheet.getRange('C112').getValue()

  log('De: ' + from);
  log('Assunto: ' + subject);
  log('Aba: ' + sheetName);
  log('Coluna de E-mails: ' + emailsColumn.toString());
  log('Coluna de Confirmação: ' + confirmColumn.toString());
  log('Linha Inicial: ' + startLine.toString());

  var sheet = ssheet.getSheetByName(sheetName);
  var rows = sheet.getLastRow();

  log('Linha Final: ' + rows);

  for (var i = startLine; i <= rows; i++) {
    var emailAddress = sheet.getRange(i, emailsColumn).getValue();
    var emailSent = sheet.getRange(i, confirmColumn).getValue();

    if (emailSent == EMAIL_SENT) {
      continue;
    }

    MailApp.sendEmail({
      name: from,
      to: emailAddress,
      subject: subject,
      htmlBody: message,
      replyTo: replyTo
    });

    log('E-mail enviado para ' + emailAddress);
    sheet.getRange(i, confirmColumn).setValue(EMAIL_SENT);
    SpreadsheetApp.flush();
  }
}

function onOpen() {
  var menuEntries = [];
  menuEntries.push({name: "E-mails entrevista", functionName: "send_interview"});
  menuEntries.push({name: "E-mails palestra geral", functionName: "send_palestra"});
  menuEntries.push({name: "E-mails aprovados", functionName: "send_approved"});
  menuEntries.push({name: "E-mails reprovados", functionName: "send_reproved"});

  ssheet.addMenu("E-mails", menuEntries);
  Logger.log('Menu criado');
}
