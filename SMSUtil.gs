{\rtf1\ansi\ansicpg1252\cocoartf1504\cocoasubrtf830
{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx566\tx1133\tx1700\tx2267\tx2834\tx3401\tx3968\tx4535\tx5102\tx5669\tx6236\tx6803\pardirnatural\partightenfactor0

\f0\fs24 \cf0 function sendInvoicesBySMS() \{\
  //TODO remove\
  var date = new Date(2017, 7, 1);\
  //var date = new Date();\
  sendInvoicesBySMSByDate(date);\
\}\
\
\
function sendInvoicesBySMSByDate(date) \{\
  var formattedDate = Utilities.formatDate(date, "EAT", "dd/MM/yyyy");\
  var month = date.getMonth();\
  var year = date.getYear();\
  var billSheetName = generateMonthYearString(month, year);\
  var billForMonthSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(billSheetName);\
  var billForMonthSheetValues = billForMonthSheet.getDataRange().getValues();\
  var contactSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLIENT_CONTACT_SHEET);\
  var contactSheetValues = contactSheet.getDataRange().getValues();\
  var clientReferenceColumn = getColumnIndexByName(billSheetName, 'Reference', billForMonthSheetValues);\
  var soldValueColumn = getColumnIndexByName(billSheetName, 'Sold', billForMonthSheetValues);\
  var phone1Column = getColumnIndexByName(CLIENT_CONTACT_SHEET, 'Phone number 1', contactSheetValues);\
  var phone2Column = getColumnIndexByName(CLIENT_CONTACT_SHEET, 'Phone number 2', contactSheetValues);\
  //go through bills in the bill and send SMS\
  var nbRows = billForMonthSheet.getMaxRows();\
  for (var i = 1; i < nbRows; i++) \{\
    if(billForMonthSheetValues[i]) \{\
      var clientRef = billForMonthSheetValues[i][clientReferenceColumn];\
      var billValue = billForMonthSheetValues[i][soldValueColumn];\
      Logger.log(clientRef + ', ' + billValue);\
      \
      //get phone numbers\
      var phone1 = getCellByKey('Reference', 'Phone number 1', clientRef, contactSheet, contactSheetValues);\
      var phone2 = getCellByKey('Reference', 'Phone number 2', clientRef, contactSheet, contactSheetValues);\
      \
      Logger.log(phone1 + ', ' + phone2);\
      \
      var smsMessage = 'Votre facture du mois de : ' + getMonthAsAString(month) + ' s\\'\'e9l\'e8ve \'e0 : ' + billValue;\
      \
      Logger.log(smsMessage);\
      if(phone1) \{\
        //TODO remove\
        sendSMS('0326728879', smsMessage);\
        //sendSMS(phone1, smsMessage);\
      \}\
    \}\
  \}\
\}\
\
function sendSMS(phoneNumber, message) \{\
  var smsRequest = 'https://smswebservices.public.mtarget.fr/SmsWebServices/ServletSms?method=sendText' + \
    '&username=' + SMS_USER + \
    '&password=' + SMS_PASS + \
    '&serviceid=' + SMS_SERVICEID + \
    '&destinationAddress=' + '%2B261' + phoneNumber.substring(1) + \
    '&originatingAddress=00000&operatorid=0&paycode=0' + \
    '&msgtext=' + encodeURIComponent(message);\
  var options = \{\
    'method' : 'get',\
    'validateHttpsCertificates' : false\
 \};\
  var response = UrlFetchApp.fetch(smsRequest, options);\
  if(true) \{\
    Logger.log('smsRequest : ' + smsRequest);\
    Logger.log('');\
    Logger.log('smsResponse : ' + response);\
  \}\
\}\
\
}