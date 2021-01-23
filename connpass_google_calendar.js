// Creates an event for the moon landing and logs the ID.
//Google Calendar日時の更新
function updateGoogleCalendar(eventId, start, end){
    let event = CalendarApp.getDefaultCalendar().getEventById(eventId);
    const startDate = new Date(start);
    const endDate = new Date(end);
    event.setTime(startDate, endDate);  
  }
  
  //スプレットシートの更新
  function updateSheet(num, eventId, eventName, start, end, sheet){
    let idSetUpFlag = false;
    let timeEqualFlag = false;
    let lastRow = sheet.getLastRow();
    let targetRange = 0;
    
    for(let i = 2; i <= lastRow; i++){
      if(eventId == sheet.getRange(i, 1).getValue()){
        idSetUpFlag = true;
        sheet.getRange(i, 6).setValue("-");
        
        if(start != sheet.getRange(i, 3).getValue()
          || end != sheet.getRange(i, 4).getValue() ){
            timeEqualFlag = true;
            targetRange = i;
        }
        break;
      }
    }
    
    if(!idSetUpFlag){    
      let startTime = new Date(start);
      let endTime = new Date(end);
      let event = CalendarApp.getDefaultCalendar().createEvent(eventName, startTime, endTime);
      
      let newRow = lastRow + 1
      sheet.getRange(newRow, 1).setValue(eventId);
      sheet.getRange(newRow, 2).setValue(eventName);
      sheet.getRange(newRow, 3).setValue(start);
      sheet.getRange(newRow, 4).setValue(end);
      sheet.getRange(newRow, 5).setValue(event.getId());
      sheet.getRange(newRow, 6).setValue("-");   
      
      return;
    }
    
    if(timeEqualFlag){
      sheet.getRange(targetRange, 3).setValue(start);
      sheet.getRange(targetRange, 4).setValue(end);
      let targetEventId = sheet.getRange(targetRange, 5).getValue();
      updateGoogleCalendar(targetEventId, start, end);
    }
  }
  
  //削除用のキー
  function signDeleteKey(sheet){
    let lastRow = sheet.getLastRow();
    for(let i = 2; i <= lastRow; i++){
        sheet.getRange(i, 6).setValue("");
    } 
  }
  
  //シートとカレンダー削除
  function deleteSheet(sheet){
    let lastRow = sheet.getLastRow();
    for(let i = lastRow; i >= 2; i--){
      if(sheet.getRange(i, 6).getValue() == ""){
        let targetEventId = sheet.getRange(i, 5).getValue();
        let calendar = CalendarApp.getDefaultCalendar();
        let event = calendar.getEventById(targetEventId);
        event.deleteEvent();
        sheet.deleteRows(i);
      }
    } 
  }
  
  //メイン関数
  function mainFunction(){  
    let URL = 'https://connpass.com/api/v1/event/';
    let date = new Date();
    let year = date.getFullYear();
    let mouth = date.getMonth() + 1;
    let nickname = 'Hiraken123';
    
    let query = '?ym=' + year + mouth + '&owner_nickname=' + nickname;
    let QUERYURL = URL + query;
    let resp = UrlFetchApp.fetch(QUERYURL);
    let data = JSON.parse(resp.getContentText());
    //スプレットシートのID
    let spreadsheet = SpreadsheetApp.openById('<SpreedSheetId>');
    //シート名を指定
     let sheet;
    if(spreadsheet.getSheetByName(sheetName)==null){
        spreadsheet.insertSheet(sheetName);
        sheet = spreadsheet.getSheetByName(sheetName);
    }else{
        sheet = spreadsheet.getSheetByName(sheetName);
    }
  consol
    signDeleteKey(sheet);
      
    data.events.forEach((element, index)=> updateSheet(index + 2,element.event_id, element.title, element.started_at, element.ended_at, sheet));
    deleteSheet(sheet);
  }
  
