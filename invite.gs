var props=PropertiesService.getScriptProperties();

function createCalendar() {
  /* Studio-seginusカレンダーのIDを取得 */
  var calID=props.getProperty("calID");
  
  /*  */
  var plansSheet=props.getProperty("plans");

  /* メールアドレスリストの取得 */
  var mailList=getAddress();
  
  /* カレンダーの取得 */
  var cal=CalendarApp.getCalendarById(calID);
  
  /* プランを取得 */
  var plans=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(plansSheet).getDataRange().getValues();
  
  /* 撮影計画について、1行ずつ取得する */
  for(var i=1;i<plans.length;i++){
    var title=plans[i][0]; //イベント名を取得
    var start=plans[i][1]; //開始日時を取得
    var end=plans[i][2];  //終了日時を取得
    var description=plans[i][3];  //説明を取得
    var location=plans[i][4]; //場所を取得
    
    var guests=""; //必要人員のメールアドレス取得用変数
    
    /* getMemメソッドを呼び出し、イベント名から必要人員リストを配列で取得する */
    var memList=getMem(title);
    
    /* 必要人員ごとにメールアドレスリストからメールアドレスを取得し、変数にカンマ区切りで追記する */
    for(var j=0;j<memList.length;j++){
      guests+=mailList[memList[j]];
      guests+=",";
    }
    
    /* 末尾のカンマを除去する */
    guests=guests.slice(0,-1);
    
    /* カレンダー作成にあたってのオブション情報オブジェクトを作成する */
    var options={
      description: description,
      location: location,
      guests: guests,
      sendInvites: true
    };
    
    /* カレンダーを作成する */ 
    cal.createEvent(title, start, end, options);
  }
}


/*　フォームの回答からメールアドレスリストを返却する  */
function getAddress(){
  /* フォーム回答取得のための値をプロパティから取得する */
  var formID=props.getProperty("formID");
  var formSheet=props.getProperty("formSheet");
  
  /* フォーム回答を取得する */
  var lists=SpreadsheetApp.openById(formID).getSheetByName(formSheet).getDataRange().getValues();
  
  var returnObj={} //返却用連想配列を宣言
  
  /* フォームの情報から、役職をキーに、メールアドレスを値にセットする */
  for(var i=1;i<lists.length;i++){
    returnObj[lists[i][2]]=lists[i][1];
  }
  return returnObj;
}

/* イベント名をキーに、必要人員をコンフィグから取得する */
function getMem(event){
  /* コンフィグのシート名をプロパティから取得する */
  var confSheet=props.getProperty("configSheet");
  
  /* コンフィグシートを取得する */
  var conf=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(confSheet).getDataRange().getValues();
  
  /* コンフィグの情報から、必要人員リストを取得する */
  for(var i=1;i<conf.length;i++){
    if(conf[i][0]==event){
      return conf[i][1].split(",");
    }
  }
  Logger.log("該当するイベントがありません。");
}