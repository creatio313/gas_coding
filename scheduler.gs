/* プロパティを取得する */
var props=PropertiesService.getScriptProperties();
var configID=props.getProperty("configID");
var answerID=props.getProperty("answerID");
var colName=props.getProperty("colName");

var confSheet=props.getProperty("confSheet");
var answerSheet=props.getProperty("answerSheet");
var resultSheet=props.getProperty("resultSheet");

function main(){
  var config=getConfig(configID,confSheet);
  var schedule=getSchedule(answerID,answerSheet);
  var list=listUp(config,schedule);
  
  var result=[]; //Excel出力用二次元配列
  var datas=[]; //二次元配列内の配列
  var rowlen=0; //カラム数
  var collen=0; //行数
  
  /* 結果の要素数から、二次元配列を生成する */
  for(plan in list){
    for(col in list[plan]){
      datas.push(col);
    }
    rowlen=plan.length;
    datas.unshift("イベント名");
    break;
  }
  result.push(datas.slice(0));
  for(plan in list){
    datas.length=0;
    datas.push(plan);
    for(col in list[plan]){
      datas.push(list[plan][col].join(", "));
    }
    result.push(datas.slice(0));
  }
  
  collen=Object.keys(list).length+1;
  rowlen=result[0].length;

  if(collen==0||rowlen==0){Logger.log("まだ回答が不十分！　どんまい！"); return;}

  //Logger.log("行高さ"+rowlen+",列数"+collen);
  SpreadsheetApp.openById(configID).getSheetByName(resultSheet).getRange(1, 1, collen, rowlen).setValues(result);
}

/*
コンフィグ取得用関数
予定名と必要キャストをオブジェクト形式で返却する
*/
function getConfig(ID,Sheet) {
  /* 返却用のオブジェクトを宣言する */
  var confObj={};
  
  /* コンフィグのデータセルの値を二次元配列で取得し、タイトル行を削除する */
  var config=SpreadsheetApp.openById(ID).getSheetByName(Sheet).getDataRange().getValues();
  config.shift();
  
  /* 二次元配列から予定名と必要キャストを抜き出し、オブジェクトプロパティとして追加する */
  for(var i=0;i<config.length;i++){
    var row=config[i];
    confObj[row[0]]=row[1];
  }
  return confObj;
}

/*
スケジュール回答取得用関数
返却するオブジェクト形式は以下の通りである
{
  役職名={
    日付候補1={時間1,時間2},
    日付候補2={時間1,時間2}
  }
}
*/
function getSchedule(ID,Sheet){
  /* 処理に必要なプロパティ　*/
  var dayIndex=[]; //日付候補のインデックス
  var dayLists=[]; //日付の名前
  var answerObj={};  //返却用オブジェクト
  
  /* コンフィグのデータセルの値を二次元配列で取得する */
  var answer=SpreadsheetApp.openById(ID).getSheetByName(Sheet).getDataRange().getValues();
  
  /*　タイトル行を取得し、日付候補のインデックスを取得する */
  var header=answer[0];
  for(var i=0;i<header.length;i++){
    if(header[i].indexOf(colName)!=-1){
      dayIndex.push(i);
    }
  }
  
  /*　タイトル行を削除する */
  answer.shift();
  
  /* 各人の回答をそれぞれオブジェクト化し、連想配列に格納する */
  for(var i=0;i<answer.length;i++){
    answerObj[answer[i][2]]={}
    for(var j=0;j<dayIndex.length;j++){
      var index=dayIndex[j];
      answerObj[answer[i][2]][header[index]]=answer[i][index];
    }
  }
  return answerObj;
}

function listUp(config,schedule){

  var answerList=[]; //回答したメンバー
  var returnObj={};

  /* 回答したメンバーを回答オブジェクトから取得 */
  for(cast in schedule){
    answerList.push(cast);
  }
  
  for(plan in config){
    /* 実施予定のリストを１つずつ抽出。会議名はplanに、必要人員はpersonlistに配列として取得 */
    var personList=config[plan].split(",");
    var fullflg=true;  //必要キャスト全員回答フラグ
    var dayList=[] //撮影候補日

    /* 必要人員全員が回答しているかをチェックし、回答していない場合は全員回答フラグをfalseに設定する */
    for(required in personList){
      var judge=answerList.some(function(element){
        return element==personList[required];
      });
      if(!judge){
        fullflg=false;
        break;
      }
    }
    /* 全員回答フラグがtrueのとき、調整処理を実行する */
    if(fullflg){
      Logger.log(plan+"ぜんいんかいとう！！！");
      dayList=scheduleSearch(schedule,personList);
    }else{
      Logger.log(plan+"の人員全員が回答していません");
      continue;
    }
    returnObj[plan]=dayList;
  }
  return returnObj;
}

/* 一致するスケジュールを取得する */
function scheduleSearch(answer,casts){
  /* 候補日と時間のオブジェクトを取得する */
  var dayList=[]; //候補日・時間のオブジェクトの配列
  var time;
  var timesTemp=[];
  var times=[]; //日付格納
  var returnObj={};
  /* 人ごとの連想配列から候補日の配列に置き換える */
  for(cast in casts){
    dayList.push(answer[casts[cast]]);
  }
  /* １人目の日付候補を取り出す */
  for(day in dayList[0]){
    /* １人目の日付候補n日目の時間候補を分割して配列にする */
    var timeList=dayList[0][day].split(", ");
    /* もし時間候補がnullでない場合、２人目以降と比較する */
    if(timeList){
      /* 日付候補の値を照合する */
      for(var i=1;i<dayList.length;i++){
        /* ２人目以降の日付候補のn日目の時間候補を分割して配列にする */
        var timeList2=dayList[i][day].split(", ");
        /* 一致するものについて、変数timesTempに追加する */
        for(var j=0;j<timeList.length;j++){
          var judge=timeList2.some(function(element){
            return timeList[j]==element;
          });
          if(judge){
            time=timeList[j];
          }else{
            time=null;
            continue;
          }
          timesTemp.push(time);
        }
        /* 照合して残った時間候補をコピーして、次の人の時間候補と照合できるようにする */
        timeList=timesTemp.slice(0);
        /* timesTempを初期化する */
        timesTemp.length=0;
      }
      /* 絞られた値をtimesにコピーし、timeListは初期化する */
      times=timeList.slice(0);
      timeList.length=0;
    }else{
        continue;
    }
    /* 時間候補を日付をキーにしてオブジェクトに登録し、時間候補を初期化して次の日付へ進む */
    returnObj[day]=times.slice(0);
    times.length=0;
  }
  return returnObj;
}
