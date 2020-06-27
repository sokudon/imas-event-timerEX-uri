/**
 * Return a list of sheet names in the Spreadsheet with the given ID.
 * @param {String} a Spreadsheet ID.
 * @return {Array} A list of sheet names.
 */


var sid="1CpwNLrurUVVLX2dmMgZHU-uQC7WQfyfWqLlaiooRaN8";
var sname="EXたいま";


function doGet() {
  var ss = SpreadsheetApp.openById(sid);
  var sheets = ss.getSheetByName(sname);
  
　var last_row = sheets.getRange(1,6).getValue();
　var last_col = 6;
  
  var values= sheets.getRange(1,1,last_row ,last_col).getValues();
 var value = JSON.parse(JSON.stringify(values));
 
  var moment = Moment.load();
  
  var html="";
  var url="http://www.shurey.com/js/timer/countdown.html?C,"; //154626840,
  var url2='https://www.timeanddate.com/countdown/gaming?iso=YYYYMMDDThhmm&p0=248&msg=event&font=slab&csz=1'
  var neta="http://sokudon.s17.xrea.com/neta/imm.html#";
  var　wtime='http://sokudon.s17.xrea.com/sekai-dere.html';
  var kaigai ="";
  
  for(var k=1;k<value.length+1;k++){
    
    
   
 if(value[k-1][0]=="ミリシタ海外版"){
      url2=url2.replace("235","102");
 var type="["+value[k-1][1]+"]";
 var titleraw=type + "偶像大師百萬人演唱會！ 劇場時光"; 
 var title=type    + "偶像大師百萬人演唱會！ 劇場時光" +"～　限定活動開始";
 var title2=type   + "偶像大師百萬人演唱會！ 劇場時光"  +"～　限定活動結束"; 
 var stat= moment(value[k-1][2]).add("h",1);
 var end= moment(value[k-1][3]).add("h",1);      
  var dformat=",M$D2,MM/DD(ddd) HH:mm:ss(GMTZ%DST)"       
  kaigai="YYYY-MM-DDTHH:mm:00+08:00"
    }
    else{ 
      
  var type="["+value[k][1]+"]";
  var titleraw=type + value[k][0]; 
  var title=type    + value[k][0] +"～　開始";
  var title2=type   + value[k][0]  +"～　終了"; 
  var stat= value[k][2];
  var end= value[k][3];
  var timest= value[k][5];
  var dformat=",M$EB,MM/DD(ddd) HH:mm:ss(GMTZ%DST)"    
  if(value[k][0]=="ミリシタ海外版"){
    
  var titleraw=type + "아이돌마스터 밀리언 라이브! 시어터 데이즈"; 
  var title=type    + "아이돌마스터 밀리언 라이브! 시어터 데이즈" +"～　이벤트시작";
  var title2=type   + "아이돌마스터 밀리언 라이브! 시어터 데이즈"  +"～　이벤트종료"; 
      url2=url2.replace("248","235");
  var dformat=",M$E6,MM/DD(ddd) HH:mm:ss(GMTZ%DST)"    
    }
    }
  
  //{}|\^[]`<>#"%
  var reg = new RegExp('[<>#"%]',"gm");
  title = title.replace(reg,"");
  titleraw = titleraw.replace(reg,"");
  title = title.replace(/\?/,"？");
  titleraw = titleraw.replace(/\?/,"？");
  var endst ="null";
  if(end.toString().match(/Z$/)){
    endst=  moment(end).format();
  }
  
  if(stat!=""){
  stat = (moment(stat).valueOf()/10000).toFixed(0);
  
  
   var i, len, arr;
        for(i=0,len=title.length,arr=[]; i<len; i++) {
          if(title.charCodeAt(i) < 0x80){
            arr += title.substr(i, 1);
          }
          else{
            arr +="%25u"+  ("00"+title.charCodeAt(i).toString(16)).slice(-4);
          }
        }
     if(kaigai=="YYYY-MM-DDTHH:mm:00+08:00"){
     html += "<tr><td>"+ hyperlink(neta + encodeURIComponent(titleraw)+","+moment(stat*10000-3600000).format(kaigai)+","+endst+dformat,titleraw) +"</td>"
    html+= "<td>"+ hyperlink( url +stat +"," +arr,moment(stat*10000-3600000).format(kaigai)) +"</td>";
    }
    else{
     html += "<tr><td>"+ hyperlink(neta + encodeURIComponent(titleraw)+","+moment(stat*10000).format(kaigai)+","+endst+dformat,titleraw) +"</td>"
    html+= "<td>"+ hyperlink( url +stat +"," +arr,moment(stat*10000).format(kaigai)) +"</td>";
    
    }}
  if(end!=""){
  end = (moment(end).valueOf()/10000).toFixed(0);
    var i, len, arr;
        for(i=0,len=title2.length,arr=[]; i<len; i++) {
          if(title.charCodeAt(i) < 0x80){
            arr += title2.substr(i, 1);
          }
          else{
            arr +="%25u"+  ("00"+title2.charCodeAt(i).toString(16)).slice(-4);
          }
        }
    if(kaigai=="YYYY-MM-DDTHH:mm:00+08:00"){
      
      if(end!="NaN"){
      html += "<td>"+ hyperlink( url +end +"," +arr,moment(end*10000 -3600000).format(kaigai))+" " + hyperlink(url2.replace(/YYYYMMDDThhmm/,moment(end*10000 -3600000).format("YYYYMMDDTHHmm")).replace("event",title2),"time&date.com") +"</td>";
      }
      else{
        html += "<td>--(ENDpending)</td>"
      }
        if(type=="[ミリシタ]"){
        html+= "<td>"+ hyperlink(wtime +"#"+encodeURIComponent(titleraw)+","+encodeURIComponent(moment(stat*10000-3600000).format(kaigai)) +","+encodeURIComponent(moment(end*10000-3600000).format(kaigai))+dformat.replace(/,MM.+$/,"")+",MT495","view worldtime,HKT495") +"</td>";
      }
    }
    else{
      if(end!="NaN"){
  html += "<td>"+ hyperlink( url +end +"," +arr,moment(end*10000).format(kaigai))+" " + hyperlink(url2.replace(/YYYYMMDDThhmm/,moment(end*10000).format("YYYYMMDDTHHmm")).replace("event",title2),"time&date.com") +"</td>";
      }
      else{
        html += "<td>--(ENDpending)</td>"
      }
      
        html+= "<td>"+ hyperlink(wtime +"#"+encodeURIComponent(titleraw)+","+encodeURIComponent(moment(stat*10000).format(kaigai)) +","+encodeURIComponent(moment(end*10000).format(kaigai))+dformat.replace(/,MM.+$/,"")+","+timest,"view worldtime," +timest) +"</td>";

    }
  }
    html += "</tr>";
  }  
  
  var header= "<style>th,td{  border:solid 1px #aaaaaa;},.table-scroll{  overflow-x : auto}</style>";
  var h ="<table><thead><tr><th>いべんと名(EVENTNAME)(はんようたいまたいむぞーん対応)</th><th>開始日(START)からのみ</th><th>終了日(END)からのみ</th><th>せかいどけい(WLDtime)</th></tr></thead>";  
  
  html= h+"<tbody>"+ header  +html + "<tbody></table>";
  
  html += '<table><tbody><tr><td>IMAS-GAME-CALENDER</td><td><a href="http://sokudon.s17.xrea.com/imas_event_calender.html" target="_blank">GOOG_CALENDAR</a></td></td></tbody></table>';
  html += "datetime source;<br>DERESUTE JPN https://imascg-slstage-wiki.gamerch.com/<br>MIRSITA JPN wiki https://imasml-theater-wiki.gamerch.com/<br>MIRISITA KOREA https://cafe.naver.com/imasmltheaterkr/120485<br>MIRISITA CHINA/TAIWAN https://www.facebook.com/idolmastermlTD.ch/<br>"
  return HtmlService.createHtmlOutput(html);
  //return ContentService.createTextOutput(url).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

//<>#"%　"#"はURI参照として、"%"はエスケープ用文字として使われます。
//除外されている記号 (RFC2396 に定義がないもの)
//以下の文字は使用できません。
// {}|\^[]`<>#"%

function hyperlink(link,a){
  link= "<a href='" + link +"' target=\"_blank\" rel=\"noopener\">" +a +"</a>";
  
  return link;
}

function wmap_getSheetsName(sheets){
  //var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet_names = new Array();
  
  if (sheets.length >= 1) {  
    for(var i = 0;i < sheets.length; i++)
    {
      sheet_names.push(sheets[i].getName());
    }
  }
  return sheet_names;
}