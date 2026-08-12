var MASTER_SHEET='スタッフマスタ',ACTIVE_SHEET='在籍スタッフ',RESIGNED_SHEET='退職者';
var COL_SHA_CD=1,COL_HPCD=2,COL_GENDER=3,COL_NAME=4,COL_ENTER=5,COL_RESIGN=6;
var COL_KUBUN=7,COL_DETAIL=8,COL_BIRTH=9,COL_CLUSTER=10,COL_TRANSPORT=11;
var COL_PARKING=12,COL_SKILL=13,COL_WAGE=14,COL_WORK_INFO=15,COL_SIZE_TOP=16;
var COL_LEND_CNT=17,COL_SIZE_SHOE=18,COL_CAR_INFO=19,COL_MOBILE=20;
var COL_EMERGENCY=21,COL_POSTCODE=22,COL_ADDRESS=23,COL_TENURE=24,COL_AGE=25;

function onOpen(){
  var ui=SpreadsheetApp.getUi();
  ui.createMenu('スタッフ管理')
    .addItem('新規スタッフ登録','showNewStaffForm')
    .addItem('スタッフ情報変更','showEditStaffForm')
    .addItem('退職処理','showRetirementDialog')
    .addSeparator()
    .addItem('在籍スタッフ一覧を更新','formatForA4')
    .addItem('退職者一覧を更新','formatResignedList')
    .addSubMenu(ui.createMenu('チェックリストを更新')
      .addItem('中（A4に6枚・大きめ）','updateCheckListMedium')
      .addItem('小（A4に8枚・小さめ）','updateCheckListSmall'))
    .addItem('面談シートに新規在籍者を追加','updateInterviewSheet')
    .addSeparator()
    .addItem('【テンプレート化】全データ削除','clearForTemplate')
    .addToUi();
}

function showNewStaffForm(){
  var h=HtmlService.createHtmlOutput(buildFormHtml(getNextHpcd(),Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd'),getFormDefaults())).setWidth(500).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(h,'新規スタッフ登録');
}

function getFormDefaults(){
  var rows=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET).getDataRange().getValues();
  var maxCD=0,lastWage='';
  for(var i=1;i<rows.length;i++){
    var s=Number(rows[i][0]),w=Number(rows[i][COL_WAGE-1]);
    if(!isNaN(s)&&s>maxCD)maxCD=s;
    if(!isNaN(w)&&w>0)lastWage=w;
  }
  var cy=new Date().getFullYear();
  return{shaCD:maxCD>0?maxCD+1:'',wage:lastWage||'',defaultBirthYear:cy-20,currentYear:cy};
}

function getNextHpcd(){
  var rows=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET).getDataRange().getValues();
  var max=462000;
  for(var i=1;i<rows.length;i++){var h=Number(rows[i][COL_HPCD-1]);if(!isNaN(h)&&h>max)max=h;}
  return max+1;
}

function westernToEraGAS(y){
  if(y>=2019)return{name:'令和',year:y-2018};
  if(y>=1989)return{name:'平成',year:y-1988};
  if(y>=1926)return{name:'昭和',year:y-1925};
  return{name:'大正',year:y-1911};
}

function buildYearOptions(d){
  var o='<option value="">年を選択してください</option>';
  for(var y=d.currentYear-14;y>=1955;y--){
    var e=westernToEraGAS(y),sel=y===d.defaultBirthYear?' selected':'';
    o+='<option value="'+y+'"'+sel+'>'+y+'年（'+e.name+e.year+'年）</option>';
  }
  return o;
}

function buildShoeSizeOptions(){
  var o='<option value="">選択</option>';
  for(var i=215;i<=290;i+=5)o+='<option value="'+(i/10)+'">'+(i/10)+'</option>';
  return o;
}

function lookupPostalCode(zip){
  var c=zip.replace(/[^0-9]/g,'');
  if(c.length!==7)return null;
  try{
    var r=UrlFetchApp.fetch('https://api.zipaddress.net/?zipcode='+c,{muteHttpExceptions:true});
    if(r.getResponseCode()===200){
      var j=JSON.parse(r.getContentText());
      if(j&&j.code===200&&j.data&&j.data.fullAddress)return j.data.fullAddress;
    }
  }catch(e){throw new Error('住所検索エラー: '+e.message);}
  return null;
}

function buildFormHtml(nextHpcd,today,d){
  var yo=buildYearOptions(d),so=buildShoeSizeOptions(),h='';
  h+='<!DOCTYPE html><html><head><meta charset="UTF-8"><style>';
  h+='*{box-sizing:border-box;margin:0;padding:0}body{font-family:Meiryo,sans-serif;font-size:13px;color:#202124}';
  h+='.fb{padding:14px;overflow-y:auto;max-height:600px}.sec{margin-bottom:14px}';
  h+='.st{font-size:11px;font-weight:bold;color:#1a73e8;background:#e8f0fe;padding:4px 8px;border-radius:3px;margin-bottom:8px}';
  h+='.gr{display:grid;grid-template-columns:1fr 1fr;gap:7px 12px}.fu{grid-column:1/-1}';
  h+='label,.slb{display:block;font-size:11px;font-weight:bold;color:#555;margin-bottom:3px}';
  h+='.bj{font-size:10px;padding:1px 4px;border-radius:3px;margin-left:3px}';
  h+='.rq{background:#fce8e6;color:#c5221f}.at{background:#e6f4ea;color:#137333}.op{background:#f1f3f4;color:#80868b}';
  h+='input,select{width:100%;padding:6px 8px;border:1px solid #dadce0;border-radius:4px;font-size:13px;font-family:Meiryo,sans-serif}';
  h+='input:focus,select:focus{outline:none;border-color:#1a73e8}';
  h+='.hv{font-size:22px;font-weight:bold;color:#1a73e8;padding:3px 0}';
  h+='.byr{display:grid;grid-template-columns:1fr 1fr;gap:8px}';
  h+='.eg{display:flex;gap:4px;align-items:center}.eg select{width:80px;flex:none}.eg input{width:52px;text-align:center;flex:none}.eg span{white-space:nowrap}';
  h+='.cw{border:1px solid #dadce0;border-radius:6px;padding:8px;background:#f8f9fa;margin-top:8px}';
  h+='.cn{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px}';
  h+='.nb{background:#e8f0fe;color:#1a73e8;border:none;border-radius:4px;padding:3px 12px;cursor:pointer;font-size:15px}';
  h+='.ct{font-size:13px;font-weight:bold;color:#333}';
  h+='.cg{display:grid;grid-template-columns:repeat(7,1fr);gap:1px;text-align:center}';
  h+='.ch{font-size:11px;font-weight:bold;padding:3px 0;color:#666}.ch.su{color:#d93025}.ch.sa{color:#1a73e8}';
  h+='.cd{padding:5px 0;border-radius:50%;font-size:12px}.cd:not(.em){cursor:pointer}.cd:not(.em):hover{background:#e8f0fe}';
  h+='.cd.su{color:#d93025}.cd.sa{color:#1a73e8}.cd.sl{background:#1a73e8!important;color:white!important}.cd.em{cursor:default}';
  h+='.bd{text-align:center;font-size:12px;font-weight:bold;color:#1a73e8;padding-top:6px;min-height:20px}';
  h+='.zr{display:flex;gap:6px}.zr input{flex:1}';
  h+='.zb{padding:6px 10px;background:#e8f0fe;color:#1a73e8;border:1px solid #1a73e8;border-radius:4px;cursor:pointer;font-size:12px;white-space:nowrap}';
  h+='.zs{font-size:11px;margin-top:4px;min-height:16px}';
  h+='.msg{display:none;margin:0 14px 10px;padding:8px 12px;border-radius:4px;font-size:12px;font-weight:bold}';
  h+='.msg.ok{background:#e6f4ea;color:#137333}.msg.err{background:#fce8e6;color:#c5221f}';
  h+='.ft{padding:10px 14px;border-top:1px solid #e0e0e0;display:flex;gap:8px}';
  h+='.bn{flex:1;padding:10px;border:none;border-radius:4px;font-size:13px;font-weight:bold;cursor:pointer;font-family:Meiryo,sans-serif}';
  h+='.bc{background:#f1f3f4;color:#3c4043}.bo{background:#1a73e8;color:white}.bo:disabled{background:#a8c7fa}';
  h+='</style></head><body><div class="fb">';
  h+='<div class="sec"><div class="st">基本情報</div><div class="gr">';
  h+='<div><label>HPCD <span class="bj at">自動採番（変更可）</span></label><input type="text" id="hpcd" value="'+nextHpcd+'" placeholder="例: '+nextHpcd+'"></div>';
  h+='<div><label>社員CD <span class="bj op">任意</span></label><input type="text" id="shaCD" value="'+d.shaCD+'" placeholder="例: 101056"></div>';
  h+='<div class="fu"><label>名前 <span class="bj rq">必須</span></label><input type="text" id="name" placeholder="例: 山田　太郎"></div>';
  h+='<div><label>性別</label><select id="gender"><option value="">選択</option><option>男</option><option>女</option></select></div>';
  h+='<div><label>入社日 <span class="bj rq">必須</span></label><input type="date" id="enterDate" value="'+today+'"></div>';
  h+='<div class="fu"><label>生年月日</label><div class="byr">';
  h+='<div><div class="slb">西暦</div><select id="birthYear" onchange="onWC()">'+yo+'</select></div>';
  h+='<div><div class="slb">和暦（直接入力可）</div><div class="eg">';
  h+='<select id="eraName" onchange="onEC()"><option value="">元号</option><option value="reiwa">令和</option><option value="heisei">平成</option><option value="showa">昭和</option><option value="taisho">大正</option></select>';
  h+='<input type="number" id="eraYear" min="1" max="99" placeholder="年" onblur="onEC()"><span>年</span></div></div></div>';
  h+='<div id="calSec" style="display:none"><div class="cw"><div class="cn">';
  h+='<button class="nb" type="button" onclick="prevM()">◀</button><span id="calTitle" class="ct"></span><button class="nb" type="button" onclick="nextM()">▶</button>';
  h+='</div><div class="cg" id="calGrid"></div><div class="bd" id="birthDisp"></div></div></div>';
  h+='</div></div></div>';
  h+='<div class="sec"><div class="st">雇用区分・役職</div><div class="gr">';
  h+='<div><label>区分</label><select id="kubun"><option value="">選択</option><option>高校</option><option>大学</option><option>短大</option><option>専門</option><option>主婦</option><option>フリーター</option><option>フリー</option><option>社員</option></select></div>';
  h+='<div><label>詳細</label><input type="text" id="detail" placeholder="例: 近大、主婦 など"></div>';
  h+='<div><label>クラスター</label><select id="cluster"><option value="">選択</option><option selected>ルーキー</option><option>1クルー</option><option>2クルー</option><option>3クルー</option><option>1トレーナー</option><option>マネージャー</option></select></div>';
  h+='<div><label>時給（円）</label><input type="number" id="wage" value="'+d.wage+'" placeholder="例: 1100"></div>';
  h+='<div class="fu"><label>勤務メモ</label><input type="text" id="workInfo" placeholder="例: 週3、土日〇、掛持ちあり"></div>';
  h+='</div></div>';
  h+='<div class="sec"><div class="st">勤務条件</div><div class="gr">';
  h+='<div><label>交通手段</label><select id="transport"><option value="">選択</option><option>車</option><option>原付</option><option>自転車</option><option>電車</option><option>バス</option><option>歩き</option><option>送迎え</option></select></div>';
  h+='<div><label>駐車場</label><select id="parking"><option value="">なし</option><option value="〇">あり（〇）</option></select></div>';
  h+='<div><label>スキル</label><select id="skill"><option value="">選択</option><option value="D・K">D・K（デシャップ＋キッチン）</option><option value="D">D（デシャップ）</option><option value="K">K（キッチン）</option></select></div>';
  h+='</div></div>';
  h+='<div class="sec"><div class="st">サイズ・貸与品</div><div class="gr">';
  h+='<div><label>上サイズ</label><select id="sizeTop"><option value="">選択</option><option>S</option><option>M</option><option>L</option><option>LL</option><option>3L</option></select></div>';
  h+='<div><label>貸与枚数</label><select id="lendCnt"><option value="">選択</option><option value="1" selected>1枚</option><option value="2">2枚</option><option value="3">3枚</option></select></div>';
  h+='<div><label>靴サイズ（cm）</label><select id="sizeShoe">'+so+'</select></div>';
  h+='<div><label>車情報</label><input type="text" id="carInfo" placeholder="例: 筑豊500あ1234 プリウス"></div>';
  h+='</div></div>';
  h+='<div class="sec"><div class="st">連絡先・住所</div><div class="gr">';
  h+='<div><label>携帯番号</label><input type="tel" id="mobile" placeholder="例: 090-1234-5678"></div>';
  h+='<div><label>緊急連絡先</label><input type="tel" id="emergency" placeholder="例: 080-1234-5678"></div>';
  h+='<div class="fu"><label>郵便番号</label><div class="zr"><input type="text" id="postcode" placeholder="例: 820-0001" maxlength="8"><button class="zb" type="button" onclick="lookupZip()">住所を検索</button></div></div>';
  h+='<div class="fu"><label>住所</label><input type="text" id="address" placeholder="郵便番号検索後、番地・号を追加してください"><div class="zs" id="zipSt"></div></div>';
  h+='</div></div>';
  h+='</div>';
  h+='<div id="msg" class="msg"></div>';
  h+='<div class="ft"><button class="bn bc" onclick="google.script.host.close()">キャンセル</button><button class="bn bo" id="saveBtn" onclick="doSave()">登録する</button></div>';
  h+='<script>';
  h+='var cY,cM,bSel=null;';
  h+='function getMD(){var v=document.getElementById("enterDate").value;if(v){var p=v.split("-");return{m:parseInt(p[1]),d:parseInt(p[2])};}var n=new Date();return{m:n.getMonth()+1,d:n.getDate()};}';
  h+='function w2e(y){y=parseInt(y);if(y>=2019)return{code:"reiwa",name:"令和",year:y-2018};if(y>=1989)return{code:"heisei",name:"平成",year:y-1988};if(y>=1926)return{code:"showa",name:"昭和",year:y-1925};return{code:"taisho",name:"大正",year:y-1911};}';
  h+='function e2w(era,eY){eY=parseInt(eY);if(!era||isNaN(eY)||eY<1)return null;if(era==="reiwa")return eY+2018;if(era==="heisei")return eY+1988;if(era==="showa")return eY+1925;if(era==="taisho")return eY+1911;return null;}';
  h+='function syncJp(y){var e=w2e(y);document.getElementById("eraName").value=e.code;document.getElementById("eraYear").value=e.year;}';
  h+='function syncYSel(y){var s=document.getElementById("birthYear");s.value=y;if(parseInt(s.value)===parseInt(y))syncJp(y);}';
  h+='function openCal(y){var md=getMD();cY=parseInt(y);cM=md.m;bSel={year:cY,month:cM,day:md.d};renderCal();document.getElementById("calSec").style.display="block";}';
  h+='function onWC(){var y=parseInt(document.getElementById("birthYear").value);if(!y){document.getElementById("calSec").style.display="none";bSel=null;return;}syncJp(y);openCal(y);}';
  h+='function onEC(){var era=document.getElementById("eraName").value,eY=document.getElementById("eraYear").value,w=e2w(era,eY);if(!w)return;var s=document.getElementById("birthYear");s.value=w;if(parseInt(s.value)===w)openCal(w);}';
  h+='function prevM(){cM--;if(cM<1){cM=12;cY--;syncYSel(cY);}renderCal();}';
  h+='function nextM(){cM++;if(cM>12){cM=1;cY++;syncYSel(cY);}renderCal();}';
  h+='function selDay(day){bSel={year:cY,month:cM,day:day};syncYSel(cY);renderCal();}';
  h+='function renderCal(){';
  h+='document.getElementById("calTitle").textContent=cY+"年 "+cM+"月";';
  h+='var fd=new Date(cY,cM-1,1).getDay(),ld=new Date(cY,cM,0).getDate();';
  h+='var g=document.getElementById("calGrid");g.innerHTML="";';
  h+='var hd=["日","月","火","水","木","金","土"];';
  h+='for(var i=0;i<7;i++){var th=document.createElement("div");th.className="ch"+(i===0?" su":i===6?" sa":"");th.textContent=hd[i];g.appendChild(th);}';
  h+='for(var j=0;j<fd;j++){var em=document.createElement("div");em.className="cd em";g.appendChild(em);}';
  h+='for(var day=1;day<=ld;day++){';
  h+='var dow=(fd+day-1)%7,cl=document.createElement("div");';
  h+='cl.className="cd"+(dow===0?" su":dow===6?" sa":"")+(bSel&&bSel.year===cY&&bSel.month===cM&&bSel.day===day?" sl":"");';
  h+='cl.textContent=day;cl.setAttribute("data-d",day);';
  h+='cl.onclick=function(){selDay(parseInt(this.getAttribute("data-d")));};';
  h+='g.appendChild(cl);}';
  h+='if(bSel){var era=w2e(bSel.year);document.getElementById("birthDisp").textContent=bSel.year+"年"+bSel.month+"月"+bSel.day+"日（"+era.name+era.year+"年）";}}';
  h+='function lookupZip(){';
  h+='var zip=document.getElementById("postcode").value.replace(/[^0-9]/g,""),st=document.getElementById("zipSt");';
  h+='if(zip.length!==7){st.textContent="7桁の郵便番号を入力してください";st.style.color="#c5221f";return;}';
  h+='st.textContent="検索中...";st.style.color="#888";';
  h+='google.script.run';
  h+='.withSuccessHandler(function(addr){if(addr){document.getElementById("address").value=addr;st.textContent="✓ 住所を入力しました（番地・号を追記してください）";st.style.color="#137333";}else{st.textContent="該当する住所が見つかりませんでした";st.style.color="#c5221f";}})';
  h+='.withFailureHandler(function(e){st.textContent="エラー: "+e.message;st.style.color="#c5221f";})';
  h+='.lookupPostalCode(zip);}';

  h+='function showMsg(t,isErr){var m=document.getElementById("msg");m.textContent=t;m.className="msg "+(isErr?"err":"ok");m.style.display="block";}';
  h+='function doSave(){';
  h+='var name=document.getElementById("name").value.trim();if(!name){showMsg("名前を入力してください",true);return;}';
  h+='var enter=document.getElementById("enterDate").value;if(!enter){showMsg("入社日を入力してください",true);return;}';
  h+='var birth="";if(bSel){var bm=("0"+bSel.month).slice(-2),bd=("0"+bSel.day).slice(-2);birth=bSel.year+"-"+bm+"-"+bd;}';
  h+='var data={shaCD:document.getElementById("shaCD").value.trim(),hpcd:document.getElementById("hpcd").value,gender:document.getElementById("gender").value,name:name,enterDate:enter,birthDate:birth,kubun:document.getElementById("kubun").value,detail:document.getElementById("detail").value,cluster:document.getElementById("cluster").value,wage:document.getElementById("wage").value,workInfo:document.getElementById("workInfo").value,transport:document.getElementById("transport").value,parking:document.getElementById("parking").value,skill:document.getElementById("skill").value,sizeTop:document.getElementById("sizeTop").value,lendCnt:document.getElementById("lendCnt").value,sizeShoe:document.getElementById("sizeShoe").value,carInfo:document.getElementById("carInfo").value,mobile:document.getElementById("mobile").value,emergency:document.getElementById("emergency").value,postcode:document.getElementById("postcode").value,address:document.getElementById("address").value};';
  h+='var btn=document.getElementById("saveBtn");btn.disabled=true;btn.textContent="登録中...";';
  h+='google.script.run.withSuccessHandler(function(msg){showMsg(msg,false);setTimeout(function(){google.script.host.close();},1200);}).withFailureHandler(function(e){showMsg("エラー: "+e.message,true);btn.disabled=false;btn.textContent="登録する";}).saveNewStaff(data);}';
  h+='var _iy=document.getElementById("birthYear").value;if(_iy){syncJp(_iy);openCal(_iy);}';
  h+='<\/script></body></html>';
  return h;
}

function saveNewStaff(d){
  var master=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
  // HPCDの検証（編集可にしたため）と重複チェック
  var hpcdNum=Number(d.hpcd);
  if(d.hpcd===''||d.hpcd==null||isNaN(hpcdNum)){
    throw new Error('HPCDは数字で入力してください（自動採番の候補: '+getNextHpcd()+'）。');
  }
  var all=master.getDataRange().getValues();
  for(var k=1;k<all.length;k++){
    if(Number(all[k][COL_HPCD-1])===hpcdNum){
      var dupName=all[k][COL_NAME-1]||'(名前なし)';
      throw new Error('HPCD '+hpcdNum+' は既に「'+dupName+'」さんで使われています。\n後から登録するこの番号を別の番号に変更してください（次の空き番号: '+getNextHpcd()+'）。');
    }
  }
  var newRow=master.getLastRow()+1;
  var enter=d.enterDate?parseLocalDate(d.enterDate):'';
  var birth=d.birthDate?parseLocalDate(d.birthDate):'';
  master.getRange(newRow,1,1,23).setValues([[
    d.shaCD||'',Number(d.hpcd),d.gender||'',d.name,
    enter||'','',
    d.kubun||'',d.detail||'',birth||'',d.cluster||'',
    d.transport||'',d.parking||'',d.skill||'',
    d.wage?Number(d.wage):'',
    d.workInfo||'',d.sizeTop||'',
    d.lendCnt?Number(d.lendCnt):'',
    d.sizeShoe?Number(d.sizeShoe):'',
    d.carInfo||'',d.mobile||'',d.emergency||'',
    d.postcode||'',d.address||''
  ]]);
  if(enter)master.getRange(newRow,COL_ENTER).setNumberFormat('yyyy/MM/dd');
  if(birth)master.getRange(newRow,COL_BIRTH).setNumberFormat('yyyy/MM/dd');
  if(d.wage)master.getRange(newRow,COL_WAGE).setNumberFormat('\\#,##0');
  master.getRange(newRow,COL_TENURE).setFormula('=IF(E'+newRow+'="","",DATEDIF(E'+newRow+',TODAY(),"Y")&"年"&DATEDIF(E'+newRow+',TODAY(),"YM")&"ヶ月")');
  master.getRange(newRow,COL_AGE).setFormula('=IF(I'+newRow+'="","",DATEDIF(I'+newRow+',TODAY(),"Y"))');
  return d.name+' を登録しました（HPCD: '+d.hpcd+'）';
}

function parseLocalDate(str){
  var p=str.split('-');
  return new Date(Number(p[0]),Number(p[1])-1,Number(p[2]));
}

function getActiveStaffList(){
  var master=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
  if(!master)throw new Error('「'+MASTER_SHEET+'」シートが見つかりません');
  return master.getDataRange().getValues().slice(1)
    .filter(function(r){return r[COL_HPCD-1]&&r[COL_NAME-1]&&!r[COL_RESIGN-1];})
    .map(function(r){return{hpcd:String(r[COL_HPCD-1]),name:String(r[COL_NAME-1])};});
}

function showRetirementDialog(){
  var list=getActiveStaffList();
  if(!list.length){SpreadsheetApp.getUi().alert('在籍スタッフが見つかりません。');return;}
  var today=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd');
  var opts=list.map(function(s){return'<option value="'+s.hpcd+'">'+s.name+'</option>';}).join('');
  var rh='<!DOCTYPE html><html><head><meta charset="UTF-8"><style>';
  rh+='*{box-sizing:border-box;margin:0;padding:0}body{font-family:Meiryo,sans-serif;padding:20px}';
  rh+='h2{font-size:16px;color:#1a73e8;margin-bottom:18px}.f{margin-bottom:14px}';
  rh+='label{display:block;font-size:12px;font-weight:bold;color:#555;margin-bottom:5px}';
  rh+='select,input{width:100%;padding:9px 11px;border:1px solid #dadce0;border-radius:4px;font-size:14px;font-family:Meiryo,sans-serif}';
  rh+='.btns{display:flex;gap:10px;margin-top:20px}.btn{flex:1;padding:10px;border:none;border-radius:4px;font-size:14px;cursor:pointer;font-weight:bold;font-family:Meiryo,sans-serif}';
  rh+='.bc{background:#f1f3f4;color:#3c4043}.bo{background:#1a73e8;color:white}';
  rh+='.msg{display:none;margin-top:14px;padding:8px 12px;border-radius:4px;font-size:12px;font-weight:bold}';
  rh+='.msg.ok{background:#e6f4ea;color:#137333}.msg.err{background:#fce8e6;color:#c5221f}';
  rh+='</style></head><body><h2>退職処理</h2>';
  rh+='<div class="f"><label>退職するスタッフ</label><select id="st">'+opts+'</select></div>';
  rh+='<div class="f"><label>退職日（デフォルト: 本日）</label><input type="date" id="dt" value="'+today+'"></div>';
  rh+='<div id="msg" class="msg"></div>';
  rh+='<div class="btns"><button class="btn bc" onclick="google.script.host.close()">キャンセル</button><button class="btn bo" id="eb" onclick="go()">退職処理を実行</button></div>';
  rh+='<script>';
  rh+='function showMsg(t,isErr){var m=document.getElementById("msg");m.textContent=t;m.className="msg "+(isErr?"err":"ok");m.style.display="block";}';
  rh+='function go(){var s=document.getElementById("st"),hpcd=s.value,name=s.options[s.selectedIndex].text,date=document.getElementById("dt").value;';
  rh+='if(!date){showMsg("退職日を入力してください",true);return;}';
  rh+='if(!confirm(name+" を退職処理します。\\n退職日: "+date.replace(/-/g,"/")+("\\nよろしいですか？")))return;';
  rh+='var b=document.getElementById("eb");b.disabled=true;b.textContent="処理中...";';
  rh+='google.script.run.withSuccessHandler(function(m){showMsg(m,false);setTimeout(function(){google.script.host.close();},1200);}).withFailureHandler(function(e){showMsg("エラー: "+e.message,true);b.disabled=false;b.textContent="退職処理を実行";}).processRetirement(hpcd,date);}';
  rh+='<\/script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(rh).setWidth(370).setHeight(275),'退職処理');
}

function processRetirement(hpcd,dateStr){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),master=ss.getSheetByName(MASTER_SHEET);
  if(!master)throw new Error('「'+MASTER_SHEET+'」シートが見つかりません');
  var data=master.getDataRange().getValues(),targetRow=-1,staffName='';
  for(var i=1;i<data.length;i++){if(String(data[i][COL_HPCD-1])===hpcd){targetRow=i+1;staffName=data[i][COL_NAME-1];break;}}
  if(targetRow===-1)throw new Error('HPCDが見つかりません: '+hpcd);
  var rd=parseLocalDate(dateStr),cell=master.getRange(targetRow,COL_RESIGN);
  cell.setValue(rd);cell.setNumberFormat('yyyy/MM/dd');
  var as=ss.getSheetByName(ACTIVE_SHEET);
  if(as&&!as.getRange(2,1).getFormula()){
    var ad=as.getDataRange().getValues();
    for(var j=ad.length-1;j>=1;j--){if(String(ad[j][COL_HPCD-1])===hpcd){as.deleteRow(j+1);break;}}
  }
  return staffName+' の退職処理が完了しました\n退職日: '+Utilities.formatDate(rd,'Asia/Tokyo','yyyy/MM/dd');
}

function showEditStaffForm(){
  var list=getActiveStaffList();
  if(!list.length){SpreadsheetApp.getUi().alert('在籍スタッフが見つかりません。');return;}
  var h=HtmlService.createHtmlOutput(buildEditFormHtml(list,getFormDefaults())).setWidth(500).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(h,'スタッフ情報変更');
}

function getStaffRecord(hpcd){
  var master=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
  if(!master)throw new Error('「'+MASTER_SHEET+'」シートが見つかりません');
  var data=master.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(String(data[i][COL_HPCD-1])===String(hpcd)){
      var r=data[i],enter=r[COL_ENTER-1],birth=r[COL_BIRTH-1];
      return{
        shaCD:r[COL_SHA_CD-1]||'',gender:r[COL_GENDER-1]||'',name:r[COL_NAME-1]||'',
        enterDate:enter instanceof Date?Utilities.formatDate(enter,'Asia/Tokyo','yyyy-MM-dd'):'',
        birthDate:birth instanceof Date?Utilities.formatDate(birth,'Asia/Tokyo','yyyy-MM-dd'):'',
        kubun:r[COL_KUBUN-1]||'',detail:r[COL_DETAIL-1]||'',cluster:r[COL_CLUSTER-1]||'',
        wage:r[COL_WAGE-1]||'',workInfo:r[COL_WORK_INFO-1]||'',transport:r[COL_TRANSPORT-1]||'',
        parking:r[COL_PARKING-1]||'',skill:r[COL_SKILL-1]||'',sizeTop:r[COL_SIZE_TOP-1]||'',
        lendCnt:r[COL_LEND_CNT-1]||'',sizeShoe:r[COL_SIZE_SHOE-1]||'',carInfo:r[COL_CAR_INFO-1]||'',
        mobile:r[COL_MOBILE-1]||'',emergency:r[COL_EMERGENCY-1]||'',postcode:r[COL_POSTCODE-1]||'',
        address:r[COL_ADDRESS-1]||''
      };
    }
  }
  throw new Error('スタッフが見つかりません: '+hpcd);
}

function updateStaff(hpcd,d){
  var master=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_SHEET);
  if(!master)throw new Error('「'+MASTER_SHEET+'」シートが見つかりません');
  var data=master.getDataRange().getValues(),targetRow=-1;
  for(var i=1;i<data.length;i++){if(String(data[i][COL_HPCD-1])===String(hpcd)){targetRow=i+1;break;}}
  if(targetRow===-1)throw new Error('HPCDが見つかりません: '+hpcd);
  var enter=d.enterDate?parseLocalDate(d.enterDate):null;
  var birth=d.birthDate?parseLocalDate(d.birthDate):null;
  master.getRange(targetRow,COL_SHA_CD).setValue(d.shaCD||'');
  master.getRange(targetRow,COL_GENDER).setValue(d.gender||'');
  master.getRange(targetRow,COL_NAME).setValue(d.name);
  if(enter)master.getRange(targetRow,COL_ENTER).setValue(enter).setNumberFormat('yyyy/MM/dd');
  master.getRange(targetRow,COL_KUBUN).setValue(d.kubun||'');
  master.getRange(targetRow,COL_DETAIL).setValue(d.detail||'');
  if(birth)master.getRange(targetRow,COL_BIRTH).setValue(birth).setNumberFormat('yyyy/MM/dd');
  master.getRange(targetRow,COL_CLUSTER).setValue(d.cluster||'');
  master.getRange(targetRow,COL_TRANSPORT).setValue(d.transport||'');
  master.getRange(targetRow,COL_PARKING).setValue(d.parking||'');
  master.getRange(targetRow,COL_SKILL).setValue(d.skill||'');
  master.getRange(targetRow,COL_WAGE).setValue(d.wage?Number(d.wage):'').setNumberFormat('\\#,##0');
  master.getRange(targetRow,COL_WORK_INFO).setValue(d.workInfo||'');
  master.getRange(targetRow,COL_SIZE_TOP).setValue(d.sizeTop||'');
  master.getRange(targetRow,COL_LEND_CNT).setValue(d.lendCnt?Number(d.lendCnt):'');
  master.getRange(targetRow,COL_SIZE_SHOE).setValue(d.sizeShoe?Number(d.sizeShoe):'');
  master.getRange(targetRow,COL_CAR_INFO).setValue(d.carInfo||'');
  master.getRange(targetRow,COL_MOBILE).setValue(d.mobile||'');
  master.getRange(targetRow,COL_EMERGENCY).setValue(d.emergency||'');
  master.getRange(targetRow,COL_POSTCODE).setValue(d.postcode||'');
  master.getRange(targetRow,COL_ADDRESS).setValue(d.address||'');
  return d.name+' の情報を更新しました';
}

function buildEditFormHtml(list,d){
  var opts='<option value="">選択してください</option>'+list.map(function(s){return'<option value="'+s.hpcd+'">'+s.name+'（'+s.hpcd+'）</option>';}).join('');
  var yo=buildYearOptions(d),so=buildShoeSizeOptions(),h='';
  h+='<!DOCTYPE html><html><head><meta charset="UTF-8"><style>';
  h+='*{box-sizing:border-box;margin:0;padding:0}body{font-family:Meiryo,sans-serif;font-size:13px;color:#202124}';
  h+='.fb{padding:14px;overflow-y:auto;max-height:600px}.sec{margin-bottom:14px}';
  h+='.st{font-size:11px;font-weight:bold;color:#1a73e8;background:#e8f0fe;padding:4px 8px;border-radius:3px;margin-bottom:8px}';
  h+='.gr{display:grid;grid-template-columns:1fr 1fr;gap:7px 12px}.fu{grid-column:1/-1}';
  h+='label,.slb{display:block;font-size:11px;font-weight:bold;color:#555;margin-bottom:3px}';
  h+='.bj{font-size:10px;padding:1px 4px;border-radius:3px;margin-left:3px}';
  h+='.rq{background:#fce8e6;color:#c5221f}.at{background:#e6f4ea;color:#137333}.op{background:#f1f3f4;color:#80868b}';
  h+='input,select{width:100%;padding:6px 8px;border:1px solid #dadce0;border-radius:4px;font-size:13px;font-family:Meiryo,sans-serif}';
  h+='input:focus,select:focus{outline:none;border-color:#1a73e8}';
  h+='.hv{font-size:22px;font-weight:bold;color:#1a73e8;padding:3px 0}';
  h+='.byr{display:grid;grid-template-columns:1fr 1fr;gap:8px}';
  h+='.eg{display:flex;gap:4px;align-items:center}.eg select{width:80px;flex:none}.eg input{width:52px;text-align:center;flex:none}.eg span{white-space:nowrap}';
  h+='.cw{border:1px solid #dadce0;border-radius:6px;padding:8px;background:#f8f9fa;margin-top:8px}';
  h+='.cn{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px}';
  h+='.nb{background:#e8f0fe;color:#1a73e8;border:none;border-radius:4px;padding:3px 12px;cursor:pointer;font-size:15px}';
  h+='.ct{font-size:13px;font-weight:bold;color:#333}';
  h+='.cg{display:grid;grid-template-columns:repeat(7,1fr);gap:1px;text-align:center}';
  h+='.ch{font-size:11px;font-weight:bold;padding:3px 0;color:#666}.ch.su{color:#d93025}.ch.sa{color:#1a73e8}';
  h+='.cd{padding:5px 0;border-radius:50%;font-size:12px}.cd:not(.em){cursor:pointer}.cd:not(.em):hover{background:#e8f0fe}';
  h+='.cd.su{color:#d93025}.cd.sa{color:#1a73e8}.cd.sl{background:#1a73e8!important;color:white!important}.cd.em{cursor:default}';
  h+='.bd{text-align:center;font-size:12px;font-weight:bold;color:#1a73e8;padding-top:6px;min-height:20px}';
  h+='.zr{display:flex;gap:6px}.zr input{flex:1}';
  h+='.zb{padding:6px 10px;background:#e8f0fe;color:#1a73e8;border:1px solid #1a73e8;border-radius:4px;cursor:pointer;font-size:12px;white-space:nowrap}';
  h+='.zs{font-size:11px;margin-top:4px;min-height:16px}';
  h+='.msg{display:none;margin:0 14px 10px;padding:8px 12px;border-radius:4px;font-size:12px;font-weight:bold}';
  h+='.msg.ok{background:#e6f4ea;color:#137333}.msg.err{background:#fce8e6;color:#c5221f}';
  h+='.ft{padding:10px 14px;border-top:1px solid #e0e0e0;display:flex;gap:8px}';
  h+='.bn{flex:1;padding:10px;border:none;border-radius:4px;font-size:13px;font-weight:bold;cursor:pointer;font-family:Meiryo,sans-serif}';
  h+='.bc{background:#f1f3f4;color:#3c4043}.bo{background:#1a73e8;color:white}.bo:disabled{background:#a8c7fa}';
  h+='</style></head><body><div class="fb">';
  h+='<div class="sec"><div class="st">対象スタッフ</div><div class="gr">';
  h+='<div class="fu"><label>スタッフを選択 <span class="bj rq">必須</span></label><select id="pick" onchange="onPick()">'+opts+'</select></div>';
  h+='</div></div>';
  h+='<div id="editArea" style="display:none">';
  h+='<div class="sec"><div class="st">基本情報</div><div class="gr">';
  h+='<div><label>HPCD</label><div class="hv" id="hpcdView">-</div></div>';
  h+='<div><label>社員CD <span class="bj op">任意</span></label><input type="text" id="shaCD"></div>';
  h+='<div class="fu"><label>名前 <span class="bj rq">必須</span></label><input type="text" id="name"></div>';
  h+='<div><label>性別</label><select id="gender"><option value="">選択</option><option>男</option><option>女</option></select></div>';
  h+='<div><label>入社日 <span class="bj rq">必須</span></label><input type="date" id="enterDate"></div>';
  h+='<div class="fu"><label>生年月日</label><div class="byr">';
  h+='<div><div class="slb">西暦</div><select id="birthYear" onchange="onWC()">'+yo+'</select></div>';
  h+='<div><div class="slb">和暦（直接入力可）</div><div class="eg">';
  h+='<select id="eraName" onchange="onEC()"><option value="">元号</option><option value="reiwa">令和</option><option value="heisei">平成</option><option value="showa">昭和</option><option value="taisho">大正</option></select>';
  h+='<input type="number" id="eraYear" min="1" max="99" placeholder="年" onblur="onEC()"><span>年</span></div></div></div>';
  h+='<div id="calSec" style="display:none"><div class="cw"><div class="cn">';
  h+='<button class="nb" type="button" onclick="prevM()">◀</button><span id="calTitle" class="ct"></span><button class="nb" type="button" onclick="nextM()">▶</button>';
  h+='</div><div class="cg" id="calGrid"></div><div class="bd" id="birthDisp"></div></div></div>';
  h+='</div></div></div>';
  h+='<div class="sec"><div class="st">雇用区分・役職</div><div class="gr">';
  h+='<div><label>区分</label><select id="kubun"><option value="">選択</option><option>高校</option><option>大学</option><option>短大</option><option>専門</option><option>主婦</option><option>フリーター</option><option>フリー</option><option>社員</option></select></div>';
  h+='<div><label>詳細</label><input type="text" id="detail"></div>';
  h+='<div><label>クラスター</label><select id="cluster"><option value="">選択</option><option>ルーキー</option><option>1クルー</option><option>2クルー</option><option>3クルー</option><option>1トレーナー</option><option>マネージャー</option></select></div>';
  h+='<div><label>時給（円）</label><input type="number" id="wage"></div>';
  h+='<div class="fu"><label>勤務メモ</label><input type="text" id="workInfo"></div>';
  h+='</div></div>';
  h+='<div class="sec"><div class="st">勤務条件</div><div class="gr">';
  h+='<div><label>交通手段</label><select id="transport"><option value="">選択</option><option>車</option><option>原付</option><option>自転車</option><option>電車</option><option>バス</option><option>歩き</option><option>送迎え</option></select></div>';
  h+='<div><label>駐車場</label><select id="parking"><option value="">なし</option><option value="〇">あり（〇）</option></select></div>';
  h+='<div><label>スキル</label><select id="skill"><option value="">選択</option><option value="D・K">D・K（デシャップ＋キッチン）</option><option value="D">D（デシャップ）</option><option value="K">K（キッチン）</option></select></div>';
  h+='</div></div>';
  h+='<div class="sec"><div class="st">サイズ・貸与品</div><div class="gr">';
  h+='<div><label>上サイズ</label><select id="sizeTop"><option value="">選択</option><option>S</option><option>M</option><option>L</option><option>LL</option><option>3L</option></select></div>';
  h+='<div><label>貸与枚数</label><select id="lendCnt"><option value="">選択</option><option value="1">1枚</option><option value="2">2枚</option><option value="3">3枚</option></select></div>';
  h+='<div><label>靴サイズ（cm）</label><select id="sizeShoe">'+so+'</select></div>';
  h+='<div><label>車情報</label><input type="text" id="carInfo"></div>';
  h+='</div></div>';
  h+='<div class="sec"><div class="st">連絡先・住所</div><div class="gr">';
  h+='<div><label>携帯番号</label><input type="tel" id="mobile"></div>';
  h+='<div><label>緊急連絡先</label><input type="tel" id="emergency"></div>';
  h+='<div class="fu"><label>郵便番号</label><div class="zr"><input type="text" id="postcode" maxlength="8"><button class="zb" type="button" onclick="lookupZip()">住所を検索</button></div></div>';
  h+='<div class="fu"><label>住所</label><input type="text" id="address"><div class="zs" id="zipSt"></div></div>';
  h+='</div></div>';
  h+='</div>';
  h+='</div>';
  h+='<div id="msg" class="msg"></div>';
  h+='<div class="ft"><button class="bn bc" onclick="google.script.host.close()">キャンセル</button><button class="bn bo" id="saveBtn" onclick="doUpdate()" disabled>更新する</button></div>';
  h+='<script>';
  h+='var cY,cM,bSel=null,curHpcd="";';
  h+='function getMD(){var v=document.getElementById("enterDate").value;if(v){var p=v.split("-");return{m:parseInt(p[1]),d:parseInt(p[2])};}var n=new Date();return{m:n.getMonth()+1,d:n.getDate()};}';
  h+='function w2e(y){y=parseInt(y);if(y>=2019)return{code:"reiwa",name:"令和",year:y-2018};if(y>=1989)return{code:"heisei",name:"平成",year:y-1988};if(y>=1926)return{code:"showa",name:"昭和",year:y-1925};return{code:"taisho",name:"大正",year:y-1911};}';
  h+='function e2w(era,eY){eY=parseInt(eY);if(!era||isNaN(eY)||eY<1)return null;if(era==="reiwa")return eY+2018;if(era==="heisei")return eY+1988;if(era==="showa")return eY+1925;if(era==="taisho")return eY+1911;return null;}';
  h+='function syncJp(y){var e=w2e(y);document.getElementById("eraName").value=e.code;document.getElementById("eraYear").value=e.year;}';
  h+='function syncYSel(y){var s=document.getElementById("birthYear");s.value=y;if(parseInt(s.value)===parseInt(y))syncJp(y);}';
  h+='function openCal(y,m,dd){var md=m?{m:m,d:dd}:getMD();cY=parseInt(y);cM=md.m;bSel={year:cY,month:cM,day:md.d};renderCal();document.getElementById("calSec").style.display="block";}';
  h+='function onWC(){var y=parseInt(document.getElementById("birthYear").value);if(!y){document.getElementById("calSec").style.display="none";bSel=null;return;}syncJp(y);openCal(y);}';
  h+='function onEC(){var era=document.getElementById("eraName").value,eY=document.getElementById("eraYear").value,w=e2w(era,eY);if(!w)return;var s=document.getElementById("birthYear");s.value=w;if(parseInt(s.value)===w)openCal(w);}';
  h+='function prevM(){cM--;if(cM<1){cM=12;cY--;syncYSel(cY);}renderCal();}';
  h+='function nextM(){cM++;if(cM>12){cM=1;cY++;syncYSel(cY);}renderCal();}';
  h+='function selDay(day){bSel={year:cY,month:cM,day:day};syncYSel(cY);renderCal();}';
  h+='function renderCal(){';
  h+='document.getElementById("calTitle").textContent=cY+"年 "+cM+"月";';
  h+='var fd=new Date(cY,cM-1,1).getDay(),ld=new Date(cY,cM,0).getDate();';
  h+='var g=document.getElementById("calGrid");g.innerHTML="";';
  h+='var hd=["日","月","火","水","木","金","土"];';
  h+='for(var i=0;i<7;i++){var th=document.createElement("div");th.className="ch"+(i===0?" su":i===6?" sa":"");th.textContent=hd[i];g.appendChild(th);}';
  h+='for(var j=0;j<fd;j++){var em=document.createElement("div");em.className="cd em";g.appendChild(em);}';
  h+='for(var day=1;day<=ld;day++){';
  h+='var dow=(fd+day-1)%7,cl=document.createElement("div");';
  h+='cl.className="cd"+(dow===0?" su":dow===6?" sa":"")+(bSel&&bSel.year===cY&&bSel.month===cM&&bSel.day===day?" sl":"");';
  h+='cl.textContent=day;cl.setAttribute("data-d",day);';
  h+='cl.onclick=function(){selDay(parseInt(this.getAttribute("data-d")));};';
  h+='g.appendChild(cl);}';
  h+='if(bSel){var era=w2e(bSel.year);document.getElementById("birthDisp").textContent=bSel.year+"年"+bSel.month+"月"+bSel.day+"日（"+era.name+era.year+"年）";}}';
  h+='function lookupZip(){';
  h+='var zip=document.getElementById("postcode").value.replace(/[^0-9]/g,""),st=document.getElementById("zipSt");';
  h+='if(zip.length!==7){st.textContent="7桁の郵便番号を入力してください";st.style.color="#c5221f";return;}';
  h+='st.textContent="検索中...";st.style.color="#888";';
  h+='google.script.run';
  h+='.withSuccessHandler(function(addr){if(addr){document.getElementById("address").value=addr;st.textContent="✓ 住所を入力しました（番地・号を追記してください）";st.style.color="#137333";}else{st.textContent="該当する住所が見つかりませんでした";st.style.color="#c5221f";}})';
  h+='.withFailureHandler(function(e){st.textContent="エラー: "+e.message;st.style.color="#c5221f";})';
  h+='.lookupPostalCode(zip);}';
  h+='function showMsg(t,isErr){var m=document.getElementById("msg");m.textContent=t;m.className="msg "+(isErr?"err":"ok");m.style.display="block";}';
  h+='function setVal(id,v){document.getElementById(id).value=v||"";}';
  h+='function onPick(){';
  h+='curHpcd=document.getElementById("pick").value;';
  h+='if(!curHpcd){document.getElementById("editArea").style.display="none";document.getElementById("saveBtn").disabled=true;return;}';
  h+='document.getElementById("saveBtn").disabled=true;';
  h+='google.script.run.withSuccessHandler(function(r){';
  h+='document.getElementById("editArea").style.display="block";';
  h+='document.getElementById("hpcdView").textContent=curHpcd;';
  h+='setVal("shaCD",r.shaCD);setVal("name",r.name);document.getElementById("gender").value=r.gender||"";';
  h+='setVal("enterDate",r.enterDate);document.getElementById("kubun").value=r.kubun||"";setVal("detail",r.detail);';
  h+='document.getElementById("cluster").value=r.cluster||"";setVal("wage",r.wage);setVal("workInfo",r.workInfo);';
  h+='document.getElementById("transport").value=r.transport||"";document.getElementById("parking").value=r.parking||"";';
  h+='document.getElementById("skill").value=r.skill||"";document.getElementById("sizeTop").value=r.sizeTop||"";';
  h+='document.getElementById("lendCnt").value=r.lendCnt?String(r.lendCnt):"";document.getElementById("sizeShoe").value=r.sizeShoe?String(r.sizeShoe):"";';
  h+='setVal("carInfo",r.carInfo);setVal("mobile",r.mobile);setVal("emergency",r.emergency);setVal("postcode",r.postcode);setVal("address",r.address);';
  h+='document.getElementById("zipSt").textContent="";';
  h+='if(r.birthDate){var p=r.birthDate.split("-");syncYSel(Number(p[0]));openCal(Number(p[0]),Number(p[1]),Number(p[2]));}else{document.getElementById("birthYear").value="";document.getElementById("eraName").value="";document.getElementById("eraYear").value="";document.getElementById("calSec").style.display="none";bSel=null;}';
  h+='document.getElementById("saveBtn").disabled=false;';
  h+='}).withFailureHandler(function(e){showMsg("エラー: "+e.message,true);document.getElementById("editArea").style.display="none";}).getStaffRecord(curHpcd);}';
  h+='function doUpdate(){';
  h+='if(!curHpcd){showMsg("スタッフを選択してください",true);return;}';
  h+='var name=document.getElementById("name").value.trim();if(!name){showMsg("名前を入力してください",true);return;}';
  h+='var enter=document.getElementById("enterDate").value;if(!enter){showMsg("入社日を入力してください",true);return;}';
  h+='var birth="";if(bSel){var bm=("0"+bSel.month).slice(-2),bd=("0"+bSel.day).slice(-2);birth=bSel.year+"-"+bm+"-"+bd;}';
  h+='var data={shaCD:document.getElementById("shaCD").value.trim(),gender:document.getElementById("gender").value,name:name,enterDate:enter,birthDate:birth,kubun:document.getElementById("kubun").value,detail:document.getElementById("detail").value,cluster:document.getElementById("cluster").value,wage:document.getElementById("wage").value,workInfo:document.getElementById("workInfo").value,transport:document.getElementById("transport").value,parking:document.getElementById("parking").value,skill:document.getElementById("skill").value,sizeTop:document.getElementById("sizeTop").value,lendCnt:document.getElementById("lendCnt").value,sizeShoe:document.getElementById("sizeShoe").value,carInfo:document.getElementById("carInfo").value,mobile:document.getElementById("mobile").value,emergency:document.getElementById("emergency").value,postcode:document.getElementById("postcode").value,address:document.getElementById("address").value};';
  h+='var btn=document.getElementById("saveBtn");btn.disabled=true;btn.textContent="更新中...";';
  h+='google.script.run.withSuccessHandler(function(msg){showMsg(msg,false);setTimeout(function(){google.script.host.close();},1200);}).withFailureHandler(function(e){showMsg("エラー: "+e.message,true);btn.disabled=false;btn.textContent="更新する";}).updateStaff(curHpcd,data);}';
  h+='<\/script></body></html>';
  return h;
}

function formatForA4(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var master=ss.getSheetByName(MASTER_SHEET);
  var sheet=ss.getSheetByName(ACTIVE_SHEET);
  if(!master){SpreadsheetApp.getUi().alert('「'+MASTER_SHEET+'」シートが見つかりません');return;}
  if(!sheet){SpreadsheetApp.getUi().alert('「'+ACTIVE_SHEET+'」シートが見つかりません');return;}
  function skillLabel(s){if(s==='D')return 'デシャップ';if(s==='K')return 'キッチン';return s||'';}
  var rows=master.getDataRange().getValues().slice(1)
    .filter(function(r){return r[COL_HPCD-1]&&r[COL_NAME-1]&&!r[COL_RESIGN-1];});
  sheet.clearContents();
  sheet.clearFormats();
  var hd=['No.','社員CD','HPCD','名前','性別','入社日','社歴','年齢','生年月日','クラスター','スキル','時給','区分','詳細','交通手段','上サイズ','貸与','靴サイズ','携帯番号'];
  sheet.getRange(1,1,1,hd.length).setValues([hd]);
  var data=rows.map(function(r,i){
    var en=r[COL_ENTER-1];
    var enStr=en instanceof Date?Utilities.formatDate(en,'Asia/Tokyo','yyyy/MM/dd'):(en||'');
    var bt=r[COL_BIRTH-1];
    var btStr=bt instanceof Date?Utilities.formatDate(bt,'Asia/Tokyo','yyyy/MM/dd'):(bt||'');
    var rowNum=i+2;
    var tenureF='=IF(F'+rowNum+'="","",DATEDIF(F'+rowNum+',TODAY(),"Y")&"年"&DATEDIF(F'+rowNum+',TODAY(),"YM")&"ヶ月")';
    var ageF='=IF(I'+rowNum+'="","",DATEDIF(I'+rowNum+',TODAY(),"Y"))';
    return[i+1,r[COL_SHA_CD-1]||'',r[COL_HPCD-1]||'',r[COL_NAME-1]||'',r[COL_GENDER-1]||'',
      enStr,tenureF,ageF,btStr,r[COL_CLUSTER-1]||'',
      skillLabel(r[COL_SKILL-1]),r[COL_WAGE-1]||'',r[COL_KUBUN-1]||'',r[COL_DETAIL-1]||'',r[COL_TRANSPORT-1]||'',
      r[COL_SIZE_TOP-1]||'',r[COL_LEND_CNT-1]||'',r[COL_SIZE_SHOE-1]||'',r[COL_MOBILE-1]||''];
  });
  if(data.length>0)sheet.getRange(2,1,data.length,hd.length).setValues(data);
  var lr=data.length+1,lc=hd.length;
  sheet.setRowHeight(1,26);
  sheet.getRange(1,1,1,lc).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(9).setFontFamily('メイリオ').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(false);
  for(var r=2;r<=lr;r++){
    sheet.setRowHeight(r,20);
    sheet.getRange(r,1,1,lc).setBackground(r%2===0?'#dce8fc':'#ffffff')
      .setFontSize(9).setFontFamily('メイリオ').setVerticalAlignment('middle').setWrap(false);
  }
  if(lr>=1)sheet.getRange(1,1,lr,lc).setBorder(true,true,true,true,true,true,'#9aa0a6',SpreadsheetApp.BorderStyle.SOLID);
  var minW=[30,55,60,80,30,82,58,30,82,68,72,55,52,60,55,40,30,45,105];
  for(var c=1;c<=lc;c++){
    sheet.autoResizeColumn(c);
    if(sheet.getColumnWidth(c)<minW[c-1])sheet.setColumnWidth(c,minW[c-1]);
  }
  protectSheetWarning_(ACTIVE_SHEET);
  SpreadsheetApp.getUi().alert('整形完了！（在籍 '+data.length+' 名）\n\n印刷: ファイル → 印刷\n・用紙: A4・横\n・拡大縮小: ページ幅に合わせる');
}

function formatResignedList(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var master=ss.getSheetByName(MASTER_SHEET);
  var sheet=ss.getSheetByName(RESIGNED_SHEET);
  if(!master){SpreadsheetApp.getUi().alert('「'+MASTER_SHEET+'」シートが見つかりません');return;}
  if(!sheet){SpreadsheetApp.getUi().alert('「'+RESIGNED_SHEET+'」シートが見つかりません');return;}
  function skillLabel(s){if(s==='D')return 'デシャップ';if(s==='K')return 'キッチン';return s||'';}
  var rows=master.getDataRange().getValues().slice(1)
    .filter(function(r){return r[COL_HPCD-1]&&r[COL_NAME-1]&&r[COL_RESIGN-1];});
  sheet.clearContents();
  sheet.clearFormats();
  var hd=['No.','社員CD','HPCD','名前','性別','入社日','退社日','社歴','年齢','生年月日','クラスター','スキル','時給','区分','詳細','交通手段','上サイズ','貸与','靴サイズ','携帯番号'];
  sheet.getRange(1,1,1,hd.length).setValues([hd]);
  var data=rows.map(function(r,i){
    var en=r[COL_ENTER-1];
    var enStr=en instanceof Date?Utilities.formatDate(en,'Asia/Tokyo','yyyy/MM/dd'):(en||'');
    var rs=r[COL_RESIGN-1];
    var rsStr=rs instanceof Date?Utilities.formatDate(rs,'Asia/Tokyo','yyyy/MM/dd'):(rs||'');
    var bt=r[COL_BIRTH-1];
    var btStr=bt instanceof Date?Utilities.formatDate(bt,'Asia/Tokyo','yyyy/MM/dd'):(bt||'');
    var tenure=String(r[COL_TENURE-1]||'');
    if(tenure.charAt(0)==='#')tenure='';
    return[i+1,r[COL_SHA_CD-1]||'',r[COL_HPCD-1]||'',r[COL_NAME-1]||'',r[COL_GENDER-1]||'',
      enStr,rsStr,tenure,r[COL_AGE-1]||'',btStr,r[COL_CLUSTER-1]||'',
      skillLabel(r[COL_SKILL-1]),r[COL_WAGE-1]||'',r[COL_KUBUN-1]||'',r[COL_DETAIL-1]||'',r[COL_TRANSPORT-1]||'',
      r[COL_SIZE_TOP-1]||'',r[COL_LEND_CNT-1]||'',r[COL_SIZE_SHOE-1]||'',r[COL_MOBILE-1]||''];
  });
  if(data.length>0)sheet.getRange(2,1,data.length,hd.length).setValues(data);
  var lr=data.length+1,lc=hd.length;
  sheet.setRowHeight(1,26);
  sheet.getRange(1,1,1,lc).setBackground('#5f6368').setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(9).setFontFamily('メイリオ').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(false);
  for(var r=2;r<=lr;r++){
    sheet.setRowHeight(r,20);
    sheet.getRange(r,1,1,lc).setBackground(r%2===0?'#eeeeee':'#ffffff')
      .setFontSize(9).setFontFamily('メイリオ').setVerticalAlignment('middle').setWrap(false);
  }
  if(lr>=1)sheet.getRange(1,1,lr,lc).setBorder(true,true,true,true,true,true,'#9aa0a6',SpreadsheetApp.BorderStyle.SOLID);
  var minW=[30,55,60,80,30,82,82,58,30,82,68,72,55,52,60,55,40,30,45,105];
  for(var c=1;c<=lc;c++){
    sheet.autoResizeColumn(c);
    if(sheet.getColumnWidth(c)<minW[c-1])sheet.setColumnWidth(c,minW[c-1]);
  }
  protectSheetWarning_(RESIGNED_SHEET);
  SpreadsheetApp.getUi().alert('整形完了！（退職者 '+data.length+' 名）\n\n印刷: ファイル → 印刷\n・用紙: A4・横\n・拡大縮小: ページ幅に合わせる');
}

// ★CHECKLIST_SIZE★ チェックリストのサイズ切替（大＝A4に4枚／中＝6枚／小＝8枚・各枚とも同じ内容の空チェックリスト）
// A4縦・「ページ幅に合わせる」印刷を前提に、列数を増やして枚数を増やす（枚数が多いほど1枚が小さくなる）。
// ここの数値だけ直せば見た目を微調整できます（1枚あたり最大20人を想定）。
function checkListPreset_(sizeKey){
  var P={
    '中':{cols:3, blockRowsN:2, font:10, rowH:24, titleH:28, hdrH:22, blockW:[30,96,44,44]},  // 6枚（3列×2段）
    '小':{cols:4, blockRowsN:2, font:8,  rowH:19, titleH:24, hdrH:18, blockW:[26,82,38,38]}   // 8枚（4列×2段）
  };
  return P[sizeKey]||P['中'];
}

function updateCheckListMedium(){ buildCheckList_('中'); }
function updateCheckListSmall(){ buildCheckList_('小'); }
function formatKentaiCheckList(){ buildCheckList_('中'); }  // 後方互換（既存の呼び出し用）

function buildCheckList_(sizeKey){
  var cfg=checkListPreset_(sizeKey);
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var master=ss.getSheetByName(MASTER_SHEET);
  var sheet=ss.getSheetByName('チェックシート');
  if(!master){SpreadsheetApp.getUi().alert('「'+MASTER_SHEET+'」シートが見つかりません');return;}
  if(!sheet){SpreadsheetApp.getUi().alert('「チェックシート」シートが見つかりません');return;}
  var list=getActiveStaffList();
  if(!list.length){SpreadsheetApp.getUi().alert('在籍スタッフが見つかりません。');return;}
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();
  sheet.clearContents();
  sheet.clearFormats();

  var n=list.length, blockCols=4, gapCols=1, gapRows=2;
  var blockRows=n+2;                       // タイトル＋見出し＋人数
  var sheets=cfg.cols*cfg.blockRowsN;      // 生成する枚数（大4/中6/小8）

  // 枚数分の配置（左上から、cfg.blockRowsN 段 × cfg.cols 列）
  var positions=[];
  for(var br=0;br<cfg.blockRowsN;br++){
    for(var bc=0;bc<cfg.cols;bc++){
      positions.push({row:1+br*(blockRows+gapRows), col:1+bc*(blockCols+gapCols)});
    }
  }

  positions.forEach(function(p){
    sheet.getRange(p.row,p.col,1,blockCols).merge().setValue('チェックリスト')
      .setFontWeight('bold').setFontSize(cfg.font+1).setHorizontalAlignment('center')
      .setVerticalAlignment('middle').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setRowHeight(p.row,cfg.titleH);
    var hd=['No.','名前','受取','提出'];
    sheet.getRange(p.row+1,p.col,1,blockCols).setValues([hd])
      .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center')
      .setFontFamily('メイリオ').setFontSize(cfg.font);
    sheet.setRowHeight(p.row+1,cfg.hdrH);
    var rows=list.map(function(s,i){return[i+1,s.name,'',''];});
    sheet.getRange(p.row+2,p.col,rows.length,blockCols).setValues(rows)
      .setFontFamily('メイリオ').setFontSize(cfg.font).setVerticalAlignment('middle');
    for(var r=0;r<n;r++)sheet.setRowHeight(p.row+2+r,cfg.rowH);
    sheet.getRange(p.row,p.col,blockRows,blockCols)
      .setBorder(true,true,true,true,true,true,'#9aa0a6',SpreadsheetApp.BorderStyle.SOLID);
  });

  // 列幅（枚数分の列＋区切り列）
  for(var bc2=0;bc2<cfg.cols;bc2++){
    var base=1+bc2*(blockCols+gapCols);
    for(var k=0;k<blockCols;k++)sheet.setColumnWidth(base+k,cfg.blockW[k]);
    if(bc2<cfg.cols-1)sheet.setColumnWidth(base+blockCols,14);  // 区切り列
  }

  protectSheetWarning_('チェックシート');
  SpreadsheetApp.getUi().alert('チェックリストを更新しました（在籍'+n+'名 × '+sheets+'枚 / サイズ:'+sizeKey+'）\n\n印刷: ファイル → 印刷\n・用紙: A4・縦\n・拡大縮小: ページ幅に合わせる');
}

function updateInterviewSheet(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var master=ss.getSheetByName(MASTER_SHEET);
  var sheet=ss.getSheetByName('面談シート');
  if(!master){SpreadsheetApp.getUi().alert('「'+MASTER_SHEET+'」シートが見つかりません');return;}
  if(!sheet){SpreadsheetApp.getUi().alert('「面談シート」シートが見つかりません');return;}
  var lastCol=Math.max(sheet.getLastColumn(),1);
  var lastRow=sheet.getLastRow();
  if(lastRow>1){
    var allData=sheet.getRange(2,1,lastRow-1,lastCol).getValues();
    for(var i=allData.length-1;i>=0;i--){
      var row=allData[i];
      var name=String(row[0]||'').trim();
      var restEmpty=row.slice(1).every(function(v){return String(v||'').trim()==='';});
      if(name==='#REF!'&&restEmpty)sheet.deleteRow(i+2);
    }
  }
  var activeList=getActiveStaffList();
  if(!activeList.length){SpreadsheetApp.getUi().alert('在籍スタッフが見つかりません。');return;}
  lastRow=sheet.getLastRow();
  var existingNames={};
  if(lastRow>1){
    sheet.getRange(2,1,lastRow-1,1).getValues().forEach(function(r){
      var v=String(r[0]||'').trim();
      if(v)existingNames[v]=true;
    });
  }
  var toAdd=activeList.filter(function(s){return !existingNames[s.name];});
  if(toAdd.length===0){
    SpreadsheetApp.getUi().alert('新しく追加する在籍者はいません（全員すでに一覧にあります）。');
    return;
  }
  var newRows=toAdd.map(function(s){var r=[s.name];for(var c=1;c<lastCol;c++)r.push('');return r;});
  var startRow=sheet.getLastRow()+1;
  sheet.getRange(startRow,1,newRows.length,lastCol).setValues(newRows);
  if(lastRow>=2){
    sheet.getRange(2,1,1,lastCol).copyTo(sheet.getRange(startRow,1,newRows.length,lastCol),{formatOnly:true});
    var refHeight=sheet.getRowHeight(2);
    for(var r=0;r<newRows.length;r++)sheet.setRowHeight(startRow+r,refHeight);
  }
  SpreadsheetApp.getUi().alert(newRows.length+'名を面談シートに追加しました。\n\n'+toAdd.map(function(s){return'・'+s.name;}).join('\n'));
}

function clearForTemplate(){
  var ui=SpreadsheetApp.getUi();
  var res=ui.alert('テンプレート化の確認','このスプレッドシートの全スタッフデータを削除します。この操作は元に戻せません。よろしいですか？',ui.ButtonSet.YES_NO);
  if(res!==ui.Button.YES)return;
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  function clearRows(sheetName){
    var sh=ss.getSheetByName(sheetName);
    if(!sh)return;
    var lr=sh.getLastRow();
    if(lr>1)sh.getRange(2,1,lr-1,sh.getLastColumn()).clearContent();
  }
  clearRows('スタッフマスタ');
  clearRows('在籍スタッフ');
  clearRows('退職者');
  clearRows('面談シート');
  var chk=ss.getSheetByName('チェックシート');
  if(chk){chk.clearContents();chk.clearFormats();}
  ui.alert('完了しました。スタッフデータを全て削除しました。\n\nこのメニュー「【テンプレート化】全データ削除」は、この後Apps Scriptエディタでコードごと削除して構いません（各店舗に配る前提のテンプレートなので、二度と使わない機能です）。');
}

// ------------------------------------------------------------
//  自動生成シートに「警告のみ」の保護をかける
//  ・手で直接編集しようとすると警告が出る（事故防止）
//  ・メニューのスクリプトは問題なく書き込める（誰が実行してもOK）
//  ・複数人で使う想定のため、完全ロックにはしない
// ------------------------------------------------------------
function protectSheetWarning_(sheetName){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sh)return;
  // 既存のシート保護を一旦すべて解除（重複防止）
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){p.remove();});
  // 警告のみの保護をかけ直す
  sh.protect()
    .setDescription('自動生成シート：直接編集しないでください（「スタッフ管理」メニューの更新で上書きされます）')
    .setWarningOnly(true);
}

// ------------------------------------------------------------
//  自動生成シートを手で直接編集したとき、分かりやすく通知する
//  （シンプルトリガー onEdit は手動編集でのみ発火。
//    メニューのスクリプト書き込みでは発火しないので誤通知しない）
// ------------------------------------------------------------
function onEdit(e){
  if(!e||!e.range)return;
  var name=e.range.getSheet().getName();
  if(name==='在籍スタッフ'||name==='退職者'||name==='チェックシート'){
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '直接入力しても記録に残りません（次の更新で消えます）。変更は上の「スタッフ管理」メニューから行ってください。',
      '⚠️「'+name+'」は自動作成シートです',
      10
    );
  }
}
