<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="报表系统" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="../Images/CssAdmin.css">
<script language="javascript" src="../Script/Flw.js"></script>
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/CustomAjax.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript">
function closead1(){
  $("#pingbiDiv").hide("slow");
  $("#start_date").val("");
  $("#end_date").val("");
  $("#s6partment").show("slow");
}
function showpingbi(){
  $("#s6partment").hide("slow");
  $("#pingbiDiv").show("slow");
}
function closead(){
  $("#ReplyDiv").hide("slow");
}
//处理保存回复
$("#ReplyDiv").ready(function(){
$("#submitSaveEdit").click(function(){
//如果回复内容为空，不提交
 if($('#ReplyText').val()!=''){
  jQuery.get("FlwAjaxFunction.asp", { "key": "update"+$("#Keyword").val(), "FItemid": $("#FEntryID").val(),"ReplyText":$('#ReplyText').val() },
   function(data){
		if(data.indexOf("###")>-1){
			var arryreply=data.split("###");
			if(arryreply[0].length>9)
			  curTd.innerText=arryreply[0].substring(0,8)+"...";
			else
			  curTd.innerText=arryreply[0];
			//实时改变提交后背景颜色
			if($("#Keyword").val()=="T7reply")curTd.parentNode.bgColor="#ffff66";
			else if($("#Keyword").val()=="T8reply")curTd.parentNode.bgColor="#ff99ff";
			else if($("#Keyword").val()=="T9reply")curTd.parentNode.parentNode.removeChild(curTd.parentNode);
		}
		$("#ReplyDiv").hide("slow");
   });
  }else{
    alert("没有回复内容，不需要保存！");
	$("#ReplyDiv").hide("slow");
  }
});
});


var arr = new Array();

//分页
function pageN(){
    arr = new Array();
    for(var i = 0 ; i < pageN.arguments.length ; i++){
        arr[i] = pageN.arguments[i];
    }
	//显示加载页面
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#listDiv').load("6sExecuteDetails.asp #listtable",{
	  showType:'DetailsList',
	  Page:arr[0],
	  s6partment:$('#departhide').val(),
	  start_date:$("#start_date").val(),
	  end_date:$("#end_date").val()
	 },function(response, status, xhr){
	  if (status =="success") {
	    pageNavigation('pageN',arr,response.split('###')[1],response.split('###')[2],'showDiv');
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
	//产生分页导航栏
}
function changedepart(){
  $('#departhide').val($("#s6partment").val());
  pageN(arr[0]);
}
function loadpingbidata(){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#pingbiShowDiv').load("6sExecuteDetails.asp #pingbibiao",{
	  showType:'PingbiList',
	  pbyear:$('#pbyear').val(),
	  pbmonth:$('#pbmonth').val()
	},function(response, status, xhr){
	  if (status =="success") {
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
}
function ShowDetails(depart,sd,ed){
  $("#pingbiDiv").hide("slow");
  $("#s6partment").show("slow");
  $("#start_date").val(sd);
  $("#end_date").val(ed);
  $("#departhide").val(depart);
  pageN(arr[0]);
}
function output(){
  window.open("6sExecuteDetails.asp?print_tag=1&showType=PingbiList&pbmonth="+$('#pbmonth').val()+"&pbyear="+$('#pbyear').val(),"Print","","false");
}
</script>
</HEAD>
<BODY>
<%
'if Instr(session("AdminPurviewFLW"),"|103,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
%>
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden;">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div align="center" style="margin:0 auto; ">
<font color="#FF0000"><strong>6S检查扣分工作平台</strong></font>
<p align="left" style="margin-top:0; margin-bottom:0; ">
<font style="background-color:#ffff66">已经回复</font>&nbsp;
<font style="background-color:#ff99ff">已经确认</font>&nbsp;
<input type="hidden" id="departhide" value="allpart">
<input type="hidden" id="start_date" value="">
<input type="hidden" id="end_date" value="">

<select id="s6partment" onChange="changedepart()" style="font-size:12px;height:15px; width:70px; z-index:1"><option value="allpart">所有部门</option><option value="<% =session("Depart") %>">本部门</option></select>
<input type="button" name="seachbutton" id="seachbutton" onClick="showpingbi()" value="6S评比表" style='HEIGHT: 18px;WIDTH: 65px;font-size:12px;'>
</p>
<div id="ReplyDiv" style="width:590px;height:180px;top:0;left:0;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
<form name="ReplyForm" id="ReplyForm" action="test1.asp">
<table id="ReplyTable" border="0" width="100%" cellspacing="0" cellpadding="1" align="center" bgcolor="black" height="100%">
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td width="60"> 改善人 </td>
 <td width="60">
 <input name="Replyer" type="text" id="Replyer" ></td>
 <td width="60"> 改善日期 </td>
 <td width="60">
 <input name="ReplyDate" type="text" id="ReplyDate" ></td>
 <td width="20" align="right"><img src="../images/close.jpg" onClick="javascript:closead()"></td>
</tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 改善对策 </td>
<td colspan="4">
  <textarea name="ReplyText" id="ReplyText" style="width:500px; height:100px; "></textarea>
</td>
</tr> 
<tr style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%;border-bottom: 1px solid #000000;">
<td valign="bottom" colspan="5" align="center">
<input type="hidden" name="FEntryID" id="FEntryID" value="">
<input type="hidden" name="Keyword" id="Keyword" value="">
&nbsp;<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;"  >
</td>
</tr>
</table>
</form>
</div>
<div id="listDiv"></div>
<div id="showDiv"></div>
<div id="pingbiDiv" style="width:'820px';height:'480px';top:0;left:0;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
<div align="center" style="margin:0 auto; ">
<font color="#FF0000"><strong><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;
<select id="pbyear" onChange="loadpingbidata()" style="font-size:12px;height:15px; width:40px "><option value="2011">2011</option>
<option value="2012">2012</option>
<option value="2013">2013</option>
<option value="2014">2014</option>
</select>
<select id="pbmonth" onChange="loadpingbidata()" style="font-size:12px;height:15px; width:40px ">
<option value="1">1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7">7</option>
<option value="8">8</option>
<option value="9">9</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
</select>
月份6S评比表</strong>&nbsp;<input type="button" name="output" id="output" onClick="output()" value="引出" style='HEIGHT: 18px;WIDTH: 65px;font-size:12px;'></font>
</div>
<div id="pingbiShowDiv"></div>
</div>
<script language="javascript">
arr[0] = 1;
pageN(arr);
</script>
</div>
</BODY>
</HTML>