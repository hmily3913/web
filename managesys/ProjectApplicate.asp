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
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp" -->
<% 
  dim sql,rs
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数
  dim datafrom'数据表名
      datafrom=" z_ProjectApplicate "
  dim datawhere'数据条件
'		 datawhere="where a.fid=b.fid and a.fuser>0 and ftext9!='Y'"
  '获取记录总数
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  idCount=rs("idCount")
  if(idcount>0) then'如果记录总数=0,则不处理
    if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
	  pagec=int(idcount/pages)'获取总页数
   	else
      pagec=int(idcount/pages)+1'获取总页数
    end if
  end if
  rs.close
  set rs=nothing
 %>
<script language="javascript" src="../Script/Flw.js"></script>
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/CustomAjax.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript">
function toSubmitEdit(){
	jQuery.post( "ProjectApplicateDetails.asp", $("#EditForm").serialize(), function(data, textStatus, jqXHR){
	  if (textStatus =="success") {
		if(data.indexOf("@@@")>-1){
		  alert("你没有权限进行此操作！");
		  $("#DetailsAddDiv").hide("slow");}
		else{
		  $("#DetailsAddDiv").hide("slow");
		  pageN(0);
		}
	  }	
	})
}
function toDelete(){
	jQuery.post( "ProjectApplicateDetails.asp", {SerialNum:$("#SerialNum").val(),Keyword:"Delete"}, function(data, textStatus, jqXHR){
	  if (textStatus =="success") {
		if(data.indexOf("@@@")>-1){
		  alert("你没有权限进行此操作！");
		  $("#DetailsAddDiv").hide("slow");}
		else{
		  $("#DetailsAddDiv").hide("slow");
		  pageN(0);
		}
	  }	
	})
}
function ShowEdit(snum){
	$('#DetailsAddDiv').load("ProjectApplicateDetails.asp #EditDiv",{SerialNum:snum,Keyword:"ShowEdit"},function(response, status, xhr){
	  if (status =="success") {
		$('#DetailsAddDiv').show("slow");
		}
	  });
}
function toSubmit(){
	jQuery.post( "ProjectApplicateDetails.asp", $("#AddForm").serialize(), function(data, textStatus, jqXHR){
	  if (textStatus =="success") {
		if(data.indexOf("@@@")>-1){
		  alert("你没有权限进行此操作！");
		  $("#DetailsAddDiv").hide("slow");}
		else{
		  $("#DetailsAddDiv").hide("slow");
		  pageN(0);
		}
	  }	
	})
}
function closead(){
  $("#ReplyDiv").hide("slow");
}
function closead1(){
  $("#DetailsAddDiv").hide("slow");
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

function addnew(){
	$('#DetailsAddDiv').load("ProjectApplicateDetails.asp #AddDiv",{Page:1,Keyword:"ShowAdd"},function(response, status, xhr){
	  if (status =="success") {
		$('#DetailsAddDiv').show("slow");
		}
	  });
}

var arr = new Array();

//分页
function pageN(){
    var arr = new Array();
    for(var i = 0 ; i < pageN.arguments.length ; i++){
        arr[i] = pageN.arguments[i];
    }
	//显示加载页面
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#listDiv').load("ProjectApplicateDetails.asp #listtable",{Page:arr[0],Keyword:"ShowList"},function(response, status, xhr){
	  if (status =="success") {
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
	//产生分页导航栏
    pageNavigation('pageN',arr,<%= pagec %>,<%= idcount %>,'showDiv');
}



</script>
</HEAD>
<BODY>
<%
if Instr(session("AdminPurview"),"|1007,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden;">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div align="center" style="margin:0 auto; ">
<font color="#FF0000"><strong>企业项目申请进度管理</strong></font>
<p align="left" style="margin-top:0; margin-bottom:0; ">
<font style="background-color:#ffff66">已经回复</font>&nbsp;
<font style="background-color:#ff99ff">已经确认</font>&nbsp;
<input type="button" name="addbutton" id="addbutton" onClick="addnew()" value="新增" style='HEIGHT: 18px;WIDTH: 40px;'>
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
<div id="DetailsAddDiv" style="position:fixed !important;position:absolute;top:120;left:10;height:90%; width:100%;z-index:600;overflow-y:auto; overflow-x: auto; display:none"></div>
<div id="popDiv" style="position:absolute;right:0;bottom:0;height:0px;width:0px;border:1px solid red; z-index:800;"> 
<!--<img src="../images/close.jpg" onClick="javascript:closead()">-->
</div>
<script language="javascript">
arr[0] = 1;
pageN(arr);
</script>
</div>
</BODY>
</HTML>