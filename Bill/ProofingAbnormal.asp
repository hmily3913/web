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
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="../Script/jquery.MultiFile.pack.js"></script>
<script language="javascript" src="../Script/jquery.form.js"></script>
<script language="javascript" src="../Script/jquery.easydrag.js"></script>
<script language="javascript" src="../Script/xheditor-zh-cn.js"></script>
<script language="javascript">
//关闭弹出层
function closead1(){
  $("#addDiv").hide("slow");
}
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
	$('#listDiv').load("ProofingAbnormalDetails.asp #listtable",{
	  showType:'DetailsList',
	  Page:arr[0],
	  seachword:$('#seachword').val()
	 },function(response, status, xhr){
	  if (status =="success") {
	    pageNavigation('pageN',arr,response.split('###')[1],response.split('###')[2],'showDiv');
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
	//产生分页导航栏
}
//双击弹出回复层
var thistd;
//处理添加按钮
function showpadd(obj,sid){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#addShowDiv').load("ProofingAbnormalDetails.asp #AddandEditdiv",{
	  showType:'AddEditShow',
	  detailType:obj,
	  SerialNum:sid
	},function(response, status, xhr){
	  if (status =="success") {
	    $("#addDiv").show("slow");
			$('#addDiv').easydrag(); 
			$("#AbnormalNote").xheditor();
	$("#addDiv").setHandler("formove"); 
		$("input[type=file].multi").MultiFile();
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
}
//处理提交事务
function toSubmit(){
  $.post('ProofingAbnormalDetails.asp?showType=DataProcess',$("#AddForm").serialize(),function(data){
    if(data.indexOf("###")==-1) alert("数据异常，请检查！");
	else pageN(0);
  });
  $("#addDiv").hide("slow");
}
//处理选择打样单号事务
function getInfo(obj){
  if($("#"+obj).val()==''){alert("对应编号不能为空！");return false;}
	$.get("ProofingAbnormalDetails.asp", { showType: "getInfo",detailType: obj, InfoID: $("#"+obj).val() },
	   function(data){
		 if(data.indexOf("###")==-1)alert("对应编号不存在，请检查！");
		 else{
		   if(obj=="ProofingID"){
		   $("#CustomID").val(data.split('###')[2]);
		   $("#CustomRanke").val(data.split('###')[3]);
		   $("#CustomLevel").val(data.split('###')[4]);
		   $("#Agenter").val(data.split('###')[5]);
		   $("#ProductType").val(data.split('###')[6]);
		   $("#Product_td").html(data.split('###')[7]);
		   }
		   else if(obj=="Register"){
		   $("#RegisterName").val(data.split('###')[1]);
		   $("#Department").val(data.split('###')[2]);
		   $("#Departmentname").val(data.split('###')[3]);
		   }
		 }
	   });
}
//选择产品触发
function changePro(){
  $("#ProductType").val($("#Product option:selected").text().split('/')[1]);
}
//图片上传，ajax提交
function uploadPic() { 
    var options = { 
        success:       showResponse  // post-submit callback 
    }; 
    $('#formUpload').ajaxSubmit(options); 
	// post-submit callback 
	function showResponse(responseText, statusText)  { 
		$('#formUpload').html(responseText);
	}
}
//复制图片路径
function CopyPath2Pic(FilePath,FileSize)//
{
  if($('#Pic').val()=="")$('#Pic').val(FilePath);
  else $('#Pic').val($('#Pic').val()+"~"+FilePath);
}
//显示图片
function PicShow(){
  if($('#Pic').val()=="")alert('没有图片');
  else{
    var pici=0;
	var arrpic=$('#Pic').val().split('~');
	for(pici;pici<arrpic.length;pici++)
      $('#PicShowDiv').html("<img src='"+arrpic[pici]+"' />");
	$('#addDiv').hide('slow');
	$('#PicDiv').show('slow');
  }
}
function closead12(){
  $('#PicDiv').hide('slow');
  $('#addDiv').show('slow');
}
</script>
</HEAD>
<BODY>
<%
if Instr(session("AdminPurviewFLW"),"|202,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden;">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div align="center" style="top:0;margin:0 auto; ">
<font color="#FF0000"><strong>打样异常反馈处理进度汇总</strong></font>
<p align="left" style="margin-top:0; margin-bottom:0; ">
<font style="background-color:#ff99ff">已经处理</font>&nbsp;
<input type="text" name="seachword" id="seachword" style='HEIGHT: 18px;WIDTH: 80px;'>
<input type="button" name="seachbutton" id="seachbutton" onClick="pageN(arr)" value="查找" style='HEIGHT: 18px;WIDTH: 40px;'>
<input type="button" name="addbutton" id="button" onClick="showpadd('AddNew','')" value="添加" style='HEIGHT: 18px;WIDTH: 40px;'>
</p>
<div id="addDiv" style="width:100%;height:100%;top:0;left:0;display:none;background-color:#888888;position:absolute;">
<div id="addShowDiv"></div>
</div>
<div id="listDiv"></div>
<div id="showDiv"></div>

<script language="javascript">
arr[0] = 1;
pageN(arr);
</script>
</div>
</BODY>
</HTML>