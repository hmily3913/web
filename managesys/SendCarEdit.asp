<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="zbh" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>编辑</TITLE>
<link rel="stylesheet" href="../Images/CssAdmin.css">
<link rel="stylesheet" href="../Images/jqi.css">
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/CustomAjax.js"></script>
<script language="javascript" src="../Script/jquery-impromptu.3.1.js"></script>
<script language="javascript" >
function ShowUseCar(){
  if($('#UseCarFlag').val()=="是"){$('#UseCarDiv').show('show'); $('#UseCarDiv').slideDown('slow');}
  else $('#UseCarDiv').slideUp('slow');
}
function checkcheckNum(obj){
  checkNum(obj);
  document.getElementById("mileage").value=obj.value-document.getElementById("Startemil").value;
}
function checkcheckFullTime(obj){
	var date1=new Date($('#StarteDate').val().replace(/-/g, "/"));  //开始时间
	var date2=new Date(obj.value.replace(/-/g, "/"));    //结束时间
	var date3=date2.getTime()-date1.getTime();  //时间差的毫秒数
	var minutes=Math.floor(date3/(60*1000));
	$('#totalTime').val(minutes);
}
function ShowDiv(){
		var cityOffset = $("#OutPeronName").offset();
		$("#ReplyDiv").css({left:cityOffset.left + "px", top:cityOffset.top + $("#OutPeronName").outerHeight() + "px"}).slideDown("fast");
}
function deleted(obj){
	obj.parentNode.parentNode.removeChild(obj.parentNode);
}
function closead(){
	$("#ReplyDiv").hide("slow");
}
function SaveRow(){
  if(document.ReplyForm.OutPeronD.length === undefined){//没有外出人
    $("#OutPeron").val("");
	$("#OutPeronName").val("");
	$("#CarryMans").val("0");
  }else{//有外出人
    var o="",p="",q=0
    for(var n=1;n<document.ReplyForm.OutPeronD.length;n++){
	  if(document.ReplyForm.OutPeronD[n].value!=''){
	    if(o==""){
		  o=document.ReplyForm.OutPeronD[n].value;
		  p=document.ReplyForm.OutPeronDName[n].value;
		}
		else {
		  o=o+","+document.ReplyForm.OutPeronD[n].value;
		  p=p+","+document.ReplyForm.OutPeronDName[n].value;
		}
		  q++;
	  }
	}
    $("#OutPeron").val(o);
	$("#OutPeronName").val(p);
	$("#CarryMans").val(q);
  }
	$("#ReplyDiv").hide("slow");
}
function toSubmit(obj){
  if($('#FBase1').val()==""){alert("部门不能为空，请检查！");return false;}
  if($('#PlanStarteDate').val()==""){alert("计划出发时间不能为空，请检查！");return false;}
  var submitname=document.getElementById("Keyword");
  var theform=document.getElementById("editForm");
  var checkflag=document.getElementById("CheckFlag").value;
  var snumber=document.getElementById("SerialNum").value;
  var subflag=false;
  switch (obj.value){
    case "保存" :
	  if(checkflag!=2){
		subflag=true;
	  submitname.value="SaveEdit";
		obj.disabled = true; 
		theform.submit();
		}else{
    alert("该单据当前状态不允许此操作！");
  }
	  break;
    case "删除" :
	  if(checkflag==0 && snumber!=''){
		subflag=true;
	  submitname.value="Delete";
		obj.disabled = true; 
		theform.submit();
		}else{
    alert("该单据当前状态不允许此操作！");
  }
	  break;
	case "申请审核-同意" :
	  if(checkflag==0 && snumber!=''){
		subflag=true;
	  submitname.value="check1";
		obj.disabled = true; 
		theform.submit();
		}else{
    alert("该单据当前状态不允许此操作！");
  }
	  break;
	case "申请审核-不同意" :
	  if(checkflag==0 && snumber!=''){
		subflag=true;
	  submitname.value="check_1";
		obj.disabled = true; 
		theform.submit();
		}else{
    alert("该单据当前状态不允许此操作！");
  }
	  break;
	case "车队审核" :
	  if((checkflag==1 && snumber!='')||(checkflag==4 && snumber!='')){
		subflag=true;
	  if(checkflag==1){
			$.prompt('请选择是否同意派车：<br>1、如果同意需先安排车辆信息；<br>2、如果不同意不需要指定车辆信息。',{
				buttons: { 同意: 'check2', 不同意: 'check_2' },
				submit:function(v,m,f){ 
					if(v=='check2'){
						if((document.getElementById("Startemil").value==''||document.getElementById("Planmileage").value=='')&&checkflag==1){alert("出发里程表数或者计划里程数未输入！");return false;}
						if($("#CarNumber").val()==""){alert("车辆不能为空，请检查");return false;}
						submitname.value="check2";
						$.prompt.close();
						obj.disabled = true; 
						theform.submit();
					}else if(v=='check_2'){
						submitname.value="check_2";
						$.prompt.close();
						obj.disabled = true; 
						theform.submit();
					}
					return false; 
				 }
			 });
	  	}else if(checkflag==4){
				if(document.getElementById("Endmil").value==''||document.getElementById("mileage").value==''){alert("回来里程表数或者实际里程数未输入！");return false;}
					submitname.value="check2";
				obj.disabled = true; 
				theform.submit();
			}
		}else{
    alert("该单据当前状态不允许此操作！");
  }
	  break;
	case "门卫审核" :
	  if(checkflag>=1 && snumber!=''){
		subflag=true;
	  submitname.value="check3";
			obj.disabled = true; 
			theform.submit();
		}else{
    alert("该单据当前状态不允许此操作！");
  }
	  break;
	case "反审核" :
	  if(checkflag>=1 && snumber!=''){
		subflag=true;
	  submitname.value="uncheck";
				obj.disabled = true; 
				theform.submit();
			}else{
    alert("该单据当前状态不允许此操作！");
  }
	  break;
	default : 
	  break; 
  }
}
</script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
'if Instr(session("AdminPurview"),"|1003.1,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result,Action,SerialNum,AdminName,UserName
Result=request.QueryString("Result")
Action=request.QueryString("Action")
SerialNum=request("SerialNum")
UserName=session("UserName")
AdminName=session("AdminName")
dim i,j '用于循环的整数
i=0
'定义派车单主表变量
dim RegDate,Register,RegisterName,FBase1,FBase1Name,SendReason,GoodsName,DeliveryAddr,mileage,Remark,TeamSugg
dim FBiller,FBillerName,FDate,PlanStarteDate,PlanEndDate,StarteDate,EndDate,CheckFlag,Checker1,Checker2,CheckDate1,CheckDate2
dim Checker3,CheckDate3,Importance,CarNumber,Driver,DriverName,Fee,FeeDepartment,FeeDepartmentName,DPhone,Startemil,Endmil,Planmileage 
dim totalTime,CarryGoods,CarryMans,UseCarFlag,packages,checkflagname,OutPeron,OutPeronName
call ProcessFun()
if Result="SendCar" then

%>
<div id="ReplyDiv" style="display:none; position:absolute; height:200px; min-width:150px; background-color:white;border:1px solid;overflow-y:auto;overflow-x:auto; z-index:999999">
<form name="ReplyForm" id="ReplyForm" action="test1.asp">
<table id="ReplyTable" border="0" cellspacing="0" cellpadding="1" align="center" bgcolor="black" style="overflow:auto;">
<tbody id="TbDetails" style="overflow:auto;">
<tr style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%;border-bottom: 1px solid #000000;">
<td  bgcolor="#8DB5E9" height="24"  width="60">员工</td>
<td  bgcolor="#8DB5E9">操作</td>
</tr>
<tr height="24" id="CloneNodeTr" onMouseOver = "this.style.backgroundColor = '#FFFFFF'" onMouseOut = "this.style.backgroundColor = '#FFFFFF'" style='background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%;border-bottom: 1px solid #000000;cursor:hand; display:none'>
 <td width="60">
 <input name="OutPeronD" type="hidden" id="OutPeronD">
		<input name="OutPeronDName" type="text" class="textfield" id="OutPeronDName" style="WIDTH: 60;" value="" maxlength="100" onBlur="getEmpName(this)"></td>
 <td width="20" align="right" onClick="deleted(this)"><img src="../images/close.jpg"/></td>
</tr>
</tbody>
<tr style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%;border-bottom: 1px solid #000000;">
<td valign="bottom" colspan="2" align="center">
&nbsp;<input style='HEIGHT: 18px;WIDTH: 40px;' name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="关闭"  onClick="closead()">
&nbsp;<input style='HEIGHT: 18px;WIDTH: 40px;' name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="增加"   onClick="AddRow()">
&nbsp;<input style='HEIGHT: 18px;WIDTH: 40px;' name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="确定"  onClick="SaveRow()">
</td>
</tr>
</table>
</form>
</div>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>外出单查看：添加，修改，删除宿舍人员信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="SendCarEdit.asp?Result=SendCar&Action=Add" onClick='changeAdminFlag("添加宿舍人员信息")'>添加外出单信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SendCarMana.asp?Result=Search&Page=1" onClick='changeAdminFlag("外出单列表")'>查看所有外出单信息</a>&nbsp;|&nbsp;</font><a href="javascript:history.go(-1)" >返回</a></td>
  </tr>
</table>
<br>
  <form name="editForm" id="editForm" method="post" action="SendCarEdit.asp?Result=SendCar&Action=<%=Action%>&SerialNum=<%=SerialNum%>">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table border="0" cellpadding="0" cellspacing="0" id=editNews >

      <tr>
        <td height="20" align="left" width="100">单据号：</td>
        <td  width="120">
		<input name="SerialNum" type="text" class="textfield" id="SerialNum" style="WIDTH: 120;" value="<%=SerialNum%>" maxlength="100" readonly="true" alt="系统自动生成"></td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td  width="120">&nbsp;</td>
        <td width="120" height="20" align="left" >&nbsp;</td>
        <td width="120" >&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="left">是否派车：</td>
        <td>
		<select name="UseCarFlag" id="UseCarFlag"  <%if CheckFlag>=1 then response.Write("readonly") else response.Write("onChange='ShowUseCar()'") end if%>>
		<option value="是" <%if UseCarFlag="是" then response.write ("selected")%>>是</option>
		<option value="否" <%if UseCarFlag="否" then response.write ("selected")%>>否</option>
		</select></td>
        <td height="20" align="left">申请人：</td>
        <td>
		<input name="Register" type="hidden" id="Register" value="<%=Register%>">
		<input name="RegisterName" type="text" class="textfield" id="RegisterName" style="WIDTH: 120;" value="<%=RegisterName%>" maxlength="100" onBlur="getEmpName(this)" <%if CheckFlag>=1 then response.Write("readonly") end if%> alt="由申请人填写">
		</td>
        <td height="20" align="left">申请日期：</td>
        <td><input name="RegDate" type="text" class="textfield" id="RegDate" style="WIDTH: 120;" value="<%=RegDate%>" maxlength="100" onBlur="return checkDate(this)" <%if CheckFlag>=1 then response.Write("readonly") end if%> alt="由申请人填写"></td>
      </tr>
      <tr>
        <td height="20" align="left">申请部门：</td>
        <td><input name="FBase1" type="hidden" id="FBase1" value="<%=FBase1%>">
		<input name="FBase1name" type="text" class="textfield" id="FBase1name" style="WIDTH: 120;" value="<%=FBase1name%>" maxlength="100" onBlur="return getDepartment(this)" <%if CheckFlag>=1 then response.Write("readonly") end if%> alt="由申请人填写"></td>
        <td height="20" align="left">计划出发时间：</td>
        <td><input name="PlanStarteDate" type="text" class="textfield" id="PlanStarteDate" style="WIDTH: 120;" value="<%=PlanStarteDate%>" maxlength="100" onBlur="return checkFullTime(this)" <%if CheckFlag>=1 then response.Write("readonly") end if%> alt="由申请人填写"></td>
        <td height="20" align="left">计划回来时间：</td>
        <td><input name="PlanEndDate" type="text" class="textfield" id="PlanEndDate" style="WIDTH: 120;" value="<%=PlanEndDate%>" maxlength="100" onBlur="return checkFullTime(this)" ></td>
      </tr>
      <tr>
        <td height="20" align="left">外出分类：</td>
        <td>
		<select name="SendReason" id="SendReason"  <%if CheckFlag>=1 then response.Write("readonly") end if%>>
		<option value="提货" <%if SendReason="提货" then response.write ("selected")%>>1、提货</option>
		<option value="送货" <%if SendReason="送货" then response.write ("selected")%>>2、送货</option>
		<option value="外协" <%if SendReason="外协" then response.write ("selected")%>>3、外协</option>
		<option value="接客人" <%if SendReason="接客人" then response.write ("selected")%>>4、接客人</option>
		<option value="送客人" <%if SendReason="送客人" then response.write ("selected")%>>5、送客人</option>
		<option value="接员工" <%if SendReason="接员工" then response.write ("selected")%>>6、接员工</option>
		<option value="送员工" <%if SendReason="送员工" then response.write ("selected")%>>7、送员工</option>
		<option value="公出" <%if SendReason="公出" then response.write ("selected")%>>8、公出</option>
		<option value="私出" <%if SendReason="私出" then response.write ("selected")%>>9、私出</option>
		</select></td>
        <td height="20" align="left">重要性：</td>
        <td>
		<select name="Importance" id="Importance"  <%if CheckFlag>=1 then response.Write("readonly") end if%>>
		<option value="重要紧急" <%if Importance="重要紧急" then response.write ("selected")%>>重要紧急</option>
		<option value="重要不紧急" <%if Importance="重要不紧急" then response.write ("selected")%>>重要不紧急</option>
		<option value="紧急不重要" <%if Importance="紧急不重要" then response.write ("selected")%>>紧急不重要</option>
		<option value="一般性" <%if Importance="一般性" then response.write ("selected")%>>一般性</option>
		</select></td>
        <td height="20" align="left">事由及内容：</td>
        <td ><input name="GoodsName" type="text" class="textfield" id="GoodsName" style="WIDTH: 120;" value="<%=GoodsName%>" maxlength="200" <%if CheckFlag>=1 then response.Write("readonly") end if%> alt="由申请人填写"></td>
      </tr>
      <tr>
        <td height="20" align="left">目的地及联系人：</td>
        <td colspan="3"><input name="DeliveryAddr" type="text" class="textfield" id="DeliveryAddr" style="WIDTH: 350;" value="<%=DeliveryAddr%>" maxlength="200" <%if CheckFlag>=1 then response.Write("readonly") end if%> alt="由申请人填写,必须详细，具体,准确"></td>
        <td height="20" align="left">联系方式：</td>
        <td ><input name="DPhone" type="text" class="textfield" id="DPhone" style="WIDTH: 120;" value="<%=DPhone%>" maxlength="100" <%if CheckFlag>=1 then response.Write("readonly") end if%> alt="由申请人填写"></td>
      </tr>
      <tr>
        <td height="20" align="left">外出人：</td>
        <td width="120"><input type="hidden" name="OutPeron" id="OutPeron" value="<%=OutPeron%>"><input name="OutPeronName" type="text" class="textfield" id="OutPeronName" style="WIDTH: 120;" value="<%=OutPeronName%>" maxlength="100" <%if CheckFlag<1 then %>onFocus="ShowDiv()" <% end if%> readonly></td>
        <td height="20" align="left" width="100">外出人数/个：</td>
        <td width="120"><input name="CarryMans" type="text" class="textfield" id="CarryMans" style="WIDTH: 120;" value="<%=CarryMans%>" maxlength="100" readonly></td>
        <td height="20" align="left" width="100"></td>
        <td width="120"></td>
      </tr>
   </table>
<div id="UseCarDiv"  <% 
if UseCarFlag="否" then  response.Write("style='display:none '")
 %>>
  <table>
      <tr>
        <td height="20" align="left" width="100">本次载货量/立方：</td>
        <td width="120"><input name="CarryGoods" type="text" class="textfield" id="CarryGoods" style="WIDTH: 120;" value="<%=CarryGoods%>" maxlength="100" onBlur="return checkNum(this)" alt="由申请人填写"></td>
        <td width="100" height="20" align="left">货物件数：</td>
        <td width="120"><input name="packages" type="text" class="textfield" id="packages" style="WIDTH: 120;" value="<%=packages%>" maxlength="100" onBlur="return checkNum(this)"></td>
        <td height="20" align="left" width="100">&nbsp;</td>
        <td width="120">&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="left" width="100">车牌号：</td>
        <td width="120">
		<select name="CarNumber" id="CarNumber" onChange="getDriver(this)">
		<option value=""></option>
    <%
	if CheckFlag>=1 and UseCarFlag="是" then
		dim temsql,temrs
		set temrs = server.createobject("adodb.recordset")
	  temsql="select carid from z_Car"
		response.Write(temsql)
	  temrs.open temsql,connk3,1,1
	  while (not temrs.eof)
			response.Write("<option value="""&temrs("carid")&"""")
			if CarNumber=temrs("carid") then response.write ("selected")
			response.Write(">"&temrs("carid")&"</option>")
			temrs.movenext
		wend
		temrs.close
		set temrs=nothing
	end if
		%>
		</select></td>
        <td height="20" align="left" width="100">驾驶员：</td>
        <td width="120">
		<input name="Driver" type="hidden" id="Driver" value="<%=Driver%>">
		<input name="DriverName" type="text" class="textfield" id="DriverName" style="WIDTH: 120;" value="<%=DriverName%>" maxlength="10" onBlur="getEmpName(this)" alt="由车队主管或驾驶员填写"></td>
        <td height="20" align="left" width="100">预计里程数：</td>
        <td width="120"><input name="Planmileage" type="text" class="textfield" id="Planmileage" style="WIDTH: 120;" value="<%=Planmileage%>" maxlength="100" onChange="return checkNum(this)" alt="由车队主管或驾驶员填写"></td>
      </tr>
      <tr>
        <td height="20" align="left">出发里程表数：</td>
        <td><input name="Startemil" type="text" class="textfield" id="Startemil" style="WIDTH: 120;" value="<%=Startemil%>" maxlength="100" onBlur="return checkNum(this)" alt="由车队主管或驾驶员填写,要完整,准确"></td>
        <td height="20" align="left">回来里程表数：</td>
        <td><input name="Endmil" type="text" class="textfield" id="Endmil" style="WIDTH: 120;" value="<%=Endmil%>" maxlength="100" onBlur="return checkcheckNum(this)" alt="由车队主管或驾驶员填写,要完整,准确"></td>
        <td height="20" align="left">实际里程数：</td>
        <td><input name="mileage" type="text" class="textfield" id="mileage" style="WIDTH: 120;" value="<%=mileage%>" maxlength="100" onChange="return checkNum(this)" readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left">费用：</td>
        <td><input name="Fee" type="text" class="textfield" id="Fee" style="WIDTH: 120;" value="<%=Fee%>" maxlength="100" onChange="return checkNum(this)" alt="由车队主管或驾驶员填写"></td>
        <td height="20" align="left">费用归属单位：</td>
        <td>
		<input name="FeeDepartment" type="hidden" id="FeeDepartment" value="<%=FeeDepartment%>">
		<input name="FeeDepartmentName" type="text" class="textfield" id="FeeDepartmentName" style="WIDTH: 120;" value="<%=FeeDepartmentName%>" maxlength="50" onBlur="return getDepartment(this)"></td>
        <td height="20" align="left">车队意见：</td>
        <td ><input name="TeamSugg" type="text" class="textfield" id="TeamSugg" style="WIDTH: 120;" value="<%=TeamSugg%>" maxlength="200"  alt="由车队主管或驾驶员填写"></td>
      </tr>
	</table>
</div>
	<table>
      <tr>
        <td height="20" align="left" width="100">实际出发时间：</td>
        <td><input name="StarteDate" type="text" class="textfield" id="StarteDate" style="WIDTH: 120;" value="<%=StarteDate%>" maxlength="100" onBlur="return checkFullTime(this)" alt="由门卫填写,要完整,准确"></td>
        <td height="20" align="left" width="100">实际回来时间：</td>
        <td><input name="EndDate" type="text" class="textfield" id="EndDate" style="WIDTH: 120;" value="<%=EndDate%>" maxlength="100" onBlur="return checkcheckFullTime(this)" alt="由门卫填写,要完整,准确"></td>
        <td height="20" align="left" width="100">累计使用时间/分：</td>
        <td><input name="totalTime" type="text" class="textfield" id="totalTime" style="WIDTH: 120;" value="<%=totalTime%>" maxlength="100" onChange="return checkNum(this)" readonly></td>
      </tr>
      <tr>
        <td height="20" align="left" width="100">备注：</td>
        <td colspan="5"><input name="Remark" type="text" class="textfield" id="Remark" style="WIDTH: 500;" value="<%=Remark%>" maxlength="500"></td>
      </tr>
      <tr>
        <td height="20" align="left">制单人：</td>
        <td width="120"><input type="hidden" name="FBiller" id="FBiller" value="<%=FBiller%>">
		<input name="FBillerName" type="text" class="textfield" id="FBillerName" style="WIDTH: 120;" value="<%=FBillerName%>" maxlength="100" readonly></td>
        <td height="20" align="left">制单日期：</td>
        <td width="120"><input name="FDate" type="text" class="textfield" id="FDate" style="WIDTH: 120;" value="<%=FDate%>" maxlength="100" readonly></td>
        <td height="20" align="left"></td>
        <td width="120"></td>
      </tr>
      <tr>
        <td height="20" align="left">审核状态：</td>
        <td><input id="CheckFlag" name="CheckFlag" type="hidden" value="<%=CheckFlag%>"><input name="checkflagname" type="text" class="textfield" id="checkflagname" style="WIDTH: 120;" value="<%=checkflagname%>" maxlength="100" readonly="true"></td>
        <td height="20" align="left">一审人名：</td>
        <td><input name="Checker1" type="text" class="textfield" id="Checker1" style="WIDTH: 120;" value="<%=Checker1%>" maxlength="100" readonly="true"></td>
        <td height="20" align="left">一审时间：</td>
        <td><input name="CheckDate1" type="text" class="textfield" id="CheckDate1" style="WIDTH: 120;" value="<%=CheckDate1%>" maxlength="100" readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left"></td>
        <td></td>
        <td height="20" align="left">二审人名：</td>
        <td><input name="Checker2" type="text" class="textfield" id="Checker2" style="WIDTH: 120;" value="<%=Checker2%>" maxlength="100" readonly="true"></td>
        <td height="20" align="left">二审时间：</td>
        <td><input name="CheckDate2" type="text" class="textfield" id="CheckDate2" style="WIDTH: 120;" value="<%=CheckDate2%>" maxlength="100" readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left"></td>
        <td></td>
        <td height="20" align="left">三审人名：</td>
        <td><input name="Checker3" type="text" class="textfield" id="Checker3" style="WIDTH: 120;" value="<%=Checker3%>" maxlength="100" readonly="true"></td>
        <td height="20" align="left">三审时间：</td>
        <td><input name="CheckDate3" type="text" class="textfield" id="CheckDate3" style="WIDTH: 120;" value="<%=CheckDate3%>" maxlength="100" readonly="true"></td>
      </tr>

 
      <tr>
        <td height="20" align="left" colspan="3">&nbsp;</td>
        <td valign="bottom" colspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td valign="bottom" colspan="6" align="center">
		<input type="hidden" name="Keyword" id="Keyword" value="">
<% If (FBiller=UserName or Register=UserName) and CheckFlag=0 Then %>		&nbsp;<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;"  onClick="toSubmit(this)">
		&nbsp;<input name="delete" type="button" class="button"  id="delete" value="删除" style="WIDTH: 80;"  onClick="toSubmit(this)"><% end if %>
<% If (Instr(session("AdminPurview"),"|1003.5,")>0 or Instr(session("AdminPurview"),"|1003.9,")>0) and CheckFlag=0 Then %>		&nbsp;<input name="check1" type="button" class="button"  id="check1" value="申请审核-同意" style="WIDTH: 120;" onClick="toSubmit(this)" >&nbsp;<input name="check1" type="button" class="button"  id="check1" value="申请审核-不同意" style="WIDTH: 120;" onClick="toSubmit(this)" ><% end if %>
<% If Instr(session("AdminPurview"),"|1003.2,")>0 and (CheckFlag=1 or CheckFlag=4) and UseCarFlag="是" Then %>		&nbsp;<input name="check1" type="button" class="button"  id="check1" value="车队审核" style="WIDTH: 80;" onClick="toSubmit(this)" ><% end if %>
<% If Instr(session("AdminPurview"),"|1003.3,")>0 and (CheckFlag>0 and CheckFlag<4) Then %>		&nbsp;<input name="check2" type="button" class="button"  id="check2" value="门卫审核" style="WIDTH: 80;"  onClick="toSubmit(this)"> <% end if %>
<% If (Instr(session("AdminPurview"),"|1003.2,")>0 or Instr(session("AdminPurview"),"|1003.3,")>0 or Instr(session("AdminPurview"),"|1003.5,")>0) and CheckFlag>0 Then %>		&nbsp;<input name="check2" type="button" class="button"  id="check2" value="反审核" style="WIDTH: 80;"  onClick="toSubmit(this)"><% end if %>
		</td>
      </tr>
      <tr>
        <td height="20" align="left" colspan="3">&nbsp;</td>
        <td valign="bottom" colspan="3">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
  </form>
<%
end if
%>
</BODY>
</HTML>

<%
sub ProcessFun()
  dim Keyword,rsRepeat,rs,sql
  Keyword=request("Keyword")
  if Keyword="SaveEdit" then '保存事务处理
	  if Action="Add" then '增加记录
		  '主表信息添加
		  set rs = server.createobject("adodb.recordset")
		  sql="select * from z_SendCar"
		  rs.open sql,connk3,1,3
		  rs.addnew
		  rs("RegDate")=Request.Form("RegDate")
		  rs("Register")=trim(Request.Form("Register"))
		  rs("RegistDepartment")=Request.Form("FBase1")
		  rs("SendReason")=Request.Form("SendReason")
		  rs("GoodsName")=Request.Form("GoodsName")
		  rs("DeliveryAddr")=Request.Form("DeliveryAddr")
		  rs("mileage")=Request.Form("mileage")
		  rs("FBiller")=Request.Form("FBiller")
		  rs("FDate")=Request.Form("FDate")
		  rs("CarryGoods")=Request.Form("CarryGoods")
		  rs("CarryMans")=Request.Form("CarryMans")
		  rs("PlanStarteDate")=Request.Form("PlanStarteDate")
		  rs("PlanEndDate")=Request.Form("PlanEndDate")
		  rs("Remark")=Request.Form("Remark")
		  rs("Importance")=Request.Form("Importance")
		  rs("DPhone")=Request.Form("DPhone")
		  rs("UseCarFlag")=Request.Form("UseCarFlag")
		  rs("packages")=Request.Form("packages")
		  rs("OutPeron")=Request.Form("OutPeron")
		  
		  rs.update
		  rs.close
		  set rs=nothing 
		response.write "<script language=javascript> alert('成功增加派车单信息！');changeAdminFlag('派车单信息');location.replace('SendCarMana.asp?Result=Search&Page=1');</script>"
	  end if
	  if Action="Modify" then '修改记录
	  	'保存主表信息编辑
		  set rs = server.createobject("adodb.recordset")
		  sql="select * from z_SendCar where checkflag <2 and SerialNum="& SerialNum
		  rs.open sql,connk3,1,3
		  if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		  end if
		  rs("RegDate")=Request.Form("RegDate")
		  rs("Register")=trim(Request.Form("Register"))
		  rs("RegistDepartment")=Request.Form("FBase1")
		  rs("SendReason")=Request.Form("SendReason")
		  rs("GoodsName")=Request.Form("GoodsName")
		  rs("DeliveryAddr")=Request.Form("DeliveryAddr")
		  rs("mileage")=Request.Form("mileage")
		  rs("FBiller")=Request.Form("FBiller")
		  rs("FDate")=Request.Form("FDate")
		  rs("CarryGoods")=Request.Form("CarryGoods")
		  rs("CarryMans")=Request.Form("CarryMans")
		  rs("PlanStarteDate")=Request.Form("PlanStarteDate")
		  if Request.Form("PlanEndDate")<>"" then
		  rs("PlanEndDate")=Request.Form("PlanEndDate")
		  end if
		  rs("Remark")=Request.Form("Remark")
		  rs("Importance")=Request.Form("Importance")
		  rs("DPhone")=Request.Form("DPhone")
		  rs("UseCarFlag")=Request.Form("UseCarFlag")
		  rs("packages")=Request.Form("packages")
		  rs("OutPeron")=Request.Form("OutPeron")
		  rs.update
		  rs.close
		  set rs=nothing 
		response.write "<script language=javascript> alert('成功编辑派车信息！');changeAdminFlag('派车列表');location.replace('SendCarMana.asp?Result=Search&Page=1');</script>"
	  end if
  elseif Keyword="Delete" then
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_SendCar where checkflag =0 and SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  else
	    if rs("FBiller")=UserName then' or Instr(session("AdminPurview"),"|1003.4,")>0
		  rs("DeleteFlag")=1
		  rs.update
'		  sql="delete from z_SendCar where SerialNum="& SerialNum
'		  connk3.execute(sql)
		else
		  response.write ("只能删除自己登记的派车单！")
		  response.end
		end if
	  end if
	  rs.close
	  set rs=nothing 
	response.write "<script language=javascript> alert('派车信息删除成功！');changeAdminFlag('派车列表');location.replace('SendCarMana.asp?Result=Search&Page=1');</script>"
  elseif Keyword="check1" then
  '申请部门主管审核
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_SendCar where checkflag =0 and SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
		if len(rs("OutPeron"))>0 then
			dim rs2
			set rs2=connk3.Execute("select  a.FNumber,a.Name from HM_Employees a,HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber  like '%"&rs("OutPeron")&"%'")
			if ((not rs2.eof) and Instr(session("AdminPurview"),"|1003.9,")>0) or (rs2.eof and Instr(session("AdminPurview"),"|1003.5,")>0) then
				rs("RejecteFlag")=0
				rs("CheckFlag")=1
				rs("Checker1")=UserName
				rs("CheckDate1")=now()
			else
				response.write ("你没有权限进行此操作！")
				response.end
			end if
		elseif Instr(session("AdminPurview"),"|1003.5,")>0 then
			rs("RejecteFlag")=0
			rs("CheckFlag")=1
			rs("Checker1")=UserName
			rs("CheckDate1")=now()
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
	  rs.update
	  rs.close
	  set rs=nothing 
	response.write "<script language=javascript> alert('派车信息审核成功！');changeAdminFlag('派车列表');location.replace('SendCarMana.asp?Result=Search&Page=1');</script>"
  elseif Keyword="check_1" and Instr(session("AdminPurview"),"|1003.5,")>0 then
  '申请部门主管审核
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_SendCar where checkflag =0 and SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
	    rs("RejecteFlag")=1
		rs("CheckFlag")=0
		rs("Checker1")=UserName
		rs("CheckDate1")=now()
	  rs.update
	  rs.close
	  set rs=nothing 
	response.write "<script language=javascript> alert('派车信息审核成功！');changeAdminFlag('派车列表');location.replace('SendCarMana.asp?Result=Search&Page=1');</script>"
  elseif Keyword="check2" and Instr(session("AdminPurview"),"|1003.2,")>0 then
  '车队审核
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_SendCar where SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
	  if Request.Form("CarNumber")="" then
		response.write ("车牌号不能为空，审核失败！")
		response.end
	  end if
	  if rs("CheckFlag")=1 then
	    rs("CheckFlag")=2
	    rs("CarNumber")=Request.Form("CarNumber")
	    rs("Driver")=Request.Form("Driver")
		rs("FeeDepartment")=Request.Form("FeeDepartment")
		rs("TeamSugg")=Request.Form("TeamSugg")
		rs("Startemil")=Request.Form("Startemil")
		rs("Planmileage")=Request.Form("Planmileage")
		if Request.Form("PlanEndDate")<>"" then rs("PlanEndDate")=Request.Form("PlanEndDate")
		rs("Checker2")=UserName
		rs("CheckDate2")=now()
	  rs.update
	  elseif rs("CheckFlag")=4 then
	    rs("CheckFlag")=5
		rs("TeamSugg")=Request.Form("TeamSugg")
		rs("Endmil")=Request.Form("Endmil")
		rs("mileage")=Request.Form("mileage")
		rs("Fee")=Request.Form("Fee")
		rs("Checker2")=UserName
		rs("CheckDate2")=now()
	  rs.update
	    if Request.Form("Endmil")>0 then
		connk3.Execute("update z_Car set mileageNum="&Request.Form("Endmil")&" where CarID='"&rs("CarNumber")&"'")
		end if
	  end if
	  rs.close
	  set rs=nothing 
	response.write "<script language=javascript> alert('派车信息审核成功！');changeAdminFlag('派车列表');location.replace('SendCarMana.asp?Result=Search&Page=1');</script>"
  elseif Keyword="check_2" and Instr(session("AdminPurview"),"|1003.2,")>0 then
  '车队驳回审核
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_SendCar where SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
	  if rs("CheckFlag")=1 and rs("UseCarFlag")="是" then
	    rs("CheckFlag")=0
		rs("RejecteFlag")=1
		rs("TeamSugg")=Request.Form("TeamSugg")
		rs("Checker2")=UserName
		rs("CheckDate2")=now()
	    rs.update
	  end if
	  rs.close
	  set rs=nothing 
	response.write "<script language=javascript> alert('派车信息驳回成功！');changeAdminFlag('派车列表');location.replace('SendCarMana.asp?Result=Search&Page=1');</script>"
  elseif Keyword="check3" and Instr(session("AdminPurview"),"|1003.3,")>0 then
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_SendCar where checkflag >=1 and SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
	  if rs("CheckFlag")=1 and rs("UseCarFlag")="否" then
	    rs("CheckFlag")=3
	    rs("StarteDate")=Request.Form("StarteDate")
		rs("Checker3")=UserName
		rs("CheckDate3")=now()
	    rs.update
	  elseif rs("CheckFlag")=2 then
	    rs("CheckFlag")=3
	    rs("StarteDate")=Request.Form("StarteDate")
		rs("Checker3")=UserName
		rs("CheckDate3")=now()
	  rs.update
	    if rs("UseCarFlag")="否" then
		else
		connk3.Execute("update z_Car set CarSatus='出行',CarryMans=CarryMans-"&rs("CarryMans")&",CarryGoods=CarryGoods-"&rs("CarryGoods")&" where CarID='"&rs("CarNumber")&"'")
		end if
	  elseif rs("CheckFlag")=3 then
		'回来门卫审核
	    if rs("UseCarFlag")="否" then
				rs("CheckFlag")=5
			else
				rs("CheckFlag")=4
			end if
			rs("totalTime")=Request.Form("totalTime")
	    rs("EndDate")=Request.Form("EndDate")
			rs("Checker3")=UserName
			rs("CheckDate3")=now()
			rs("Endmil")=Request.Form("Endmil")
			rs("mileage")=Request.Form("mileage")
			rs.update
			'根据外出人影响考勤系统-edit by zbh 2011-12-07
			if rs("OutPeron")<>"" and rs("SendReason")<>"私出" then
				OutPeron="'"&replace(rs("OutPeron"),",","','")&"'"
				sql="select date,a.userid,a.name as xm,h.deptname,a.ssn,c.num_runid,c.name,c.units,d.sdays,d.edays,e.schclassid,e.schName,  "
				sql=sql&"e.starttime,e.endtime,e.checkin,e.checkout,e.checkintime1,e.checkintime2,e.checkouttime1,e.checkouttime2,e.workday  "
				sql=sql&"from USERINFO a,USER_OF_RUN b,NUM_RUN c,NUM_RUN_DEIL d,SchClass e ,Calendar,DEPARTMENTS h "
				sql=sql&"where a.userid=b.userid and b.num_of_run_id=c.num_runid and a.defaultdeptid=h.deptid  "
				sql=sql&"and c.num_runid=d.num_runid and d.schclassid=e.schclassid and b.startdate<=date and b.enddate>=date and a.ssn in ("&OutPeron&") "
				sql=sql&"and ((DATEPART(weekday,date)-1=d.sdays%7 and c.units=1)  "
				sql=sql&"or ((day(date)+(datediff(m,c.startdate,date)%c.cyle)*31)=d.sdays and c.units=2))  "
				sql=sql&"and datediff(d,date,'"&rs("StarteDate")&"')<=0 and datediff(d,date,'"&rs("EndDate")&"')>=0 "
				sql=sql&"order by a.defaultdeptid,a.userid,date,e.starttime "
				set rs2=server.createobject("adodb.recordset")
				rs2.open sql,ConnStrkq,1,1
				while(not rs2.eof)
					dim sql2,rs3
					dim flagbool:flagbool=false
					if rs2("checkin")=1 then
						sql2="select checktime from CHECKINOUT where checktime>'"&rs2("date")&"'+' '+convert(varchar(8),'"&rs2("checkIntime1")&"',114) and checktime<'"&rs2("date")&"'+' '+convert(varchar(8),'"&rs2("starttime")&"',114) and userid="&rs2("userid")
						sql2=sql2&" union select startSpecday as checktime from USER_SPEDAY where datediff(d,startspecday,'"&rs2("date")&"')=0 and startspecday<='"&rs2("date")&"'+' '+convert(varchar(8),'"&rs2("endtime")&"',114) and endspecday>='"&rs2("date")&"'+' '+convert(varchar(8),'"&rs2("starttime")&"',114) and userid="&rs2("userid")
						set rs3=server.createobject("adodb.recordset")
						rs3.open sql2,Connkq,1,1
						if rs3.eof and rs3.bof and datediff("s",rs("StarteDate"),rs2("date")&" "&left(rs2("starttime"),12))>=0 and datediff("s",rs("EndDate"),rs2("date")&" "&left(rs2("starttime"),12))<=0 then flagbool=true
					end if
					if rs2("checkout")=1 then
						sql2="select checktime from CHECKINOUT where checktime>'"&rs2("date")&"'+' '+convert(varchar(8),'"&rs2("endtime")&"',114) and checktime<'"&rs2("date")&"'+' '+convert(varchar(8),'"&rs2("checkouttime2")&"',114) and userid="&rs2("userid")
						sql2=sql2&" union select startSpecday as checktime from USER_SPEDAY where datediff(d,startspecday,'"&rs2("date")&"')=0 and startspecday<='"&rs2("date")&"'+' '+convert(varchar(8),'"&rs2("endtime")&"',114) and endspecday>='"&rs2("date")&"'+' '+convert(varchar(8),'"&rs2("starttime")&"',114) and userid="&rs2("userid")
						set rs3=server.createobject("adodb.recordset")
						rs3.open sql2,Connkq,1,1
						if rs3.eof and rs3.bof and datediff("s",rs("StarteDate"),rs2("date")&" "&left(rs2("endtime"),12))>=0 and datediff("s",rs("EndDate"),rs2("date")&" "&left(rs2("endtime"),12))<=0 then flagbool=true
					end if
					if flagbool then
						'写入到考勤系统
						connkq.Execute("insert into USER_SPEDAY select "&rs2("userid")&",'"&rs2("date")&" "&left(rs2("starttime"),12)&"','"&rs2("date")&" "&left(rs2("endtime"),12)&"',5,'资讯平台外出单审核通过，单号："&rs("SerialNum")&"',getdate(),'z_SendCar',"&rs("SerialNum")&" ")
						'外出单考勤写入标志修改，方便后期查询
						rs("ISCQ")=1
						rs.update
					end if
					rs2.movenext
				wend
				rs2.close
				set rs2=nothing
			end if
			'更新车辆信息
	    if rs("UseCarFlag")="否" then
			else
		  dim tmcar:tmcar=rs("CarNumber")
		  if Request.Form("Endmil")>0 then
			  connk3.Execute("update z_Car set mileageNum="&Request.Form("Endmil")&",CarryMans=CarryMans+"&rs("CarryMans")&",CarryGoods=CarryGoods+"&rs("CarryGoods")&" where CarID='"&tmcar&"'")
		  else
			  connk3.Execute("update z_Car set CarryMans=CarryMans+"&rs("CarryMans")&",CarryGoods=CarryGoods+"&rs("CarryGoods")&" where CarID='"&tmcar&"'")
		  end if
			'更新车辆状态
		  set rs = server.createobject("adodb.recordset")
		  sql="select * from z_SendCar where checkflag =3 and UseCarFlag='是' and CarNumber='"& tmcar&"'"
		  rs.open sql,connk3,0,1
		  if rs.bof and rs.eof then
				connk3.Execute("update z_Car set CarSatus='空闲' where CarID='"&tmcar&"'")
		  end if
		end if
	  end if
	  rs.close
	  set rs=nothing 
	response.write "<script language=javascript> alert('派车信息审核成功！');changeAdminFlag('派车列表');location.replace('SendCarMana.asp?Result=Search&Page=1');</script>"
  elseif Keyword="uncheck" then
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_SendCar where checkflag >=1 and SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
	  if rs("CheckFlag")=5 then
	    if rs("UseCarFlag")="否" and Instr(session("AdminPurview"),"|1003.3,")>0 and rs("Checker3")=UserName then
				rs("CheckFlag")=3
				rs("Checker3")=UserName
				rs("CheckDate3")=now()
				if rs("ISCQ") then
					connkq.Execute("Delete from USER_SPEDAY where billsrcType='z_SendCar' and BillSrcNum="&SerialNum)
					rs("ISCQ")=0
				end if
			elseif rs("UseCarFlag")="是" and Instr(session("AdminPurview"),"|1003.2,")>0 then
				rs("CheckFlag")=4
				rs("Checker2")=UserName
				rs("CheckDate2")=now()
			end if
			rs.update
	  elseif rs("CheckFlag")=4 and Instr(session("AdminPurview"),"|1003.3,")>0 then
			rs("CheckFlag")=3
			rs("Checker3")=UserName
			rs("CheckDate3")=now()
			if rs("ISCQ") then
				connkq.Execute("Delete from USER_SPEDAY where billsrcType='z_SendCar' and BillSrcNum="&SerialNum)
				rs("ISCQ")=0
			end if
			rs.update
			connk3.Execute("update z_Car set CarSatus='出行',CarryMans=CarryMans-"&rs("CarryMans")&",CarryGoods=CarryGoods-"&rs("CarryGoods")&" where CarID='"&rs("CarNumber")&"'")
	  elseif rs("CheckFlag")=3 then
	    if rs("UseCarFlag")="否" and Instr(session("AdminPurview"),"|1003.3,")>0 then
	    rs("CheckFlag")=1
		rs("Checker3")=UserName
		rs("CheckDate3")=now()
		rs.update
		elseif rs("UseCarFlag")="是" and Instr(session("AdminPurview"),"|1003.3,")>0 then
	    rs("CheckFlag")=2
		rs("Checker3")=UserName
		rs("CheckDate3")=now()
		rs.update
		tmcar=rs("CarNumber")
	    connk3.Execute("update z_Car set CarryMans=CarryMans+"&rs("CarryMans")&",CarryGoods=CarryGoods+"&rs("CarryGoods")&" where CarID='"&tmcar&"'")
		  set rs = server.createobject("adodb.recordset")
		  sql="select * from z_SendCar where checkflag =3 and UseCarFlag='是' and CarNumber='"&tmcar&"'"
		  rs.open sql,connk3,0,1
		  if rs.bof and rs.eof then
			connk3.Execute("update z_Car set CarSatus='空闲' where CarID='"&tmcar&"'")
		  end if
		end if
	  elseif rs("CheckFlag")=2 and Instr(session("AdminPurview"),"|1003.2,")>0 then
		rs("CheckFlag")=1
		rs("Checker2")=UserName
		rs("CheckDate2")=now()
	  rs.update
	  elseif rs("CheckFlag")=1 and Instr(session("AdminPurview"),"|1003.5,")>0  and rs("Checker1")=UserName then
		rs("CheckFlag")=0
		rs("Checker1")=UserName
		rs("CheckDate1")=now()
	  rs.update
	  else
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
	  rs.close
	  set rs=nothing 
	response.write "<script language=javascript> alert('派车信息审核成功！');changeAdminFlag('派车列表');location.replace('SendCarMana.asp?Result=Search&Page=1');</script>"
  else
  	if Action="Modify" then'提出编辑信息
	  '提取主表信息
	  set rs = server.createobject("adodb.recordset")
      sql="select * from z_SendCar where SerialNum="& SerialNum
      rs.open sql,connk3,1,1
      if rs.bof and rs.eof then
        response.write ("数据库读取记录出错！")
        response.end
      end if
	  RegDate=rs("RegDate")
	  Register=rs("Register")
	  RegisterName=getUser(rs("Register"))
	  FBase1=rs("RegistDepartment")
	  FBase1Name=getDepartment(rs("RegistDepartment"))
	  SendReason=rs("SendReason")
	  GoodsName=rs("GoodsName")
	  UseCarFlag=rs("UseCarFlag")
	  DeliveryAddr=rs("DeliveryAddr")
      mileage=rs("mileage")
	  Remark=rs("Remark")
      FBiller=rs("FBiller")
	  FBillerName=getUser(rs("FBiller"))
      FDate=rs("FDate")
	  PlanStarteDate=rs("PlanStarteDate")
	  PlanEndDate=rs("PlanEndDate")
	  StarteDate=rs("StarteDate")
	  EndDate=rs("EndDate")
	  CheckFlag=rs("CheckFlag")
	  CarryGoods=rs("CarryGoods")
	  CarryMans=rs("CarryMans")
	  totalTime=rs("totalTime")
	  packages=rs("packages")
	  if rs("CheckFlag")=2 or (rs("CheckFlag")=1 and rs("UseCarFlag")="否") then
	  StarteDate=now()
	  end if
	  if rs("CheckFlag")=3 and Instr(session("AdminPurview"),"|1003.3,")>0 then
	  EndDate=now()
	  totalTime=datediff("n",StarteDate,now())
	  end if
	  CheckFlag=rs("CheckFlag")
	  if rs("CheckFlag")=1 then
	  checkflagname="主管审核"
	  elseif rs("CheckFlag")=2 then
	  checkflagname="车队一审"
	  elseif rs("CheckFlag")=3 then
	  checkflagname="门卫一审"
	  elseif rs("CheckFlag")=4 then
	  checkflagname="门卫二审"
	  elseif rs("CheckFlag")=5 then
	  checkflagname="已结案"
	  else
	  checkflagname="未审核"
	  end if
	  Checker1=getUser(rs("Checker1"))
	  Checker2=getUser(rs("Checker2"))
	  CheckDate1=rs("CheckDate1")
	  CheckDate2=rs("CheckDate2")
	  Checker3=getUser(rs("Checker3"))
	  CheckDate3=rs("CheckDate3")
	  Importance=rs("Importance")
	  CarNumber=rs("CarNumber")
	  Driver=rs("Driver")
	  DriverName=getUser(rs("Driver"))
	  Fee=rs("Fee")
	  TeamSugg=rs("TeamSugg")
	  FeeDepartment=rs("FeeDepartment")
	  DPhone=rs("DPhone")
	  Startemil=rs("Startemil")
	  Endmil=rs("Endmil")
	  Planmileage=rs("Planmileage")
	  OutPeron=rs("OutPeron")
	  if OutPeron<>"" then
	  dim iii
	  iii=0
	  while (iii<=UBound(split(OutPeron,",")))
	  if iii<>UBound(split(OutPeron,",")) then
	  OutPeronName=OutPeronName+getUser(split(OutPeron,",")(iii))+","
	  else
	  OutPeronName=OutPeronName+getUser(split(OutPeron,",")(iii))
	  end if
	  iii=iii+1
	  wend
	  end if
	  if rs("FeeDepartment")<>"" then FeeDepartmentName=getDepartment(rs("FeeDepartment"))
	  rs.close
      set rs=nothing 
	else'提取增加时所需信息,制单人，制单日期，单号
	  RegDate=date()
	  Register=UserName
	  RegisterName=AdminName
      FBiller=UserName
	  FBillerName=AdminName
	  mileage="0"
	  GoodsName=""
	  DeliveryAddr=""
	  Remark=""
      FDate=now()
	  if (Hour(now())+6)<=17 then
	  PlanStarteDate=dateadd("h",6,now())
	  PlanEndDate=dateadd("h",6,now())
	  else
	  PlanStarteDate=dateadd("h",18,now())
	  PlanEndDate=dateadd("h",18,now())
	  end if
	  CarryGoods="0"
	  CarryMans="0"
	  totalTime="0"
	  Fee="0"
	  Startemil="0"
	  Endmil="0"
	  Planmileage="0"
	  packages="0"
'	  StarteDate=now()
'	  EndDate=now()
	  CheckFlag=0
	  
	end if
  end if
end sub

Function getUser(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From t_emp where fnumber='"&ID&"'"
  rs.open sql,connk3,1,1
  if rs.bof and rs.eof then
  getUser=""
  else
  getUser=rs("Fname")
  end if
  rs.close
  set rs=nothing
End Function    
Function getDepartment(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From t_item where fitemclassid=2 and Fitemid="&ID
  rs.open sql,connk3,1,1
  getDepartment=rs("Fname")
  rs.close
  set rs=nothing
End Function    
%>