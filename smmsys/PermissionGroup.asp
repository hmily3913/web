<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="zbh" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>员工列表</TITLE>
<link rel="stylesheet" href="../Images/CssAdmin.css">
<link rel="stylesheet" href="../Images/zTreeStyle.css">
<style type="text/css">
	.zhengt
	{
		float: left;
		height: 400px;
		width: 19%;
		display:block;
	}
/* ------------- 右键菜单 -----------------  */

div#rMenu {
	background-color:#555555;
	text-align: left;
	padding:2px;
}

div#rMenu ul {
	margin:1px 0;
	padding:0 5px;
	cursor: pointer;
	list-style: none outside none;
	background-color:#DFDFDF;
}
div#rMenu ul li {
	margin:0;
	padding:2px 0;
}
</style>

<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="../Script/jquery.ztree-2.6.min.js"></script>
<script type="text/javascript">
//组别操作按钮触发	
function ShowAdd(obj){
  $("#DetailTypeName").val(obj);
  if(obj=='Edit' ||obj=='unForbid' || obj=='Forbid' || obj=='Delete'){
    //检查删除，编辑时是否多选或者未选
    if($("#selectgroup option:selected").length!=1){alert('请选择1条记录，再进行此操作！');return false;}
    $("#GroupName").val($("#selectgroup option:selected").text());
    $("#GroupID").val($("#selectgroup option:selected").val());
	//删除或禁止时直接提交
	if(obj=='unForbid' ||obj=='Forbid' || obj=='Delete'){
	  if (confirm("确定要执行此操作吗？"))toSave();
	}
  }
  //添加，编辑时显示编辑框
  if (obj=='Edit' ||obj=='AddNew'){
    $("#ReplyDiv").show("slow");
    $("#GroupName").focus();
  }
}
//提交组别操作
function toSave(){
  if($("#GroupName").val()==""){alert("组别名称不能为空！");return false;}
  $.post('PermissionGroupDetails.asp?showType=DataProcess',{
    detailType:$("#DetailTypeName").val(),
	GroupID:$("#GroupID").val(),
    GroupName:$("#GroupName").val()
	},function(data){
    if(data.indexOf("###")==-1) alert("数据异常，请检查！");
	else{
	 if($("#DetailTypeName").val()=='AddNew')$("#selectgroup").append("<option value='"+data.split('###')[1]+"' style='color:#0000FF'>"+data.split('###')[2]+"</option>");
	 else if($("#DetailTypeName").val()=='Edit')$("#selectgroup option:selected").text($("#GroupName").val());
	 else if($("#DetailTypeName").val()=='Delete')$("#selectgroup option:selected").remove();
	 else if($("#DetailTypeName").val()=='Forbid')$("#selectgroup option:selected").css('color','#FF0000');
	 else if($("#DetailTypeName").val()=='unForbid')$("#selectgroup option:selected").css('color','#0000FF');
	}
  });
  $("#ReplyDiv").hide("slow");
}
//点选组别，显示已有人员名单
	var zTree2;

	var setting2 = {
		expandSpeed: "",
		showLine: true,
		checkable: true,
		isSimpleData : true,
		treeNodeParentKey : "PSNum",
		treeNodeKey : "SerialNum",
		checkType : {"Y":"ps", "N":"ps"},
		callback: {
			rightClick: zTreeOnRightClick
		}
	};

	function searchTreeNode(evt) {
		hideRMenu();
		$('#SearchID').val('');
		$("#editOne").show();
		$("#editOne").css({"top":evt.clientY+"px", "left":evt.clientX+"px", "visibility":"visible"});
	}
	function submitSearch(){
		if($('#SearchID').val()!=''){
			var treeNode = zTree2.getNodeByParam("name", $('#SearchID').val());
			treeNode = treeNode?treeNode:zTree2.getNodeByParam("SerialNum", $('#SearchID').val());
			if (treeNode) {
				zTree2.selectNode(treeNode);
			} else {
				alert("没有找到匹配的节点，请更换搜索条件");
			}
		}
	}
	//右键菜单
	function zTreeOnRightClick(event, treeId, treeNode) {
		if (!treeNode && event.target.tagName.toLowerCase() != "button" && $(event.target).parents("a").length == 0) {
			zTree2.cancelSelectedNode();
			showRMenu("root", event.clientX, event.clientY);
		} else if (treeNode && !treeNode.noR) {
			zTree2.selectNode(treeNode);
			showRMenu("node", event.clientX, event.clientY);
		}
	}
	function showRMenu(type, x, y) {
		$("#rMenu ul").show();
		if (type=="root") {
			$("#m_edit").hide();
			$("#m_del").hide();
			$("#m_check").hide();
			$("#m_unCheck").hide();
		}
		$("#rMenu").css({"top":y+"px", "left":x+"px", "visibility":"visible"});
	}
	function hideRMenu() {
		if (rMenu) rMenu.style.visibility = "hidden";
	}

function showUser(){
  if($("#selectgroup option:selected").length==1){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
    $.post('PermissionGroupDetails.asp?showType=getInfo',{
      detailType:'getUser',
	  GroupID:$("#selectgroup option:selected").val()
	  },function(data){
	    if(data.indexOf("###")>-1){
		  $("#selectold option").remove();
		  $("#selectold").append(data.split('###')[1].split('@@@')[0]);
			var zNodes2 = jQuery.parseJSON(data.split('###')[1].split('@@@')[1]);
			zTree2 = $("#Users").zTree(setting2, zNodes2);
			hideRMenu();
//		  $("#selectnew option").remove();
//		  $("#selectnew").append(data.split('###')[1].split('@@@')[1]);
		  }
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  });
  }
}
//添加指定人员
function AddMove(){
	var tmp = zTree2.getCheckedNodes();
  if(tmp.length>0){
    if($("#selectgroup option:selected").length==1){
	  var userIDs="";
		var olditems="";
	  var i=0;
		for (i; i<tmp.length-1; i++) {
			if(tmp[i].PSNum!='0'&&tmp[i].PSNum!==null){
				userIDs+=tmp[i].SerialNum+",";
				olditems+="<option value='"+tmp[i].SerialNum+"'>"+tmp[i].name+"</option>";
			}
		}
	  userIDs+=tmp[i].SerialNum;
		olditems+="<option value='"+tmp[i].SerialNum+"'>"+tmp[i].name+"</option>";
	  
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
		$.post('PermissionGroupDetails.asp?showType=DataProcess',{
		detailType:'AddUser',
	  GroupID:$("#selectgroup option:selected").val(),
	  UserIDs:userIDs
	  },function(data){
	    if(data.indexOf("###")>-1){
			for (i=0; i<tmp.length; i++) {
				if(tmp[i].PSNum!='0'&&tmp[i].PSNum!==null){
					zTree2.removeNode(tmp[i]);
				}
			}
				zTree2.removeNode(tmp);
			$("#selectold").append(olditems);
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
		}
	  });
	}else{
	  alert('请选择一个权限组进行组员维护！');
	  return false;
	}
  }
}
//移除指定人员
function DeleteMove(){
  if($("#selectold option:selected").length>0){
    if($("#selectgroup option:selected").length==1){
	  var userIDs="";
	  var i=0;
	  for(i;i<$("#selectold option:selected").length-1;i++){
	    userIDs+=$("#selectold option:selected")[i].value+",";
	  }
	  userIDs+=$("#selectold option:selected")[i].value;
	  
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
      $.post('PermissionGroupDetails.asp?showType=DataProcess',{
      detailType:'DeleteUser',
	  GroupID:$("#selectgroup option:selected").val(),
	  UserIDs:userIDs
	  },function(data){
	    if(data.indexOf("###")>-1){
			var olditems = $("#selectold option:selected").remove();
//			$("#selectnew").append(olditems);
			var zNodes2 = jQuery.parseJSON(data.split('###')[2]);
			zTree2 = $("#Users").zTree(setting2, zNodes2);
			hideRMenu();
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
		}
	  });
	}else{
	  alert('请选择一个权限组进行组员维护！');
	  return false;
	}
  }
}
//全部添加
function AddMoveAll(){
    if($("#selectgroup option:selected").length==1){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
      $.post('PermissionGroupDetails.asp?showType=DataProcess',{
      detailType:'AddAllUser',
	  GroupID:$("#selectgroup option:selected").val()
	  },function(data){
	    if(data.indexOf("###")>-1){
			var olditems = $("#selectnew option").remove();
			$("#selectold").append(olditems);
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
		}
	  });
	}else{
	  alert('请选择一个权限组进行组员维护！');
	  return false;
	}
}
//全部删除
function DeleteMoveAll(){
    if($("#selectgroup option:selected").length==1){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
      $.post('PermissionGroupDetails.asp?showType=DataProcess',{
      detailType:'DeleteAllUser',
	  GroupID:$("#selectgroup option:selected").val()
	  },function(data){
	    if(data.indexOf("###")>-1){
			var olditems = $("#selectold option").remove();
			$("#selectnew").append(olditems);
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
		}
	  });
	}else{
	  alert('请选择一个权限组进行组员维护！');
	  return false;
	}
}
var PurviewType;
var GroupID;
var zTree1;
var setting;
var nodes;
function ShowPm(obj){
  GroupID=$("#selectgroup option:selected").val();
  if($("#selectgroup option:selected").length==1){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#addShowDiv').load("PermissionGroupDetails.asp #AddandEditdiv",{
	  showType:'AddEditShow',
	  GroupID:GroupID
	},function(response, status, xhr){
	  if (status =="success") {
	    $("#groupdiv").slideUp("normal");
	    $("#addDiv").show("slow");
			
nodes = [];
PurviewType=obj=='SetPm'?'RPT':'FLW';
	setting = {
		checkable: true,
		checkType : {"Y":"s", "N":"s"},
		expandSpeed:"",
		keepParent: false,
		keepLeaf: false,
		isSimpleData : true,
		treeNodeParentKey : "PSNum",
		treeNodeKey : "SerialNum",
		rootPID : 0,
		async: true,
		asyncUrl: "PermissionGroupDetails.asp",
		asyncParamOther : {"showType":"getInfo","detailType":"SetPurview","PurviewType":PurviewType,"EmpId":GroupID}
	};
		setting.async = true;
		zTree1 = $("#PermissionRPT").zTree(setting, nodes);

		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
	}else{
	  alert('请选择一个权限组进行组员维护！');
	  return false;
	}
}
function closead1(){
  $("#addDiv").hide("slow");
	$("#groupdiv").slideDown('normal');
}
function submit4pm(){
  $.post('PermissionGroupDetails.asp?showType=DataProcess',$("#editForm").serialize(),function(data){
    if(data.indexOf("###")==-1) alert("数据异常，请检查！");
	  $("#addDiv").hide("slow");
	 $("#groupdiv").slideDown('normal');
  });
}

</script>
<script language="javascript">

	function submitSave(){
		var tmp = zTree1.getCheckedNodes();
		var pall="";
		for (var i=0; i<tmp.length; i++) {
			pall+=tmp[i].PermissionID;
		}
		$.post('PermissionGroupDetails.asp?showType=DataProcess&detailType=SetEdit',{
		  PurviewType:PurviewType,
		  PermissionID:pall,
		  GroupID:GroupID
		},function(data){
		  alert(data);
		});
	}

</script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|1102,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden; display:none;">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div id="addDiv" style="width:'820px';height:'480px';top:0;left:0;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
<div id="addShowDiv"></div>
</div>
<font color="#FF0000"><strong>权限组别设置：</strong></font>
<input type="button" value="增加" onClick="ShowAdd('AddNew')" style="width: 40px; height:16px; font-size:12px" />&nbsp;
<input type="button" value="修改" onClick="ShowAdd('Edit')" style="width: 40px; height:16px; font-size:12px" />&nbsp;
<input type="button" value="删除" onClick="ShowAdd('Delete')" style="width: 40px; height:16px; font-size:12px" />&nbsp;
<input type="button" value="禁用" onClick="ShowAdd('Forbid')" style="width: 40px; height:16px; font-size:12px" />&nbsp;
<input type="button" value="启用" onClick="ShowAdd('unForbid')" style="width: 40px; height:16px; font-size:12px" />&nbsp;
<input type="button" value="报表权限" onClick="ShowPm('SetPm')" style="width: 80px; height:16px; font-size:12px" />&nbsp;
<input type="button" value="工作流权限" onClick="ShowPm('SetFLWPm')" style="width: 100px; height:16px; font-size:12px" />&nbsp;
<br />
<div id="groupdiv">
<select id="selectgroup" multiple="multiple" class="zhengt" onChange="showUser()">
<%
dim rs,sql'sql语句
sql="select * from smmsys_PermissionGroup order by SerialNum "
set rs=server.createobject("adodb.recordset")
rs.open sql,connzxpt,0,1
while (not rs.eof)
  dim forbidcl:forbidcl="#0000FF"
  if rs("ForbidFlag")=1 then forbidcl="#FF0000"
  Response.Write "<option value='"&rs("SerialNum")&"' style='color:"&forbidcl&"'>"&rs("GroupName")&"</option>"
  rs.movenext
wend
rs.close
set rs=nothing
%>
</select>
<div class="zhengt">
<div id="ReplyDiv" style="display:none; ">
<input type="hidden" id="DetailTypeName" value="">
<input type="hidden" id="GroupID" value="">
<input name="GroupName" type="text" id="GroupName" style="WIDTH: 100%;" value="" maxlength="50">
<br /><input type="button" onClick="toSave()" value="<-确定" style="width: 100%" />
</div>
</div>
<select id="selectold" multiple="multiple" class="zhengt">
</select>
<div class="zhengt">
	<input type="button" onClick="DeleteMove()" value=">" style="width: 100%" /><br />
	<input type="button" onClick="AddMove()" value="<" style="width: 100%" /><br />
<!--	<input type="button" onClick="DeleteMoveAll()" value=">>" style="width: 100%" /><br />
	<input type="button" onClick="AddMoveAll()" value="<<" style="width: 100%" /><br />-->
</div>
<div id="selectnew" class="zhengt">
<ul id="Users" class="tree"></ul>
</div>
</div>
<div id="rMenu" style="position:absolute; visibility:hidden;z-index:700;">
<li>
<ul id="m_add" onclick="searchTreeNode(event);"><li>查找</li></ul>
</li>
</div>
<div id="editOne" style="position:absolute; visibility:hidden;z-index:700;">
<table width="auto" style="margin:0;text-align: left; padding:0; background-color:#EBF2F9; border:0; color:#FFFFFF;">
<tr>
<td>姓名或工号</td><td><input type="text" id="SearchID" name="SearchID" /></td>
</tr>
<tr>
<td colspan="2">
<input type="button" value="确定" style="width:60px; height:16px; font-size:12px " onclick="submitSearch()"/>
<input type="button" value="关闭" style="width:60px; height:16px; font-size:12px " onclick="$('#editOne').hide()"/></td>
</tr>
</table>
</div>
</body>
</html>
