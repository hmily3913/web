<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|1102,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<% 
dim Result,rs,sql
Result=request.QueryString("Result")
dim ID,UserName,Purview
ID=request.QueryString("ID")
set rs = server.createobject("adodb.recordset")
sql="select * from [N-基本资料单头] where 员工代号='"&ID&"'"
rs.open sql,conn,1,1
UserName=rs("姓名")
Purview=rs("权限")
rs.close
set rs=nothing 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="zbh" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>编辑管理员</TITLE>
<link rel="stylesheet" href="../Images/CssAdmin.css">
<link rel="stylesheet" href="../Images/zTreeStyle.css">
<style type="text/css">
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
<script language="javascript">
var zTree1;
var setting;
var rMenu;
	setting = {
		checkable: true,
		checkType : {"Y":"ps", "N":"ps"},
		expandSpeed:"",
		keepParent: false,
		keepLeaf: false,
		isSimpleData : true,
		treeNodeParentKey : "PSNum",
		treeNodeKey : "SerialNum",
		rootPID : 0,
		async: true,
		asyncUrl: "PurviewDetails.asp",
		asyncParamOther : {"showType":"getInfo","detailType":"SetPurview","PurviewType":"RPT","EmpId":"<%=ID%>"},
		callback: {
			rightClick: zTreeOnRightClick
		}
	};
	$(document).ready(function(){
		refreshTree();
		rMenu = document.getElementById("rMenu");
		$("body").bind("mousedown", 
			function(event){
				if (!(event.target.id == "rMenu" || $(event.target).parents("#rMenu").length>0)) {
					rMenu.style.visibility = "hidden";
				}
			});
	});
	//右键菜单
	function zTreeOnRightClick(event, treeId, treeNode) {
		if (!treeNode && event.target.tagName.toLowerCase() != "button" && $(event.target).parents("a").length == 0) {
			zTree1.cancelSelectedNode();
			showRMenu("root", event.clientX, event.clientY);
		} else if (treeNode && !treeNode.noR) {
			zTree1.selectNode(treeNode);
			showRMenu("node", event.clientX, event.clientY);
		}
	}
	function showRMenu(type, x, y) {
		$("#rMenu ul").show();
		if (type=="root") {
			$("#m_check").hide();
			$("#m_unCheck").hide();
		}
		$("#rMenu").css({"top":y+"px", "left":x+"px", "visibility":"visible"});
	}
	function hideRMenu() {
		if (rMenu) rMenu.style.visibility = "hidden";
	}
	
	function checkTreeNode(checked) {
		var node = zTree1.getSelectedNode();
		if (node) {
			node.checked = checked;
			zTree1.updateNode(node, true);
		}
		hideRMenu();
	}

	function refreshTree() {
		hideRMenu();
		var nodes = [];
		setting.async = true;
		zTree1 = $("#PermissionRPT").zTree(setting, nodes);
	}

	function submitSave(){
		var tmp = zTree1.getCheckedNodes();
		var pall="";
		for (var i=0; i<tmp.length; i++) {
			pall+=tmp[i].PermissionID;
		}
		$.post('PurviewDetails.asp?showType=DataProcess&detailType=SetEdit',{
		  PurviewType:"RPT",
		  PermissionID:pall,
		  EmpId:"<%=ID%>"
		},function(data){
		  alert(data);
		});
	}

</script>
</HEAD>

<BODY>
<div id="rMenu" style="position:absolute; visibility:hidden;">
<li>
<ul id="m_check" onclick="checkTreeNode(true);"><li>Check节点</li></ul>
<ul id="m_unCheck" onclick="checkTreeNode(false);"><li>unCheck节点</li></ul>
<ul id="m_reset" onclick="refreshTree();"><li>权限重新载入</li></ul>
</li>
</div>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>网站管理员：添加，修改管理员信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a >添加管理员</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="PurviewSet.asp" onClick='changeAdminFlag("网站管理员")'>查看所有员工</a></td>    
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <form name="editForm" method="post" action="PurviewEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">登&nbsp;录&nbsp;名：</td>
        <td><input name="UserName" type="text" class="textfield" id="UserName" style="WIDTH: 120;" value="<%=UserName%>" maxlength="16" readonly>&nbsp;*&nbsp;3-10位字符，不可修改</td>
      </tr>
      <tr >
        <td height="20" align="right">操作权限：</td>
        <td nowrap>
  		<div class="zTreeDemoBackground">
			<ul id="PermissionRPT" class="tree"></ul>
		</div>		
	  </td>

      </tr>

   <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 60;" onclick="return submitSave()"></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</BODY>
</HTML>
