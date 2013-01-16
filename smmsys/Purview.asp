<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!--#include file="../CheckAdmin.asp" -->
<%
if Instr(session("AdminPurview"),"|1104,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="报表系统" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
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
		expandSpeed:"",
		dragCopy:true,
		dragCopy:true,
		keepParent: false,
		keepLeaf: false,
		async: true,
		asyncParam: ["SerialNum"],
		asyncParamOther : {"showType":"getInfo","detailType":"EditPurview"},
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
	function renameTreeNode(evt) {
		var srcNode = zTree1.getSelectedNode();
		if (!srcNode) {
			alert("请先选中一个节点");
			return;
		}
		hideRMenu();
		$('#PermissionID').val(srcNode.PermissionID);
		$('#PermissionName').val(srcNode.name);
		$("#editOne").show();
		$("#editOne").css({"top":evt.clientY+"px", "left":evt.clientX+"px", "visibility":"visible"});
	}
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
	var addCount = 1;
	function addTreeNode() {
		hideRMenu();
		var srcNode=zTree1.getSelectedNode();
		var p_s=0;
		var p_pname=($('#PurviewType').val()=='RPT')?'报表系统':'工作流系统';
		var p_pid="";
		var p_lv=0;
		if(srcNode){
		  p_s=srcNode.SerialNum;
		  p_pname=srcNode.LongName+"-"+srcNode.name;
		  p_pid=srcNode.PermissionID;
		  p_lv=srcNode.level;
		}
		$.post('PurviewDetails.asp?showType=DataProcess&detailType=AddNew',{
		  PermissionClass:$('#PurviewType').val(),
		  PermissionID:p_pid,
		  LongName:p_pname,
		  PSNum:p_s,
		  N_lv:p_lv
		},function(data){
		  var newNode=jQuery.parseJSON(data);
		  zTree1.addNodes(srcNode, newNode);
		});
	}
	function removeTreeNode() {
		hideRMenu();
		var node = zTree1.getSelectedNode();
		if (node) {
			if (node.nodes && node.nodes.length > 0) {
				alert("要删除的节点是父节点，不允许执行删除操作，请先删除子节点！");
				return ;
			} else {
				$.post('PurviewDetails.asp?showType=DataProcess&detailType=Delete',{
				  SerialNum:node.SerialNum
				},function(data){
				  if(data.indexOf("###")==-1) alert(data);
				  else {
					zTree1.removeNode(node);
				  }
				});
			}
		}
	}
	
	function refreshTree() {
		var paraUrl=$('#PurviewType').length==0?'RPT':$('#PurviewType').val();
		setting.asyncUrl= "PurviewDetails.asp?PurviewType="+paraUrl;
		hideRMenu();
		$('#editOne').hide();
		var nodes = [];
		setting.async = true;
		zTree1 = $("#Permission").zTree(setting, nodes);
	}
	function submitSave(){
		var srcNode = zTree1.getSelectedNode();
		if (!srcNode) {
			alert("请先选中一个节点");
			return;
		}
		$.post('PurviewDetails.asp?showType=DataProcess&detailType=Edit',{
		  PermissionID:$('#PermissionID').val(),
		  PermissionName:$('#PermissionName').val(),
		  SerialNum:srcNode.SerialNum
		},function(data){
		  if(data.indexOf("###")==-1) alert(data);
		  else {
			$('#editOne').hide();
			srcNode.PermissionID=$('#PermissionID').val();
			srcNode.name=$('#PermissionName').val();
			zTree1.updateNode(srcNode, true);
		  }
		});
	}

</script>
</HEAD>

<BODY>
<div id="rMenu" style="position:absolute; visibility:hidden;">
<li>
<ul id="m_add" onclick="addTreeNode();"><li>增加节点</li></ul>
<ul id="m_edit" onclick="renameTreeNode(event);"><li>编辑节点</li></ul>
<ul id="m_del" onclick="removeTreeNode();"><li>删除节点</li></ul>
<ul id="m_reset" onclick="refreshTree();"><li>权限重新载入</li></ul>
</li>
</div>
<div id="editOne" style="position:absolute; visibility:hidden;">
<table width="auto" style="margin:0;text-align: left; padding:0; background-color:#EBF2F9; border:0; color:#FFFFFF;">
<tr>
<td>权限代码</td><td><input type="text" id="PermissionID" name="PermissionID" /></td>
</tr>
<tr>
<td>权限名称</td><td><input type="text" id="PermissionName" name="PermissionName" /></td>
</tr>
<tr>
<td colspan="2">
<input type="button" value="确定" style="width:60px; height:16px; font-size:12px " onclick="submitSave()"/>
<input type="button" value="取消" style="width:60px; height:16px; font-size:12px " onclick="$('#editOne').hide()"/></td>
</tr>
</table>
</div>
<table width="100%" border="1">
  <tr>
    <td width="10%">
	<select id="PurviewType" onchange="refreshTree()">
	<option value="RPT">报表系统</option>
	<option value="FLW">工作流平台</option>
	</select>
	</td>
	<td width="80%">
			<ul id="Permission" class="tree"></ul>
	</td>
  </tr>
</table>
</BODY>
</HTML>