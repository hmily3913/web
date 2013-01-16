<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<HTML>
<HEAD>
<TITLE>后台管理导航</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012 - zbh-STUDIO" />
<META NAME="Author" CONTENT="---zbh" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<script language="javascript" src="Script/Admin.js"></script>
<SCRIPT language=JavaScript>
//切换平台
function ChangePlat(obj){
	if(obj==1){
		parent.frames["leftFrame"].location.href="SysLeft.asp";
		if(parent.frames["mainFrame"].location.href.indexOf('SysCome.asp')==-1)
			parent.frames["mainFrame"].location.href="SysCome.asp";
		changeAdminFlag("报表系统");
	}else if (obj==2){
		parent.frames["leftFrame"].location.href="FLW_SysLeft.asp";
		if(parent.frames["mainFrame"].location.href.indexOf('SysCome.asp')==-1)
			parent.frames["mainFrame"].location.href="SysCome.asp";
		changeAdminFlag("工作流平台");
	}
}

</SCRIPT>
</head>
<!--#include file="CheckAdmin.asp"-->

<body topmargin="0" bottom="0" leftmargin="0" rightmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td align="right" style="width:200px;height:30px;background: url('Images/Head-bg.gif') repeat-x scroll center center #7F99BE;">
	<td align="left" style="height:30px;background: url('Images/Head-bg.gif') repeat-x scroll center center #7F99BE;">
	&nbsp;<font color="#0000FF">|</font>
	<a href="javascript:ChangePlat(1)"><font color="#FFFFFF" style="font-size:12;font-weight:bold;">报表系统</font></a>&nbsp;<font color="#0000FF">|</font>
	<a href="javascript:ChangePlat(2)"><font color="#FFFFFF" style="font-size:12;font-weight:bold;">工作流系统</font></a>&nbsp;<font color="#0000FF">|</font>&nbsp;&nbsp;</td>
	<td align="right" style="background:url('Images/Head-bg.gif') repeat-x scroll center center #7F99BE;"><font color="#0000FF">|</font>&nbsp;
	<a href="javascript:parent.initmessList()" id="duser"><font color="#FFFFFF" style="font-size:12;font-weight:bold;" id="messageMana">短信管理</font></a>&nbsp;
	<font color="#0000FF">|</font>&nbsp;
	<a href="javascript:parent.refreshTree()" id="duser"><font color="#FFFFFF" style="font-size:12;font-weight:bold;">同事管理</font></a>&nbsp;<font color="#0000FF">|</font>&nbsp;&nbsp;</td>
  </tr>
</table>
</body>