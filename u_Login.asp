<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012 - zbh" />
<META NAME="Author" CONTENT="企业管理系统" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>用户登录</TITLE>
<LINK href="images/User_Login.css" type=text/css rel=stylesheet>
<script language="javascript" >
//检查用户登录------------------------------------------------------------------------------
function CheckAdminLogin()
{
   var check; 
   if (!voidNum(document.AdminLogin.LoginName.value))
   {
	  alert("请正确输入用户名称（由0-9,a-z,-_任意组合的字符串）。");
      document.AdminLogin.LoginName.focus();
	  return false;
	  exit;
   }    
   if (!voidNum(document.AdminLogin.LoginPassword.value))
   {
	  alert("请输入用户密码。");
	  document.AdminLogin.LoginPassword.focus();
	  return false;
	  exit;
   }
/*   if (!voidNum(document.AdminLogin.VerifyCode.value))
   {
      alert("请正确输入验证码。");
      document.AdminLogin.VerifyCode.focus();
	  return false;
	  exit;
   }*/
   return true;
}
</script>
<script language="javascript" src="Script/Admin.js"></script>
</HEAD>
<BODY id=userlogin_body>
<DIV></DIV>
	<form action="CheckLogin.asp" method="post" name="AdminLogin" id="AdminLogin"  onSubmit="return CheckAdminLogin()">
<DIV id=user_login>
<DL>
  <DD id=user_top>
  <UL>
    <LI class=user_top_l></LI>
    <LI class=user_top_c></LI>
    <LI class=user_top_r></LI></UL>
  <DD id=user_main>
  <UL>
    <LI class=user_main_l></LI>
    <LI class=user_main_c>
    <DIV class=user_main_box>
    <UL>
      <LI class=user_main_text>用户名： </LI>
      <LI class=user_main_input><INPUT class=TxtUserNameCssClass id=LoginName 
      maxLength=20 name=LoginName> </LI></UL>
    <UL>
      <LI class=user_main_text>密 码： </LI>
      <LI class=user_main_input><INPUT class=TxtPasswordCssClass id=LoginPassword 
      type=password name=LoginPassword> </LI></UL>
    </DIV></LI>
    <LI class=user_main_r><INPUT class=IbtnEnterCssClass id=submitLogin 
    style="BORDER-TOP-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; BORDER-RIGHT-WIDTH: 0px" 
    type=image src="images/user_botton.gif" name="submitLogin"> </LI></UL>
  <DD id=user_bottom>
  <UL>
    <LI class=user_bottom_l></LI>
    <LI class=user_bottom_c><SPAN style="MARGIN-TOP: 40px"></SPAN> </LI>
    <LI class=user_bottom_r></LI></UL></DD></DL></DIV><SPAN id=ValrUserName 
style="DISPLAY: none; COLOR: red"></SPAN><SPAN id=ValrPassword 
style="DISPLAY: none; COLOR: red"></SPAN><SPAN id=ValrValidateCode 
style="DISPLAY: none; COLOR: red"></SPAN>
<DIV id=ValidationSummary1 style="DISPLAY: none; COLOR: red"></DIV>

<DIV></DIV>

</FORM></BODY>


</HTML>
