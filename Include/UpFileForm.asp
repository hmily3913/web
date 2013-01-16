<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/CssAdmin.css">
<title>文件选择</title>
</head> 
<!--#include file="CheckAdmin.asp"-->
<body>
<table width="400" border="0" align="center" cellpadding="12" cellspacing="1" bgcolor="#99BBE8">
  <form action="UpFileSave.asp" method="post" enctype="multipart/form-data" name="formUpload" id="formUpload">
  <tr>
    <td bgcolor="#EBF2F9">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="60" height="30" nowrap>选择文件：</td>
        <td><input name="FromFile" type="file" class="textfield" id="FromFile" size="41" class="multi"></td>
      </tr>
      <tr>
        <td height="36" colspan="2" align="center" valign="bottom"><input name="reset" type="reset" class="button" value=" 重置 ">
          &nbsp;<input name="Submit" type="submit" class="button" value=" 上传 "></td>
        </tr>
    </table>
	</td>
  </tr>
  </form>
</table>
</body>
</html>

