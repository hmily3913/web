<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="zbh" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>产品列表</TITLE>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<script language="JavaScript">
<!--
function window_onload(){
 printjq();
}

function printjq()
{
  try {
      var ExcelID    = new ActiveXObject ( "Excel.Application" );
     }
  catch(e) {
         alert( e+"要打印该表，您必须安装Excel电子表格软件，同时浏览器须使用“ActiveX 控件”，您的浏览器须允许执行控件。 请点击【帮助】了解浏览器设置方法！");
         return "";
  }
  ExcelID.visible = true;
  var newBook=ExcelID.Workbooks.Add;
  var kwb= newBook.Worksheets.Add; 
  var ksheet = newBook.Worksheets(1);
  ksheet.ActiveSheet;
<%
dim sql,rs
dim i
i=1
sql="select * from ( select FBillNo as id,FBillNo_SRC as 销售订单号,FAlterReason as 变更原因, FAlterDate as 变更日期,ICItem.fnumber,ICItem.fname,falterqty,falteramount from t_DDBGTZDEntry inner join t_DDBGTZD on t_DDBGTZDEntry.finterid=t_DDBGTZD.finterid inner join dbo.t_ICItem AS ICItem ON ICItem.FItemID = t_DDBGTZDEntry.FItemID where FCheckerID>0 and FAlterDate>='2011-4-3' and FAlterDate<='2011-4-3' ) as aaa"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  while(not rs.eof)
%>
   ksheet.Cells(<%=i%>,1).value="<%=rs("变更原因")%>";
<%  
i=i+1
rs.movenext
wend
  rs.close
  set rs=nothing
%>
   ExcelID.Visible = true; 
   ExcelID.UserControl = true; 
//   ExcelID.DisplayAlerts = false;
  window.close();
}

-->
</script>
</head>

<body onLoad="window_onload()">
</body>
</html>
