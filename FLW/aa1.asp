<%
Dim Connk3,ConnStrk3

Set Connk3=Server.CreateObject("Adodb.Connection")
ConnStrk3="driver={SQL Server};server=192.168.0.9;UID=sa;PWD=lovemaster;Database=AIS20081217153921"
Connk3.open ConnStrk3

dim varfileid,varfilename,varfilesize,varcontent
response.Expires=0
response.Buffer=true
response.Clear()


dim SQL,rs
dim Fentry,fields
Fentry=request("id")
fields=request("fields")
set rs=server.createobject("adodb.recordset")
SQL = "select "&fields&" from t_dhtzd where fid="&Fentry'fentryid=34121
rs.open SQL,connk3,1,1
If Not rs.Eof Then
varfilename = "sss.rtf"'rs("fbigtext2")
varfilesize=rs(fields).ActualSize
varcontent = rs(fields).GetChunk(varfilesize)
Response.ContentType = "*/*"
'Response.AddHeader "Content-Length",varfilesize
'Response.AddHeader "Content-Disposition", "attachment;filename=""" & varfilename & """"
Response.binarywrite varcontent
'dim strChartAbsPath
'strChartAbsPath=Server.MapPath("../temp/")
'Set   iStm   =   Server.CreateObject("ADODB.Stream")
'        With   iStm 
'                .Mode   =   3 
'                .Type   =   1 
'				.charset = "gb2312"
'                .Open 
'                .Write   rs("fbigtext5_tag")
'				.SaveToFile   strChartAbsPath   &   "\test1.txt"
'				End   With

End If
rs.Close
Set rs = Nothing



%>
