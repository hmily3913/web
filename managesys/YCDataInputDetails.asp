<!--#include file="../CheckAdmin.asp" -->
<%
if Instr(session("AdminPurview"),"|1009,")=0 then 
  response.write ("你不具有该管理模块的操作权限，请返回！")
  response.end
end if
'========判断是否具有管理权限
Dim Conn,ConnStr,Conn2,ConnStr2
On error resume next
Set Conn=Server.CreateObject("Adodb.Connection")
ConnStr="driver={SQL Server};server=192.168.0.5;UID=sa;PWD=loveradmin;Database=YC2011"
Conn.open ConnStr
Set Conn2=Server.CreateObject("Adodb.Connection")
ConnStr2="driver={SQL Server};server=192.168.0.5;UID=sa;PWD=loveradmin;Database=KQ2011"
Conn2.open ConnStr2
if err then
   err.clear
   Set Conn = Nothing
   Set Conn2 = Nothing
   Response.Write "系统错误：数据库连接出错，请检查'系统管理>>站点常量设置',请联系管理员!"
   Response.End
end if

dim showType,start_date,end_date,print_tag,UserName,AdminName
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")
print_tag=request("print_tag")
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
if showType="xls2sql" then
	Server.ScriptTimeout = 999999
	conn.execute("delete from YCtemp")
	set rs=server.createobject("adodb.recordset")
	sql="select * from YCtemp"
	rs.open sql,conn,3,3
	InfoID=request("InfoID")
	Set xlApp=Server.CreateObject("Excel.Application")          '/******** VBA方法 连接Excel *********/
	Set xlbook=xlApp.Workbooks.Open(Server.mappath(InfoID))  
	Set xlsheet=xlbook.Worksheets(2)  
	i=3
	While cstr(xlsheet.cells(i,2))<>""           '/********** 使用第3列 帐号为空时判断为结束标志  **********/
	
	if xlsheet.cells(i,8)<>"" or xlsheet.cells(i,9)<>"" then
		rs.Addnew()  
		for n=1 to 10
			rs("z"&n)=xlsheet.cells(i,n)
		next
		rs.Update  
	end if
	i=i+1  
	Wend  
	xlsheet.close
	Set xlsheet=nothing  
	xlbook.Close  
	Set xlbook=Nothing  
	xlApp.DisplayAlerts=false
	xlApp.Quit  
	
	sql="insert into CHECKINOUT  "&_
" select a.Userid,z6+' '+z8 as checktime,'I',1,1,null,0,null,0 "&_
" from USERINFO a,YCtemp as b "&_
" where a.ssn=z2 and z8 is not null and z9<>'' "&_
" union all  "&_
" select a.Userid,z6+' '+z9 as checktime,'I',1,1,null,0,null,0 "&_
" from USERINFO a,YCtemp as b "&_
" where a.ssn=b.z2 and z9 is not null and z9<>'' "&_
" order by checktime ,userid"
	conn.execute(sql)
	conn.execute("delete from YCtemp")
	rs.Close
	Set rs=Nothing
	response.Write("共计"&i-1&"条数据导入成功!")
elseif showType="cleardata" then
	conn.execute("delete from YCtemp")
	conn.execute("delete from CHECKINOUT")
	response.Write("数据删除完毕！")
elseif showType="ClearUsers" then
	Conn2.execute("delete from USERINFO where ssn is null")
	response.Write("数据删除完毕！")
end if
 %>
