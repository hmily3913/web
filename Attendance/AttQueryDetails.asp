<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|203,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
	 if request("cq")="bh" then
		Set Connkq=Server.CreateObject("Adodb.Connection")
		ConnStrkq="driver={SQL Server};server=192.168.0.184;UID=sa;PWD=ldrz;Database=KQ2011"
		Connkq.open ConnStrkq
	 elseif request("cq")="lq" then
		Set Connkq=Server.CreateObject("Adodb.Connection")
		ConnStrkq="driver={SQL Server};server=122.228.158.226;UID=sa;PWD=lovemaster;Database=att2000"
		Connkq.open ConnStrkq
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
if showType="DetailsList" then 
%>
 <div id="listtable" style="width:100%; height:100%; overflow:auto">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr class="TitleRow">
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>姓名</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>时段</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>上班时间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>下班时间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>签到时间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>签退时间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>例外情况</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>说明</strong></font></td>
  </tr>
 <%
  dim Depart,wherestr
  wherestr=""
  if request("ID")="" and request("DepartID")="" then
		response.Write("部门和工号不能同时为空！")
		response.End()
	end if
  if request("Sdate")="" or request("Edate")="" then
		response.Write("日期不能为空！")
		response.End()
	end if

'  sql="select date,a.userid,a.name as xm,h.deptname,a.ssn,c.num_runid,c.name,c.units,d.sdays,d.edays,e.schclassid,e.schName,  "
'	sql=sql&"e.starttime,e.endtime,e.checkin,e.checkout,e.checkintime1,e.checkintime2,e.checkouttime1,e.checkouttime2,e.workday  "
'	sql=sql&"from USERINFO a,USER_OF_RUN b,NUM_RUN c,NUM_RUN_DEIL d,SchClass e ,Calendar,DEPARTMENTS h "
'	sql=sql&"where a.userid=b.userid and b.num_of_run_id=c.num_runid and a.defaultdeptid=h.deptid  "
'	sql=sql&"and c.num_runid=d.num_runid and d.schclassid=e.schclassid and b.startdate<=date and b.enddate>=date "
  sql="select date,a.userid,a.name as xm,h.deptname,a.ssn,c.num_runid,c.name,c.units,d.sdays, "
	sql=sql&"d.edays,e.schclassid,e.schName, e.starttime,e.endtime,e.checkin,e.checkout,e.checkintime1, "
	sql=sql&"e.checkintime2,e.checkouttime1,e.checkouttime2,e.workday ,min(i.checktime) as ttime1,max(j.checktime) as ttime2,k.leavename,k.YUANYING "
	sql=sql&"from USERINFO a inner join USER_OF_RUN b on a.userid=b.userid "
	if request("ID")<>"" then
		sql=sql&" and a.ssn='"&request("ID")&"' "
	else
		sql2="exec sp_getChildDEPT "&request("DepartID")
		set rs2=server.createobject("adodb.recordset")
		rs2.open sql2,ConnStrkq,0,1
		sql=sql&" and (1=2 "
		while(not rs2.eof)
			sql=sql&" or a.ssn='"&rs2("ssn")&"' "
			rs2.movenext
		wend
		sql=sql&" ) "
	end if
	sql=sql&"inner join NUM_RUN c on b.num_of_run_id=c.num_runid inner join  "
	sql=sql&"NUM_RUN_DEIL d on c.num_runid=d.num_runid inner join "
	sql=sql&"SchClass e on d.schclassid=e.schclassid inner join  "
	sql=sql&"Calendar on b.startdate<=date and b.enddate>=date and  "
	sql=sql&"date>='"&request("Sdate")&"' and date<='"&request("Edate")&"'  inner join "
	sql=sql&"DEPARTMENTS h on a.defaultdeptid=h.deptid  and ((DATEPART(weekday,date)-1=d.sdays%7 and c.units=1)  "
	sql=sql&"or ((day(date)+(datediff(m,c.startdate,date)%c.cyle)*31)=d.sdays and c.units=2)) left join "
	sql=sql&"CHECKINOUT i on datediff(d,i.checktime,date)=0 and datediff(s,i.checktime,date+' '+convert(varchar(8),e.checkintime1,114))<0  "
	sql=sql&"and datediff(s,i.checktime,date+' '+convert(varchar(8),e.checkintime2,114))>0 and i.userid=a.userid left join "
	sql=sql&"CHECKINOUT j on datediff(d,j.checktime,date)=0 and datediff(s,j.checktime,date+' '+convert(varchar(8),e.checkouttime1,114))<0  "
	sql=sql&"and datediff(s,j.checktime,date+' '+convert(varchar(8),e.checkouttime2,114))>0 and j.userid=a.userid left join "
	sql=sql&"(select b.leavename,a.YUANYING,startspecday,endspecday,userid from USER_SPEDAY a,LeaveClass b  "
	sql=sql&"where a.dateid=b.leaveid and datediff(d,startspecday,'"&request("Sdate")&"')<=0 and datediff(d,startspecday,'"&request("Edate")&"')>=0) k on "
	sql=sql&"datediff(d,k.startspecday,date)=0  "
	sql=sql&"and (datediff(s,k.startspecday,date+' '+convert(varchar(8),e.endtime,114))>=0  "
	sql=sql&"and datediff(s,endspecday,date+' '+convert(varchar(8),e.starttime,114))<=0)  "
	sql=sql&"and k.userid=a.userid "
	sql=sql&"group by date,a.userid,a.name,h.deptname,a.ssn,c.num_runid,c.name,c.units,d.sdays, "
	sql=sql&"d.edays,e.schclassid,e.schName, e.starttime,e.endtime,e.checkin,e.checkout,e.checkintime1, "
	sql=sql&"e.checkintime2,e.checkouttime1,e.checkouttime2,e.workday,a.defaultdeptid,k.leavename,k.YUANYING "
	sql=sql&"order by a.defaultdeptid,a.userid,date,e.starttime "

'	sql=sql&"and ((DATEPART(weekday,date)-1=d.sdays%7 and c.units=1)  "
'	sql=sql&"or ((day(date)+(datediff(m,c.startdate,date)%c.cyle)*31)=d.sdays and c.units=2))  "
'	sql=sql&"and date>='"&request("Sdate")&"' and date<='"&request("Edate")&"' "
'	sql=sql&"order by a.defaultdeptid,a.userid,date,e.starttime "

    set rs=server.createobject("adodb.recordset")
    rs.open sql,ConnStrkq,1,1
    while(not rs.eof)'填充数据到表格
		dim yclb:yclb=""
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("deptname")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ssn")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("xm")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("date")&"</td>"
      Response.Write "<td nowrap>"&rs("schName")&"</td>"
      Response.Write "<td nowrap>"&right(rs("starttime"),8)&"</td>"
      Response.Write "<td nowrap>"&right(rs("endtime"),8)&"</td>"
'			sql2="select min(checktime) as ttime from CHECKINOUT where datediff(d,checktime,'"&rs("date")&"')=0 and datediff(s,checktime,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("checkintime1")&"',114))<0 and datediff(s,checktime,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("checkintime2")&"',114))>0 and userid="&rs("userid")
'				set rs2=server.createobject("adodb.recordset")
'				rs2.open sql2,ConnStrkq,0,1
				if isnull(rs("ttime1")) then
					if rs("checkin")=1 then yclb="旷工"
		response.Write("<td></td>")
				else
					if datediff("s",rs("ttime1"),rs("date")&" "&right(rs("starttime"),8))<0 then
		response.Write("<td>"&right(rs("ttime1"),8)&"</td>")
		if rs("checkin")=1 then yclb="迟到"
					else
		response.Write("<td>"&right(rs("ttime1"),8)&"</td>")
					end if
				end if
'				sql2="select max(checktime) as ttime from CHECKINOUT where datediff(d,checktime,'"&rs("date")&"')=0 and datediff(s,checktime,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("checkouttime1")&"',114))<0 and datediff(s,checktime,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("checkouttime2")&"',114))>0 and userid="&rs("userid")
'				set rs2=server.createobject("adodb.recordset")
'				rs2.open sql2,ConnStrkq,0,1
				if isnull(rs("ttime2")) then
					if rs("checkout")=1 then yclb="旷工"
		response.Write("<td></td>")
				else
					if datediff("s",rs("ttime2"),rs("date")&" "&left(rs("endtime"),12))>0 then
		response.Write("<td>"&right(rs("ttime2"),8)&"</td>")
			if rs("checkout")=1 then yclb="早退"
					else
		response.Write("<td>"&right(rs("ttime2"),8)&"</td>")
					end if
				end if
'			sql2="select b.leavename,a.YUANYING from USER_SPEDAY a,LeaveClass b where a.dateid=b.leaveid and datediff(d,startspecday,'"&rs("date")&"')=0 and (datediff(s,startspecday,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("endtime")&"',114))>=0 and datediff(s,endspecday,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("starttime")&"',114))<=0) and userid="&rs("userid")
'			set rs2=server.createobject("adodb.recordset")
'			rs2.open sql2,ConnStrkq,0,1
			if isnull(rs("leavename")) then
				response.Write("<td>"&yclb&"</td>")
				response.Write("<td></td>")
			else
				response.Write("<td>"&rs("leavename")&"</td>")
				response.Write("<td>"&rs("YUANYING")&"</td>")
			end if
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
'	  rs2.close
'	  set rs2=nothing
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
elseif showType="getInfo" then 
	if request("DepartID")<>"" then
		sql="exec sp_getChildDEPT "&request("DepartID")
	else
		sql="select ssn,name from USERINFO "
	end if
		set rs=server.createobject("adodb.recordset")
		rs.open sql,ConnStrkq,1,1
		response.write "["
		do until rs.eof
		Response.Write("{")
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
			next
			if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
			else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}")
			end if
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
elseif showType="DepartMent" then
    set rs = server.createobject("adodb.recordset")
    sql="select * from DEPARTMENTS order by DeptID"
    rs.open sql,Connkq,1,1
		response.Write("[")
		do until rs.eof
	%>
	{"SerialNum": "<%=rs("DeptID")%>", "name": "<%=rs("DEPTName")%>", "PSNum": "<%=rs("SUPDeptID")%>"
	<%
		response.Write("}")
			rs.movenext
			If Not rs.eof Then
			  Response.Write ","
			End If
		loop
		Response.Write "]"
		rs.close
		set rs=nothing
end if
 %>

