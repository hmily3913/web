<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurview"),"|1002,")=0 then 
  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
  response.end
end if

'========判断是否具有管理权限
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
  dim page'页码
      page=clng(request("page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=request("rp")
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" t_dormperson a left join t_dormpersonentry b on a.FId=b.FId left join t_emp c on b.person=c.fitemid "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where 1=1 "
	Dim searchterm,searchcols
	
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = Request.Form("qtype")
	if isnumeric(searchterm) then
	datawhere = datawhere&" and " & searchcols & " = " & searchterm & " "
	else
	datawhere = datawhere&" and " & searchcols & " LIKE '%" & searchterm & "%' "
	end if
	End if
		 if request("lh")<>"" then datawhere=datawhere&" and a.Ftext1='"&request("lh")&"' "
		 if request("yg")<>"" then datawhere=datawhere&" and (c.FNumber like '%"&request("yg")&"%' or c.FName like '%"&request("yg")&"%') "
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "a.FID" 
	Else
	sortname = Request.Form("sortname")
	End If
	Dim sortorder
	if Request.Form("sortorder") = "" then
	sortorder = "desc"
	Else
	sortorder = Request.Form("sortorder")
	End If
      taxis=" order by "&sortname&" "&sortorder
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(distinct a.FID) as idCount from "& datafrom &" " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  idCount=rs("idCount")
  if(idcount>0) then'如果记录总数=0,则不处理
    if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
	  pagec=int(idcount/pages)'获取总页数
   	else
      pagec=int(idcount/pages)+1'获取总页数
    end if
  end if
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="select distinct a.FID from "& datafrom &" " & datawhere 
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("FID")
	  else
	    sqlid=sqlid &","&rs("FID")
	  end if
	  rs.movenext
    next
  '获取本页需要用到的id结束============================================
'-----------------------------------------------------------
'-----------------------------------------------------------
%>
{"page":"<%=page%>","total":"<%=idcount%>","rows":[
<%
  if sqlid<>"" then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    sql="select  a.fbillno,a.Ftext,a.Ftext1,a.louceng,a.maxperson,a.sumperson,a.FDecimal,a.waternum,a.electnum,a.HotWater,a.FID from t_dormperson a where a.FID in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    do until rs.eof'填充数据到表格'
%>		
		{"id":"<%=rs("FID")%>",
		"cell":[
<%		
	  for i=0 to rs.fields.count-2
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).value&""",")
		else
		response.write (""""&JsonStr(rs.fields(i).value)&""",")
		end if
	  next
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).value&"""]}")
		else
		response.write (""""&JsonStr(rs.fields(i).value)&"""]}")
		end if

	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
  end if
  rs.close
  set rs=nothing
response.Write"]}"
'-----------------------------------------------------------'
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
  	if  Instr(session("AdminPurview"),"|1002.1,")>0 then
		dim rs2,FIDone
		set rs = server.createobject("adodb.recordset")
		sql="select * from t_dormperson"
		rs.open sql,connk3,1,3
		rs.addnew
'		rs("FClassTypeId")="257709051"
		set rs2=connk3.Execute("select max(Fid) as FID,max(FbillNo) as FbillNo from t_dormperson")
		if rs2("FID")<>"" then
		rs("FID")=Cint(rs2("FID"))+1
		FIDone=Cint(rs2("FID"))+1
		rs("FbillNo")=rs2("FbillNo")+1
		else
		rs("FID")="1000"
		FIDone=1000
		rs("FbillNo")="10000000"
		end if
		rs2.close
		set rs2=nothing
		  set rsRepeat = connk3.execute("select ftext from t_dormperson where ftext='" & trim(Request.Form("ftext")) & "'")
		  if not (rsRepeat.bof and rsRepeat.eof) then '宿舍号
			response.write "此宿舍号已经被使用，请换一个编号再试试！"
			response.end
		  else
			rs("ftext")=trim(Request.Form("ftext"))
		  end if
		  rs("ftext1")=trim(Request.Form("ftext1"))
		  rs("louceng")=Request.Form("louceng")
		  rs("maxperson")=Request.Form("maxperson")
		  rs("electnum")=Request.Form("electnum")
		  rs("FDecimal")=Request.Form("FDecimal")
		  rs("waternum")=Request.Form("waternum")
		  rs("HotWater")=Request.Form("HotWater")
			rs("useflag")=Request.Form("useflag")
			rs("showflag")=Request.Form("showflag")
		  rs("FBiller")=trim(Request.Form("FBiller"))
		  rs("FDate")=trim(Request.Form("FDate"))
		rs.update
		rs.close
		set rs=nothing 
		for   i=1   to   Request.form("FEntryID").count
			if Request.form("person")(i)<>"" then
				'添加子表
				connk3.Execute("insert into t_dormpersonEntry values ("&FIDone&","&i&",'"&Request.form("person")(i)&"',"&Request.form("finteger3")(i)&",'"&Request.form("fbase1")(i)&"','"&Request.form("fdate1")(i)&"')")
				'更新宿舍信息表
			end if
		next
		connk3.Execute("update t_dormperson set sumperson=a.a1 from (select count(1) as a1,fid as  a2 from t_dormpersonEntry group by fid) a where a.a2=fid and fid="&FID)
		response.write "###"
	else
		response.write ("你没有权限进行此操作！")
		response.end
	end if
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
    FID=request("FID")
		sql="select * from t_dormperson where FID="&Request.Form("FID")
		rs.open sql,connk3,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurview"),"|1002.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
    end if
		  set rsRepeat = connk3.execute("select ftext from t_dormperson where ftext='" & trim(Request.Form("ftext")) & "' and FID<>"&Request.Form("FID"))
		  if not (rsRepeat.bof and rsRepeat.eof) then '宿舍号
			response.write "此宿舍号已经被使用，请换一个编号再试试！"
			response.end
		  else
			rs("ftext")=trim(Request.Form("ftext"))
		  end if
		  rs("ftext1")=trim(Request.Form("ftext1"))
		  rs("louceng")=Request.Form("louceng")
		  rs("maxperson")=Request.Form("maxperson")
		  rs("electnum")=Request.Form("electnum")
		  rs("FDecimal")=Request.Form("FDecimal")
		  rs("waternum")=Request.Form("waternum")
		  rs("HotWater")=Request.Form("HotWater")
			rs("useflag")=Request.Form("useflag")
			rs("showflag")=Request.Form("showflag")
		  rs("FBiller")=trim(Request.Form("FBiller"))
		  rs("FDate")=trim(Request.Form("FDate"))
		rs.update
		rs.close
		set rs=nothing 
		for   i=1   to   Request.form("FEntryID").count
			if Request.form("DeleteFlag")(i)="1" or (Request.form("person")(i)="" and Request.Form("FEntryID")(i)<>"") then
				connk3.Execute("Delete from t_dormpersonEntry where FEntryID="&Request.Form("FEntryID")(i))
			elseif Request.Form("FEntryID")(i)<>"" then
				connk3.Execute("update t_dormpersonEntry set person='"&Request.form("person")(i)&"',fbase1='"&Request.form("fbase1")(i)&"',finteger3='"&Request.form("finteger3")(i)&"',FDate1='"&Request.form("FDate1")(i)&"' where FEntryID="&Request.Form("FEntryID")(i))
			elseif Request.Form("person")(i)<>"" then
				connk3.Execute("insert into t_dormpersonEntry values ("&Request.Form("FID")&","&i&",'"&Request.form("person")(i)&"',"&Request.form("finteger3")(i)&",'"&Request.form("fbase1")(i)&"','"&Request.form("fdate1")(i)&"')")
			end if
		next
		set rs2=connk3.Execute("select count(1) as a1,fid as  a2 from t_dormpersonEntry where fid="&Request.Form("FID")&" group by fid")
		if rs2.eof then
		connk3.Execute("update t_dormperson set sumperson=0 where fid="&Request.Form("FID"))
		else
		connk3.Execute("update t_dormperson set sumperson="&rs2("a1")&" where fid="&Request.Form("FID"))
		end if
		rs2.close
		set rs2=nothing
		response.write "###"
  elseif detailType="Delete" then
		set rs = server.createobject("adodb.recordset")
		sql="select * from t_dormperson where FID="&request("FID")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurview"),"|1002.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		rs.close
		set rs=nothing 
		connk3.Execute("Delete from t_dormperson where FID="&request("FID"))
		connk3.Execute("Delete from t_dormpersonEntry where FID="&request("FID"))
		response.write "###"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="fid" then
    InfoID=request("InfoID")
		sql="select * from t_dormperson where FID="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
			response.write ("对应单据不存在，请检查！")
			response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":[{"
			for i=0 to rs.fields.count-1
			if IsNull(rs.fields(i).value) then
			response.write (""""&LCASE(rs.fields(i).name) & """:"""&rs.fields(i).value&""",")
			else
			response.write (""""&LCASE(rs.fields(i).name) & """:"""&JsonStr(rs.fields(i).value)&""",")
			end if
			next
			response.write ("""Entrys"":[")
			sql="select a.*,b.fnumber as personid,b.fname as personname,(c.fnumber+'/'+c.fname) as depart from t_dormpersonEntry a left join t_emp b on a.person=b.fitemid left join t_item c on c.fitemclassid=2 and a.fbase1=c.fitemid where a.FID="&InfoID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,connk3,1,1
			do until rs.eof
				response.write "{"
				for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
				response.write (""""&LCASE(rs.fields(i).name) & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&LCASE(rs.fields(i).name) & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
				next
				if IsNull(rs.fields(i).value) then
				response.write (""""&LCASE(rs.fields(i).name) & """:"""&rs.fields(i).value&"""}")
				else
				response.write (""""&LCASE(rs.fields(i).name) & """:"""&JsonStr(rs.fields(i).value)&"""}")
				end if
				rs.movenext
			If Not rs.eof Then
				Response.Write ","
			End If
			loop
			response.write "]}]}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="Emp" then
		sql="select Fitemid,Fnumber,fname from t_Emp where Fnumber like '%"&request("InfoID")&"%' or Fname like '%"&request("InfoID")&"%'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.eof then
			response.Write("###")
			response.End()
		end if
			response.write "{"
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
  elseif detailType="Depart" then
		sql="select Fitemid,Fnumber,fname from t_item where fitemclassid=2 and (Fnumber like '%"&request("InfoID")&"%' or Fname like '%"&request("InfoID")&"%')"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.eof then
			response.Write("###")
			response.End()
		end if
			response.write "{"
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
  end if
elseif showType="Export" then 
	InfoID=request("InfoID")
	sql="select a.fbillno,a.Ftext,a.Ftext1,b.finteger3,c.fnumber,c.fname,d.fname as bumen,b.FDate1 from t_dormperson a left join t_dormpersonentry b on a.FId=b.FId left join t_emp c on b.person=c.fitemid left join t_item d on fitemclassid=2 and d.fitemid=b.FBase1 where 1=1 "
		 if request("lh")<>"" then sql=sql&" and a.Ftext1='"&request("lh")&"' "
		 if request("yg")<>"" then sql=sql&" and (c.FNumber like '%"&request("yg")&"%' or c.FName like '%"&request("yg")&"%') "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
%>
		<table width="100%" border="0" id="editDetails" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
		  <tbody id="TbDetails">
		  <tr height="24">
			<td >单号</td>
			<td >宿舍号</td>
			<td >楼号</td>
			<td >床位</td>
			<td >工号</td>
			<td >姓名</td>
			<td >部门</td>
			<td >入住日期</td>
		  </tr>
<%
	do until rs.eof
%>
	<tr height="24" bgcolor="#EBF2F9">
			<td ><%=rs("fbillno")%></td>
			<td ><%=rs("Ftext")%></td>
			<td ><%=rs("Ftext1")%></td>
			<td ><%=rs("finteger3")%></td>
			<td ><%=rs("fnumber")%></td>
			<td ><%=rs("fname")%></td>
			<td ><%=rs("bumen")%></td>
			<td ><%=rs("FDate1")%></td>
   </tr>
<%
		rs.movenext
	loop
	response.Write("</table>")
	rs.close
	set rs=nothing 
end if

Function getBillNo(ID,id2)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select max("&id2&") as maxno From "&ID
  rs.open sql,connk3,1,1
  getBillNo=rs("maxno")+1
  rs.close
  set rs=nothing
End Function  
Function getUser(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From t_emp where fnumber='"&ID&"'"
  rs.open sql,connk3,1,1
  if rs.bof and rs.eof then
  getUser=""
  else
  getUser=rs("Fname")
  end if
  rs.close
  set rs=nothing
End Function    

%>