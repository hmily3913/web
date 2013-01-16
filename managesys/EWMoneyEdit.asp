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
      datafrom=" t_electwater a inner join t_electwaterEntry b on a.FID=b.FID left join t_dormperson c on b.sushehao=c.Ftext "
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
		 if request("lh")<>"" then datawhere=datawhere&" and c.Ftext1='"&request("lh")&"' "
		 if request("ss")<>"" then datawhere=datawhere&" and a.sushehao like '%"&request("ss")&"%' "
		 if request("nf")<>"" then datawhere=datawhere&" and a.year ='"&request("nf")&"' "
		 if request("yf")<>"" then datawhere=datawhere&" and a.period ='"&request("yf")&"' "
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
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
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
    sql="select b.Fentryid from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("Fentryid")
	  else
	    sqlid=sqlid &","&rs("Fentryid")
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
    sql="select a.fbillno,b.sushehao,a.checkdate,a.year,a.period,a.waterprice,a.FDecimal2,b.water,b.elect,b.lastHotWater,b.thiswater,b.thiselect,b.thiHotWater,a.FID from "& datafrom &" where b.Fentryid in("& sqlid &") "&taxis
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
  	if  Instr(session("AdminPurview"),"|1002.3,")>0 then
		dim rs2,FIDone
		set rs = server.createobject("adodb.recordset")
		sql="select * from t_electwater"
		rs.open sql,connk3,1,3
		rs.addnew
'		rs("FClassTypeId")="257709051"
		set rs2=connk3.Execute("select max(Fid) as FID,max(FbillNo) as FbillNo from t_electwater")
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
		 if Request.Form("checkdate")<>"" then rs("checkdate")=trim(Request.Form("checkdate"))
		  rs("year")=Request.Form("year")
		  rs("period")=Request.Form("period")
		  rs("waterprice")=Request.Form("waterprice")
		  rs("FDecimal2")=Request.Form("FDecimal2")
		  rs("FBiller")=AdminName
		  rs("FDate1")=now()
		  rs("HotWaterPrice")=Request.Form("HotWaterPrice")
		rs.update
		rs.close
		set rs=nothing 
		for   i=1   to   Request.form("FEntryID").count
			if Request.form("sushehao")(i)<>"" then
				'添加子表
				connk3.Execute("insert into t_electwaterEntry (FID,FIndex,sushehao,water,elect,lastHotWater,thiswater,thiselect,thiHotWater) values ("&FIDone&","&i&",'"&Request.form("sushehao")(i)&"','"&Request.form("water")(i)&"','"&Request.form("elect")(i)&"','"&Request.form("lastHotWater")(i)&"','"&Request.form("thiswater")(i)&"','"&Request.form("thiselect")(i)&"','"&Request.form("thiHotWater")(i)&"')")
				'更新宿舍信息表
		connk3.Execute("update t_dormperson set waternum="&Request.form("thiswater")(i)&",electnum="&Request.form("thiselect")(i)&",HotWater="&Request.form("thiHotWater")(i)&" where ftext='"&Request.form("sushehao")(i)&"'")
			end if
		next
		response.write "###"
	else
		response.write ("你没有权限进行此操作！")
		response.end
	end if
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
    FID=request("FID")
		sql="select * from t_electwater where FID="&Request.Form("FID")
		rs.open sql,connk3,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurview"),"|1002.3,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
    end if
		 if Request.Form("checkdate")<>"" then rs("checkdate")=trim(Request.Form("checkdate"))
		  rs("year")=Request.Form("year")
		  rs("period")=Request.Form("period")
		  rs("waterprice")=Request.Form("waterprice")
		  rs("FDecimal2")=Request.Form("FDecimal2")
		  rs("FBiller")=AdminName
		  rs("FDate1")=now()
		  rs("HotWaterPrice")=Request.Form("HotWaterPrice")
		rs.update
		rs.close
		set rs=nothing 
		for   i=1   to   Request.form("FEntryID").count
			if Request.form("DeleteFlag")(i)="1" or (Request.form("sushehao")(i)="" and Request.Form("FEntryID")(i)<>"") then
				connk3.Execute("Delete from t_electwaterEntry where FEntryID="&Request.Form("FEntryID")(i))
			elseif Request.Form("FEntryID")(i)<>"" then
				connk3.Execute("update t_electwaterEntry set sushehao='"&Request.form("sushehao")(i)&"',water='"&Request.form("water")(i)&"',elect='"&Request.form("elect")(i)&"',lastHotWater='"&Request.form("lastHotWater")(i)&"',thiswater='"&Request.form("thiswater")(i)&"',thiselect='"&Request.form("thiselect")(i)&"',thiHotWater='"&Request.form("thiHotWater")(i)&"' where FEntryID="&Request.Form("FEntryID")(i))
			elseif Request.Form("sushehao")(i)<>"" then
				connk3.Execute("insert into t_electwaterEntry (FID,FIndex,sushehao,water,elect,lastHotWater,thiswater,thiselect,thiHotWater) values ("&FIDone&","&i&",'"&Request.form("sushehao")(i)&"','"&Request.form("water")(i)&"','"&Request.form("elect")(i)&"','"&Request.form("lastHotWater")(i)&"','"&Request.form("thiswater")(i)&"','"&Request.form("thiselect")(i)&"','"&Request.form("thiHotWater")(i)&"')")
			end if
		next
		response.write "###"
  elseif detailType="Delete" then
		set rs = server.createobject("adodb.recordset")
		sql="select * from t_electwater where FID="&request("FID")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurview"),"|1002.3,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		rs.close
		set rs=nothing 
		connk3.Execute("Delete from t_electwater where FID="&request("FID"))
		connk3.Execute("Delete from t_electwaterEntry where FID="&request("FID"))
		response.write "###"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="FID" then
    InfoID=request("InfoID")
		sql="select * from t_electwater where FID="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
			response.write ("对应单据不存在，请检查！")
			response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":[{"
			for i=0 to rs.fields.count-1
			if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
			end if
			next
			response.write ("""Entrys"":[")
			sql="select * from t_electwaterEntry where FID="&InfoID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,connk3,1,1
			do until rs.eof
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
			If Not rs.eof Then
				Response.Write ","
			End If
			loop
			response.write "]}]}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="AllEW" then
		response.write ("{""Entrys"":[")
		sql="select ftext as sushehao,waternum as water,electnum as elect,hotWater as lastHotWater from t_dormperson where ftext1='"&request("louhao")&"' order by ftext1,ftext"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		do until rs.eof
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
		If Not rs.eof Then
			Response.Write ","
		End If
		loop
		response.write "]}"
  end if
elseif showType="Export" then 
	InfoID=request("InfoID")
	sql="select a.fbillno,b.sushehao,a.checkdate,a.year,a.period,a.waterprice,a.FDecimal2,b.water,b.elect,b.lastHotWater,b.thiswater,b.thiselect,b.thiHotWater,a.FId,a.hotwaterprice,c.FText1 from t_electwater a inner join t_electwaterEntry b on a.FID=b.FID left join t_dormperson c on b.sushehao=c.Ftext where 1=1 "
		 if request("lh")<>"" then sql=sql&" and c.Ftext1='"&request("lh")&"' "
		 if request("ss")<>"" then sql=sql&" and a.sushehao like '%"&request("ss")&"%' "
		 if request("nf")<>"" then sql=sql&" and a.year ='"&request("nf")&"' "
		 if request("yf")<>"" then sql=sql&" and a.period ='"&request("yf")&"' "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
%>
		<table width="100%" border="0" id="editDetails" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
		  <tbody id="TbDetails">
		  <tr height="24">
			<td >单号</td>
			<td >宿舍号</td>
			<td >楼号</td>
			<td >年份</td>
			<td >月份</td>
			<td >水价</td>
			<td >电价</td>
			<td >热水价</td>
			<td >上月水表度数</td>
			<td >上月电表度数</td>
			<td >上月热水度数</td>
			<td >本月水表度数</td>
			<td >本月电表度数</td>
			<td >本月热水度数</td>
		  </tr>
<%
	do until rs.eof
%>
	<tr height="24" bgcolor="#EBF2F9">
			<td ><%=rs("fbillno")%></td>
			<td ><%=rs("sushehao")%></td>
			<td ><%=rs("Ftext1")%></td>
			<td ><%=rs("year")%></td>
			<td ><%=rs("period")%></td>
			<td ><%=rs("waterprice")%></td>
			<td ><%=rs("FDecimal2")%></td>
			<td ><%=rs("hotwaterprice")%></td>
			<td ><%=rs("water")%></td>
			<td ><%=rs("elect")%></td>
			<td ><%=rs("lastHotWater")%></td>
			<td ><%=rs("thiswater")%></td>
			<td ><%=rs("thiselect")%></td>
			<td ><%=rs("thiHotWater")%></td>
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