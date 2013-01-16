<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp" -->
<%
if Instr(session("AdminPurview"),"|120.5,")=0 then 
  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
  response.end
end if
'========判断是否具有管理权限'
dim showType,detailType
showType=request("showType")
detailType=request("detailType")
if showType="DetailsList" then 
  dim FDate,Department,PostName,PersonName,ItemName,wherestr
	wherestr=""
	FDate=request("FMonth")&"#"&request("FYear")
	if request("Department")<>"" then wherestr=wherestr&" and a.Department='"&request("Department")&"'"
	if request("PostName")<>"" then wherestr=wherestr&" and a.PostName='"&request("PostName")&"'"
	if request("PersonName")<>"" then wherestr=wherestr&" and c.PersonName='"&request("PersonName")&"'"
	if request("ItemName")<>"" then wherestr=wherestr&" and d.ItemName='"&request("ItemName")&"'"
	
	dim rs,sql,rs2,sql2
  dim page'页码
      page=clng(request("page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=request("rp")
  dim pagec'总页数

  dim i'用于循环的整数
	
  '获取记录总数
  sql="select count(1) as idCount "&_
	" from parametersys_PerformancePost a,"&_
	" parametersys_PostToItem b,"&_
	" parametersys_PostToPerson c,"&_
	" parametersys_PerformanceItem d"&_
	" where a.SerialNum=b.PSnum"&_
	" and a.SerialNum=c.PSnum"&_
	" and b.ItemSnum=d.SerialNum"&wherestr
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,0,1
  idCount=rs("idCount")
  if(idcount>0) then'如果记录总数=0,则不处理
    if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
	  pagec=int(idcount/pages)'获取总页数
   	else
      pagec=int(idcount/pages)+1'获取总页数
    end if
  end if

%>
{"page":"<%=page%>","total":"<%=idCount%>","rows":[
<%

	sql="select a.Department,a.PostName,a.Competent,c.PersonId,c.PersonName,d.ItemName,d.TargetData as oneTarget,d.ItemUnit,"&_
	" d.DataTable,d.DataField,d.DelayDays,b.AddSubFlag,b.AddSubRate,b.TargetData as actTarget,b.Score"&_
	" from parametersys_PerformancePost a,"&_
	" parametersys_PostToItem b,"&_
	" parametersys_PostToPerson c,"&_
	" parametersys_PerformanceItem d"&_
	" where a.SerialNum=b.PSnum"&_
	" and a.SerialNum=c.PSnum"&_
	" and b.ItemSnum=d.SerialNum"&wherestr&_
	" order by a.Department asc,a.Competent desc,a.PostName,c.PersonId,d.ItemName"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
	for i=1 to rs.pagesize
	if rs.eof then exit for  
		dim startdate,realData,strRealData,realGetNum
		startdate=dateadd("d",rs("DelayDays"),Split(getDateRangebyMonth(FDate),"###")(1))
		'如果是部门主管，从总表里取数'
		if rs("Competent")=1 then
		sql2="select top 1 "&rs("DataField")&" as realData from "&rs("DataTable")&" where uptdate>='"&startdate&"' order by uptdate asc"
		else
		sql2="select top 1 本月有效率 as realData from "&rs("DataTable")&"_"&rs("DataField")&"Sum where 个人='"&rs("PersonName")&"' and uptdate>='"&startdate&"' order by uptdate asc"
		end if
		set rs2=server.createobject("adodb.recordset")
		rs2.open sql2,connzxpt,1,1
		if rs2.bof and rs2.eof then
		realData=0
		else
		realData=rs2("realData")
		end if
		rs2.close
		set rs2=nothing
		'计算实际得分：实际得分=目标得分-增减幅度*(目标达成-实际达成)'
		if InStr(realData, ".")>0 then
			strRealData=Left(realData, InStr(realData, ".") - 1)
		else
			strRealData=realData
		End if
		realGetNum=Cint(rs("actTarget"))-cdbl(rs("Score"))*(Cint(strRealData)-Cint(rs("oneTarget")))/cdbl(rs("AddSubRate"))
		if realGetNum<0 then realGetNum=0
		if realGetNum>Cint(rs("actTarget")) then realGetNum=rs("actTarget")
	if(i=1)then
%>		
		{"id":"<%=i%>",
		"cell":["<%=rs("Department")%>","<%=rs("PostName")%>","<%=rs("PersonId")%>","<%=rs("PersonName")%>","<%=rs("ItemName")%>","<%=rs("actTarget")%>","<%=rs("oneTarget")%>","<%=realData%>","<%=realGetNum%>","<%=Split(getDateRangebyMonth(FDate),"###")(1)%>"]}
<%		
	else
%>		
		,{"id":"<%=i%>",
		"cell":["<%=rs("Department")%>","<%=rs("PostName")%>","<%=rs("PersonId")%>","<%=rs("PersonName")%>","<%=rs("ItemName")%>","<%=rs("actTarget")%>","<%=rs("oneTarget")%>","<%=realData%>","<%=realGetNum%>","<%=Split(getDateRangebyMonth(FDate),"###")(1)%>"]}
<%		
	end if
	rs.movenext
	next
	response.Write"]}"
	rs.close
	set rs=nothing
elseif showType="getInfo" then 
	if detailType="PersonName" then
		sql="select PersonId,PersonName from parametersys_PostToPerson"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.write "["
		do until rs.eof
		Response.Write("{")
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r")&""",")
				end if
			next
			if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
			else
				response.write (""""&rs.fields(i).name & """:"""&replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r")&"""}")
			end if
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
	elseif detailType="ItemName" then
		sql="select SerialNum,ItemName from parametersys_PerformanceItem"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.write "["
		do until rs.eof
		Response.Write("{")
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r")&""",")
				end if
			next
			if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
			else
				response.write (""""&rs.fields(i).name & """:"""&replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r")&"""}")
			end if
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
	elseif detailType="PostName" then
		sql="select SerialNum,PostName from parametersys_PerformancePost"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.write "["
		do until rs.eof
		Response.Write("{")
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r")&""",")
				end if
			next
			if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
			else
				response.write (""""&rs.fields(i).name & """:"""&replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r")&"""}")
			end if
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
	end if
end if
%>