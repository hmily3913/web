<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurview"),"|408,")=0 then 
  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
  response.end
end if
'========判断是否具有管理权限
dim showType,UserName,AdminName
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")

if showType="DetailsList" then 
  dim datawhere,FDate'数据条件
	FDate=request("FMonth")&"#"&request("FYear")
	datawhere=" WHERE a.FBase=b.fitemid and datediff(d,FDate,'"&Split(getDateRangebyMonth(FDate),"###")(0)&"')<=0 and datediff(d,FDate,'"&Split(getDateRangebyMonth(FDate),"###")(1)&"')>=0"

  dim i:i=1
  dim rs,sql,idCount,page,temstr
	temstr="###"
  sql="select count(1) as idCount from (select Fnumber,FName,sum(a1) as b1,sum(a3) as b3,sum(a2) as b2,round((cast(sum(a1) as float)/cast(sum(a2) as float)*100),2) as b4 "&_
" from "&_
" (select case when FComboBox2='合格' then 1 else 0 end as a1,case when FComboBox2='合格' then 0 else 1 end as a3,1 as a2,b.Fnumber,b.FName"&_
" from t_supplycheck a,"&_
" t_Supplier b "&datawhere&_
" ) as c"&_
" group by c.Fnumber,c.FName"&_
" ) as d"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  idCount=rs("idCount")

%>
{"page":"1","total":"<%=idcount%>","rows":[
<%
    sql="select Fnumber,FName,sum(a1) as b1,sum(a3) as b3,sum(a2) as b2,round((cast(sum(a1) as float)/cast(sum(a2) as float)*100),2) as b4 "&_
" from "&_
" (select case when FComboBox2='合格' then 1 else 0 end as a1,case when FComboBox2='合格' then 0 else 1 end as a3,1 as a2,b.Fnumber,b.FName"&_
" from t_supplycheck a,"&_
" t_Supplier b "&datawhere&_
" ) as c"&_
" group by c.Fnumber,c.FName"&_
" order by b4 desc,b2 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
		dim b5:b5=0
    do until rs.eof
		if temstr<>rs("b4") then
			b5=b5+1
			temstr=rs("b4")
		end if
%>		
		{"id":"<%=i%>",
		"cell":["<%=i%>","<%=rs("FName")%>","<%=rs("Fnumber")%>","<%=rs("b2")%>","<%=rs("b1")%>","<%=rs("b3")%>","<%=rs("b4")%>","<%=b5%>"]}
<%		
i=i+1
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
  rs.close
  set rs=nothing
response.Write"]}"
'-----------------------------------------------------------'
elseif showType="DataProcess" then 
  	if  Instr(session("AdminPurview"),"|408.1,")>0 then
		set rs = server.createobject("adodb.recordset")
		sql="select * from qcsys_SupplyQCSort where FYear='"&request("FYear")&"' and FMonth='"&request("FMonth")&"'"
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
		rs.addnew
		for i=0 to rs.fields.count-1
		  if rs.fields(i).name ="Biller" then
		  rs.fields(i).value=UserName
		  elseif rs.fields(i).name ="BillDate" then
		  rs.fields(i).value=now()
		  else
		  rs.fields(i).value=Request(rs.fields(i).name)
		  end if
		next
		else
		for i=0 to rs.fields.count-1
		  if rs.fields(i).name ="Biller" then
		  rs.fields(i).value=AdminName
		  elseif rs.fields(i).name ="BillDate" then
		  rs.fields(i).value=now()
		  else
		  rs.fields(i).value=Request(rs.fields(i).name)
		  end if
		next
		end if		
		rs.update
		rs.close
		set rs=nothing 
		response.write "保存成功！"
	else
		response.write ("你没有权限进行此操作！")
		response.end
	end if
elseif showType="getInfo" then 
  request("FMonth")
	sql="select * from qcsys_SupplyQCSort where Fyear='"&request("Fyear")&"' and FMonth='"&request("FMonth")&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	  response.write "{""Info"":""###"",""fieldValue"":[{"
	  for i=0 to rs.fields.count-2
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
		else
		response.write (""""&rs.fields(i).name & """:"""&replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r")&""",")
		end if
	  next
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}]}")
		else
		response.write (""""&rs.fields(i).name & """:"""&replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r")&"""}]}")
		end if
	rs.close
	set rs=nothing 
end if
 %>
