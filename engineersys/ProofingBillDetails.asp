<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurview"),"|410,")=0 then 
'  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
'  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,UserName,AdminName
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")

if showType="DetailsList" then 
  dim datawhere'数据条件
	datawhere=" where 1=1 "
	if request("AllQuery")<>"" then
		datawhere=datawhere&" and "&request("AllQuery")
	end if

  dim i:i=1
  dim rs,sql,idCount,page,temstr,pages
	page=clng(request("page"))
	pages=request("rp")
	temstr="###"
  sql="select count(1) as idCount  from [K-打样单查询] "&datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,0,1
  idCount=rs("idCount")

%>
{"page":"<%=page%>","total":"<%=idcount%>","rows":[
<%
    sql="select 车间名称,业务员姓名,客户代号,难易度,打样单号,品号,类型,订单日期,业务交期,开发回复交期,预计订单量,预计订单日期,数量,打样员,材料说明 from [K-打样单查询] "&datawhere&_
" order by 打样单号 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
		rs.pagesize = pages
		rs.absolutepage = page  
		for i=1 to rs.pagesize
		if rs.eof then exit for  
		if(i=1)then
		else
		  Response.Write ","
		end if
%>		
		{"id":"<%=i%>",
		"cell":["<%=rs("车间名称")%>","<%=rs("业务员姓名")%>","<%=rs("客户代号")%>","<%=rs("打样单号")%>","<%=rs("品号")%>","<%=rs("类型")%>","<%=rs("难易度")%>","<%=rs("订单日期")%>","<%=rs("业务交期")%>","<%=rs("开发回复交期")%>","<%=rs("预计订单量")%>","<%=rs("预计订单日期")%>","<%=rs("数量")%>","<%=rs("打样员")%>","<%=rs("材料说明")%>"]}
<%		
	    rs.movenext
    next
  rs.close
  set rs=nothing
response.Write"]}"
end if
 %>
