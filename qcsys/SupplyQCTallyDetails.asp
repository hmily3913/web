<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurview"),"|409,")=0 then 
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
  dim datawhere'数据条件
	datawhere=" WHERE a.FBase=b.fitemid and a.FBase4=c.fitemid and a.FBase1=d.fitemid and FDecimal4>0 and a.FBase2=e.fitemid "
	if request("AllQuery")<>"" then
		datawhere=datawhere&" and "&request("AllQuery")
	else
		datawhere=datawhere&" and datediff(month,a.FDate,getdate())=0"
	end if

  dim i:i=1
  dim rs,sql,idCount,page,temstr,pages
	page=clng(request("page"))
	pages=request("rp")
	temstr="###"
  sql="select count(1) as idCount  from t_supplycheck a,"&_
" t_Supplier b ,t_emp c,t_icitem d,t_measureUnit e "&datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  idCount=rs("idCount")

%>
{"page":"<%=page%>","total":"<%=idcount%>","rows":[
<%
    sql="select a.fdate,b.fname as provider,d.fnumber,d.fname as product,a.FText3,FDecimal,FDecimal4,FDecimal5,c.Fname,e.FName as UnitName,FDecimal4-FDecimal5 as cha,round((FDecimal4-FDecimal5)/FDecimal4,2) as bili from t_supplycheck a,"&_
" t_Supplier b ,t_emp c,t_icitem d,t_measureUnit e "&datawhere&_
" order by a.fdate asc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
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
		"cell":["<%=rs("fdate")%>","<%=rs("provider")%>","<%=rs("fnumber")%>","<%=rs("product")%>","<%=rs("FText3")%>","<%=rs("UnitName")%>","<%=rs("FDecimal")%>","<%=rs("FDecimal4")%>","<%=rs("FDecimal5")%>","<%=rs("cha")%>","<%=rs("bili")%>","<%=rs("Fname")%>"]}
<%		
	    rs.movenext
    next
  rs.close
  set rs=nothing
response.Write"]}"
end if
 %>
