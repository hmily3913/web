<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurview"),"|410,")=0 then 
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
	datawhere=" "
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
  sql="select count(1) as idCount  from t_RP_ARPBill a,t_RP_ARPBillEntry b,QMReject c,QMRejectEntry d,t_Supplier e ,t_icitem f "&_
" where a.FBillID=b.FBillID and b.FClassID_SRC=c.FClassTypeID and a.FCustomer=e.FItemid and c.FItemid=f.Fitemid "&_
" and c.FID=d.FId and b.FEntryID_SRC=d.FEntryID and b.FID_SRC=c.FID "&datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  idCount=rs("idCount")

%>
{"page":"<%=page%>","total":"<%=idcount%>","rows":[
<%
    sql="select a.fdate,e.FNumber as providerid,e.fname as provider,f.fnumber as productid,f.fname as product,g.FText,g.FText3,FDecimal,FQty5,FText1,h.FName as DealName,b.famount,a.FNumber as billno from t_RP_ARPBill a,t_RP_ARPBillEntry b,QMReject c,QMRejectEntry d,t_Supplier e ,t_icitem f,t_supplycheck g,QMAGroupInfo h "&_
" where a.FBillID=b.FBillID and b.FClassID_SRC=c.FClassTypeID and a.FCustomer=e.FItemid and c.FItemid=f.Fitemid "&_
" and g.FID=c.FID_SRC and c.FID=d.FId and b.FEntryID_SRC=d.FEntryID and b.FID_SRC=c.FID and d.FDefectHandlingID=h.FID "&datawhere&_
" order by a.FDate asc"
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
		"cell":["<%=rs("fdate")%>","<%=rs("providerid")%>","<%=rs("provider")%>","<%=rs("productid")%>","<%=rs("product")%>","<%=rs("FText")%>","<%=rs("FText3")%>","<%=rs("FDecimal")%>","<%=rs("FQty5")%>","<%=rs("FText1")%>","<%=rs("DealName")%>","<%=rs("famount")%>","<%=rs("billno")%>"]}
<%		
	    rs.movenext
    next
  rs.close
  set rs=nothing
response.Write"]}"
end if
 %>
