<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|214,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName,Depart
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")
print_tag=request("print_tag")
Depart=session("Depart")
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
      datafrom=" purchasesys_MaterialQuote "
  dim datawhere'数据条件
    datawhere=" where 1=1 "
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = request.Form("qtype")
		if Request.Form("qtype")="RegDate" then
			datawhere = datawhere&" and datediff(d,"&searchcols&",'"&searchterm&"')=0 "
		else
			datawhere = datawhere&" and "&searchcols&" like '%"&searchterm&"%' "
		end if
	End if
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "SerialNum" 
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
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
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
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="select SerialNum from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("SerialNum")
	  else
	    sqlid=sqlid &","&rs("SerialNum")
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
    sql="select * from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格
	dim ys:ys="#f7f7f7"
	if rs("CheckFlag")>0 then ys="#ffff66"
	dim CheckState:CheckState=""
	if rs("CheckFlag")="1" then
	  CheckState="主管审核"
	elseif rs("CheckFlag")="2" then
	  CheckState="副总审批"
	elseif rs("CheckFlag")="3" then
	  CheckState="已执行"
	else
	  CheckState="待审核"
	end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("RegDate")%>","<%=rs("Saler")%>","<%=rs("Product")%>","<%=rs("Material")%>","<%=rs("Model")%>","<%=rs("Price")%>","<%=JsonStr(rs("QualityStatus"))%>","<%=rs("MiniQty")%>","<%=JsonStr(rs("EnvironType"))%>","<%=rs("EnvironPrice")%>","<%=rs("TestFee")%>","<%=rs("ValidDays")%>","<%=JsonStr(rs("Remark"))%>","<%=rs("Biller")%>"]}
<%		
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
  end if
  rs.close
  set rs=nothing
response.Write"]}"
'-----------------------------------------------------------
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
  	if  Instr(session("AdminPurview"),"|208.1,")>0 then
			for   i=2   to   Request.form("SerialNum").count
				if (Request.form("DeleteFlag")(i)="0" or Request.form("DeleteFlag")(i)="") and Request.form("Material")(i)<>"" then
					connzxpt.Execute("insert into purchasesys_MaterialQuote select '"&Request.form("RegDate")&"','"&Request.form("Saler")&"','"&Request.form("Product")&"','"&Request.form("Material")(i)&"','"&Request.form("Model")(i)&"',"&Request.form("Price")(i)&",'"&Request.form("QualityStatus")(i)&"',"&Request.form("MiniQty")(i)&",'"&Request.form("EnvironType")(i)&"',"&Request.form("EnvironPrice")(i)&","&Request.form("TestFee")(i)&","&Request.form("ValidDays")(i)&",'"&Request.form("Remark")(i)&"','"&AdminName&"','"&UserName&"','"&now()&"'")
				end if
			next
			response.write "保存成功！"
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit"  then
		if Instr(session("AdminPurview"),"|208.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("Delete from purchasesys_MaterialQuote where SerialNum="&Request.Form("SerialNum")(i))
			elseif Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("update purchasesys_MaterialQuote set RegDate='"&Request.form("RegDate")&"',Saler='"&Request.form("Saler")&"',Product='"&Request.form("Product")&"',Material='"&Request.form("Material")(i)&"',Model='"&Request.form("Model")(i)&"',Price="&Request.form("Price")(i)&",QualityStatus='"&Request.form("QualityStatus")(i)&"',MiniQty="&Request.form("MiniQty")(i)&",EnvironType='"&Request.form("EnvironType")(i)&"',EnvironPrice="&Request.form("EnvironPrice")(i)&",Remark='"&Request.form("Remark")(i)&"',TestFee="&Request.form("TestFee")(i)&",ValidDays="&Request.form("ValidDays")(i)&",BillDate=getdate() where SerialNum="&Request.Form("SerialNum")(i))
			elseif Request.form("Material")(i)<>"" then
					connzxpt.Execute("insert into purchasesys_MaterialQuote select '"&Request.form("RegDate")&"','"&Request.form("Saler")&"','"&Request.form("Product")&"','"&Request.form("Material")(i)&"','"&Request.form("Model")(i)&"',"&Request.form("Price")(i)&",'"&Request.form("QualityStatus")(i)&"',"&Request.form("MiniQty")(i)&",'"&Request.form("EnvironType")(i)&"',"&Request.form("EnvironPrice")(i)&","&Request.form("TestFee")(i)&","&Request.form("ValidDays")(i)&",'"&Request.form("Remark")(i)&"','"&AdminName&"','"&UserName&"','"&now()&"'")
			end if
		next
		response.write "修改成功！"
  elseif detailType="Delete" then
		if Instr(session("AdminPurview"),"|208.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
			connzxpt.Execute("Delete from purchasesys_MaterialQuote where SerialNum in ("&request("SerialNum")&")")
		response.write "删除成功！"
	end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from purchasesys_MaterialQuote where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			response.write ("对应单据不存在，请检查！")
			response.end
		else
			response.write "{""fieldValue"":[{"
			for i=0 to rs.fields.count-1
			if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
			end if
			next
			response.write """Info"":""###""}]}"
		end if
		rs.close
		set rs=nothing 
	end if
end if
 %>
