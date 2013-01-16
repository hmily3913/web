<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|105,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,rs,sql
showType=request("showType")
print_tag=request("print_tag")
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
if showType="DetailsList" then 
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9" class='DataCol'><font color="#FFFFFF"><strong>日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" class='DataCol'><font color="#FFFFFF"><strong>通知单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" class='DataCol'><font color="#FFFFFF"><strong>型号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" class='DataCol'><font color="#FFFFFF"><strong>数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" class='DataCol'><font color="#FFFFFF"><strong>业务员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>颜色</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>样品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>营销交期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>评审日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>分厂</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>生管交期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工艺重点要求</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>铜模</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>菲林片</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>模治具</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物料状况</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>采购周期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>品质</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>退单原因</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工程日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>BOM表制作员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物控下单日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>延误天数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>备注</strong></font></td>
  </tr>
 <%
  dim Depart,wherestr,seachword
  dim page'页码
      page=clng(request("Page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=15
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" t_dhtzd a inner join t_dhtzdentry b on a.fid=b.fid and a.fuser>0 left join zxpt.dbo.Flw_OrderReview e on b.FEntryid=e.FEntryid "
  dim datawhere'数据条件
		 datawhere="where 1=1 "&wherestr
		 if request("sd")<>"" then datawhere=datawhere&" and datediff(d,a.FDate1,'"&request("sd")&"')<=0 "
		 if request("ed")<>"" then datawhere=datawhere&" and datediff(d,a.FDate1,'"&request("ed")&"')>=0 "
		 if request("os")<>"" then datawhere=datawhere&" and b.finteger1="&request("os")
		 if request("tdid")<>"" then datawhere=datawhere&" and a.fbillno like '%"&request("tdid")&"%'"
		 if request("psd")<>"" then datawhere=datawhere&" and e.PSDate is not null and datediff(d,e.PSDate,'"&request("psd")&"')<=0 "
		 if request("ped")<>"" then datawhere=datawhere&" and e.PSDate is not null and datediff(d,e.PSDate,'"&request("ped")&"')>=0 "
		 if request("fc")<>"" then datawhere=datawhere&" and e.Fenchang is not null and e.Fenchang like '%"&request("fc")&"%' "
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" order by a.fid desc,b.fentryid asc"
  dim i'用于循环的整数
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
    sql="select b.fentryid from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = 15 '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("fentryid")
	  else
	    sqlid=sqlid &","&rs("fentryid")
	  end if
	  rs.movenext
    next
  '获取本页需要用到的id结束============================================
'-----------------------------------------------------------
'-----------------------------------------------------------
  if sqlid<>"" then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    dim sql2,rs2,hejikoufen
	hejikoufen=0
	dim formdata(3),bgcolors
    sql="select b.fentryid as eid,a.FBillNo,a.FDate1,d.fname as name2,"&_
	"b.FDate11,f.fname as name4,b.FText7,b.FQty,b.FInteger,e.* "&_
	" from t_dhtzd a inner join t_dhtzdentry b on a.fid=b.fid inner join t_emp d on a.FBase3=d.fitemid inner join t_ICItem f on b.FBase=f.fitemid left join zxpt.dbo.Flw_OrderReview e on b.FEntryid=e.FEntryid "&_
	" where b.fentryid in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while(not rs.eof)'填充数据到表格
  	  bgcolors="#EBF2F9"
			dim PMDate,PSDate
			PMDate=""
			PSDate=""
			if rs("PSDate")<>"" and datediff("d",rs("PSDate"),"1900-01-01")<0 then
				bgcolors="#ff99ff"'粉色
				PSDate=rs("PSDate")
			end if
			if rs("PMDate")<>"" and datediff("d",rs("PMDate"),"1900-01-01")<0 then
				PMDate=rs("PMDate")
			end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand' id="&rs("eid")&">" & vbCrLf
      Response.Write "<td nowrap class='DataCol'>"&rs("FDate1")&"</td>" & vbCrLf
      Response.Write "<td nowrap class='DataCol'>"&rs("FBillNo")&"</td>" & vbCrLf
      Response.Write "<td nowrap class='DataCol'>"&rs("name4")&"</td>" & vbCrLf
      Response.Write "<td nowrap class='DataCol'>"&rs("FQty")&"</td>" & vbCrLf
      Response.Write "<td nowrap class='DataCol'>"&rs("name2")&"</td>"
      Response.Write "<td nowrap>"&rs("FText7")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FInteger")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FDate11")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=PSDate>"&PSDate&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Fenchang>"&rs("Fenchang")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=PMDate>"&PMDate&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Standard>"&rs("Standard")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Tongmo>"&rs("Tongmo")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Feilin>"&rs("Feilin")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Mozhi>"&rs("Mozhi")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Wuliao>"&rs("Wuliao")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Caigou>"&rs("Caigou")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Pinzhi>"&rs("Pinzhi")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Tuidan>"&rs("Tuidan")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Gongcheng>"&rs("Gongcheng")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Bom>"&rs("Bom")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Wukong>"&rs("Wukong")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Yanchi>"&rs("Yanchi")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Remark>"&rs("Remark")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
  else
    response.write "<tr><td height='50' align='center' colspan='22' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
	response.Write("###"&pagec&"###"&idcount&"###")
	if Instr(session("AdminPurviewFLW"),"|105.1,")=0 and Instr(session("AdminPurviewFLW"),"|105.2,")=0 then 
		response.Write("0")
	else
		response.Write("1")
	end if
elseif showType="showDetails" then 

elseif showType="DataProcess" then 

elseif showType="getInfo" then 
	sql="select * from Flw_OrderReview where FEntryid="&request("InfoID")
	set rs = server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	if Instr(session("AdminPurviewFLW"),"|105.1,")>0  then 
		if rs.bof and rs.eof then
			connzxpt.Execute("insert into Flw_OrderReview (FEntryid,"&request("detailType")&") values ("&request("InfoID")&",'"&request("values")&"')")
		else
			connzxpt.Execute("update Flw_OrderReview set "&request("detailType")&"='"&request("values")&"' where FEntryid="&request("InfoID"))
		end if
	elseif Instr(session("AdminPurviewFLW"),"|105.2,")>0 and (request("detailType")="Wuliao" or request("detailType")="Caigou") then 
		if rs.bof and rs.eof then
			connzxpt.Execute("insert into Flw_OrderReview (FEntryid,"&request("detailType")&") values ("&request("InfoID")&",'"&request("values")&"')")
		else
			connzxpt.Execute("update Flw_OrderReview set "&request("detailType")&"='"&request("values")&"' where FEntryid="&request("InfoID"))
		end if
	else
		response.Write("你没有权限进行此操作，请检查！")
		response.End()
	end if
	rs.close
	set rs=nothing 
elseif showType="Export" then 
%>
 <table width="100%" border="1" cellpadding="3" cellspacing="1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><strong>日期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>通知单号</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>型号</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>数量</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>业务员</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>颜色</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>样品</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>营销交期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>评审日期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>分厂</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>生管交期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>工艺重点要求</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>铜模</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>菲林片</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>模治具</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>物料状况</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>采购周期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>品质</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>退单原因</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>工程日期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>BOM表制作员</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>物控下单日期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>延误天数</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>备注</strong></td>
  </tr>
 <%
    sql="select b.fentryid as eid,a.FBillNo,a.FDate1,d.fname as name2,"&_
	"b.FDate11,f.fname as name4,b.FText7,b.FQty,b.FInteger,e.* "&_
	" from t_dhtzd a inner join t_dhtzdentry b on a.fid=b.fid inner join t_emp d on a.FBase3=d.fitemid inner join t_ICItem f on b.FBase=f.fitemid left join zxpt.dbo.Flw_OrderReview e on b.FEntryid=e.FEntryid "&_
	" where 1=1 "
		 if request("sd")<>"" then sql=sql&" and datediff(d,a.FDate1,'"&request("sd")&"')<=0 "
		 if request("ed")<>"" then sql=sql&" and datediff(d,a.FDate1,'"&request("ed")&"')>=0 "
		 if request("os")<>"" then sql=sql&" and b.finteger1="&request("os")
		 if request("tdid")<>"" then sql=sql&" and a.fbillno like '%"&request("tdid")&"%'"
		 if request("psd")<>"" then sql=sql&" and e.PSDate is not null and datediff(d,e.PSDate,'"&request("psd")&"')<=0 "
		 if request("ped")<>"" then sql=sql&" and e.PSDate is not null and datediff(d,e.PSDate,'"&request("ped")&"')>=0 "
		 if request("fc")<>"" then sql=sql&" and e.Fenchang is not null and e.Fenchang like '%"&request("fc")&"%' "
		 sql=sql&" order by a.fid desc,b.fentryid asc "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while(not rs.eof)'填充数据到表格
			PMDate=""
			PSDate=""
			if rs("PSDate")<>"" and datediff("d",rs("PSDate"),"1900-01-01")<0 then
				PSDate=rs("PSDate")
			end if
			if rs("PMDate")<>"" and datediff("d",rs("PMDate"),"1900-01-01")<0 then
				PMDate=rs("PMDate")
			end if
	  Response.Write "<tr>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FDate1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FBillNo")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("name4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FQty")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("name2")&"</td>"
      Response.Write "<td nowrap>"&rs("FText7")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FInteger")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FDate11")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=PSDate>"&PSDate&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Fenchang>"&rs("Fenchang")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=PMDate>"&PMDate&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Standard>"&rs("Standard")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Tongmo>"&rs("Tongmo")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Feilin>"&rs("Feilin")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Mozhi>"&rs("Mozhi")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Wuliao>"&rs("Wuliao")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Caigou>"&rs("Caigou")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Pinzhi>"&rs("Pinzhi")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Tuidan>"&rs("Tuidan")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Gongcheng>"&rs("Gongcheng")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Bom>"&rs("Bom")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Wukong>"&rs("Wukong")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Yanchi>"&rs("Yanchi")&"</td>" & vbCrLf
      Response.Write "<td nowrap class=Remark>"&rs("Remark")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
  %>
  </table>
<% 
end if
 %>