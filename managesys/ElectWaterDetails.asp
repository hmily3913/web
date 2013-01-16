<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
if request("showType")="Export" then
	response.ContentType("application/vnd.ms-excel")
	response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
dim Result,StartDate,EndDate,Keyword,Reachsum,unReachsum,Reachper,sqlstr,louhao
Result=request("Result")
StartDate=request("start_date")
if StartDate="" then StartDate=date()
EndDate=request("end_date")
if EndDate="" then EndDate=date()
Keyword=request("Keyword")
louhao=request("louhao")

      Response.Write "[<font color='red'>"&StartDate&"] 年 [<font color='red'>"&EndDate&"]月 [<font color='red'>"&louhao&"] 用 电 扣 费 明 细 表"
 
%>

 <table width="100%" border="1" >
  <tr>
    <td width="40" nowrap ><strong>宿舍</strong></td>
    <td width="50" height="24" nowrap ><strong>部门</strong></td>
    <td width="50" nowrap ><strong>工号</strong></td>
    <td width="50" nowrap ><strong>姓名</strong></td>
    <td width="60" nowrap ><strong>上月水表</strong></td>
    <td width="60" nowrap ><strong>本月水表</strong></td>
    <td width="60" nowrap ><strong>上月电表</strong></td>
    <td width="60" nowrap ><strong>本月电表</strong></td>
    <td width="60" nowrap ><strong>上月热水</strong></td>
    <td width="60" nowrap ><strong>本月热水</strong></td>
    <td width="50" nowrap ><strong>用水量</strong></td>
    <td width="50" nowrap ><strong>用电量</strong></td>
    <td width="50" nowrap ><strong>用热水</strong></td>
    <td width="60" nowrap ><strong>合计费用</strong></td>
    <td width="60" nowrap ><strong>应扣费用</strong></td>
  </tr>
 <%

  dim rs,sql'sql语句
  '获取记录总数
  sql="select * from managesys_ShuidianRpt "&_
"where year="&StartDate&" and period="&EndDate&" and louhao='"&louhao&"'"&_
"order by ftext"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,0,1
  dim totalfee,totalelect '总计
  totalelect=0.0
  totalfee=0.0
    while(not rs.eof)'填充数据到表格
			totalfee=totalfee+cdbl(rs("onefee"))
	  Response.Write "<tr >" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("name1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fnumber")&"</td>"
      Response.Write "<td nowrap>"&rs("name2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("water")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("thiswater")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("elect")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("thiselect")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("lasthotwater")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("thihotwater")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("waterdiff")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("electdiff")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("hotwaterdiff")&"</td>" & vbCrLf
	  
      Response.Write "<td nowrap>"&rs("tempnum")&"</td>" & vbCrLf
	    
      Response.Write "<td nowrap>"&formatnumber(rs("onefee"),1)&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='14' nowrap >&nbsp;合计</td>" & vbCrLf
    Response.Write "<td  nowrap  >"&totalfee&"</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
'    response.write "<tr><td height='50' align='center' colspan='9' nowrap  bgcolor='#EBF2F9'>暂无该月水电信息</td></tr>"
'-----------------------------------------------------------

  rs.close
  set rs=nothing
  %>
  </table>
<%
elseif request("showType")="Count" then
	if Instr(session("AdminPurview"),"|1002.5,")=0 then
		response.Write("你没有权限进行此操作，请检查！")
		response.End()
	end if
	StartDate=request("start_date")
	EndDate=request("end_date")
	if datediff("m",StartDate&"-"&EndDate&"-01",date())>0 then
		response.Write("不能对1个月之前的数据进行此操作！")
		response.End()
	end if
	louhao=request("louhao")
	set rs= connzxpt.Execute("select 1 from managesys_ShuidianRpt where year="&StartDate&" and period="&EndDate&" and louhao='"&louhao&"'")
	if rs.eof then
		response.Write(louhao&"对应"&StartDate&"-"&EndDate&"月未进行计算，是否现在计算？")
		response.End()
	else
		response.Write(louhao&"对应"&StartDate&"-"&EndDate&"月<font color='red'>已进行计算</font>，是否重新计算？")
		response.End()
	end if
elseif request("showType")="doCount" then
	if Instr(session("AdminPurview"),"|1002.5,")=0 then
		response.Write("你没有权限进行此操作，请检查！")
		response.End()
	end if
	StartDate=request("start_date")
	EndDate=request("end_date")
	louhao=request("louhao")
	connzxpt.Execute("Delete from managesys_ShuidianRpt where year="&StartDate&" and period="&EndDate&" and louhao='"&louhao&"'")
  sql="select distinct a.sumperson,a.ftext, t_item.fname as name1,t_emp.fnumber,t_emp.fname as name2, "&_
"d.water,d.thiswater,(d.thiswater-d.water) as waterdiff,d.elect,d.thiselect,(d.thiselect-d.elect) as electdiff, "&_
"d.lasthotwater,d.thihotwater,(d.thihotwater-d.lasthotwater) as hotwaterdiff,c.year,c.period,c.fdecimal2,c.waterprice,c.HotWaterPrice "&_
"from t_dormperson a left join t_dormpersonentry b on "&_
"a.fid=b.fid left join t_electwaterentry d on "&_
"a.ftext=d.sushehao left join t_electwater c on "&_
"d.fid=c.fid left join t_item on "&_
"b.fbase1=t_item.FItemID left join t_emp on "&_
"b.person=t_emp.FItemID "&_
"where a.useflag=1 and a.showflag=1 and c.year="&StartDate&" and c.period="&EndDate&" and a.ftext1='"&louhao&"'"&_
"order by ftext"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  totalelect=0.0
  totalfee=0.0
    while(not rs.eof)'填充数据到表格
	  dim tempnum,onefee,electdiff,fdecimal2,sumperson,waterdiff,hotwaterdiff
	  onefee=0
	  electdiff=0
	  waterdiff=0
	  hotwaterdiff=0
	  sumperson=0
	  if louhao="办公楼宿舍" then
		  fdecimal2=CDbl(rs("fdecimal2"))
		  electdiff=CDbl(rs("electdiff"))
		  sumperson=CDbl(rs("sumperson"))
		  totalelect=totalelect+electdiff
		  tempnum=electdiff*fdecimal2'电费合计
		  if sumperson > 0 and tempnum>(sumperson*15) then
			onefee=(tempnum-(sumperson*15))/sumperson
			totalfee=totalfee+onefee
		  end if
	  elseif louhao="旭日小区22号楼" then
	    sumperson=CDbl(rs("sumperson"))
	    if rs("period")=1 or rs("period")=7 or rs("period")=8 or rs("period")=12 then
		  if sumperson > 0 then
		    waterdiff=CDbl(rs("waterdiff"))-10*sumperson
				electdiff=CDbl(rs("electdiff"))-200*sumperson
			tempnum=waterdiff*CDbl(rs("waterprice"))+electdiff*CDbl(rs("fdecimal2"))
			if tempnum<0 then tempnum=0
			onefee=tempnum/sumperson
			totalfee=totalfee+onefee
		  end if
		else
		  if sumperson > 0 then
		    waterdiff=CDbl(rs("waterdiff"))-8*sumperson
				electdiff=CDbl(rs("electdiff"))-150*sumperson
			tempnum=waterdiff*CDbl(rs("waterprice"))+electdiff*CDbl(rs("fdecimal2"))
			if tempnum<0 then tempnum=0
			onefee=tempnum/sumperson
			totalfee=totalfee+onefee
		  end if
		end if
	  elseif louhao="旭日小区4号楼" then
	    sumperson=CDbl(rs("sumperson"))
			if sumperson > 0 then
				tempnum=CDbl(rs("waterdiff"))*CDbl(rs("waterprice"))+CDbl(rs("electdiff"))*CDbl(rs("fdecimal2"))+CDbl(rs("hotwaterdiff"))*CDbl(rs("HotWaterPrice"))
				onefee=tempnum/sumperson-20
				if onefee<3 then onefee=0
				totalfee=totalfee+onefee
			end if
	  end if

		connzxpt.Execute("insert into managesys_ShuidianRpt(year,period,louhao,ftext,name1,fnumber,name2,water,thiswater,elect,thiselect,lasthotwater,thihotwater,waterdiff,electdiff,hotwaterdiff,waterprice,fdecimal2,HotWaterPrice,tempnum,onefee,Biller,BillerID,BillDate) values ("&StartDate&","&EndDate&",'"&louhao&"','"&rs("ftext")&"','"&rs("name1")&"','"&rs("fnumber")&"','"&rs("name2")&"',"&rs("water")&","&rs("thiswater")&","&rs("elect")&","&rs("thiselect")&","&rs("lasthotwater")&","&rs("thihotwater")&","&rs("waterdiff")&","&rs("electdiff")&","&rs("hotwaterdiff")&","&rs("waterprice")&","&rs("fdecimal2")&","&rs("HotWaterPrice")&","&tempnum&","&onefee&",'"&AdminName&"','"&UserName&"','"&now()&"')")
	  rs.movenext
    wend
  rs.close
  set rs=nothing
	response.Write("计算成功！")
end if
%>


