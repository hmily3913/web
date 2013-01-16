<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|506,")=0 then 
'  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
'  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
'  response.end
'end if
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
      datafrom=" [N-签呈表] a left join [N-签合表单身] b on a.编号=b.编号 left join [N-签呈奖惩明细] c on a.编号=c.编号 "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where 1=1 "
		if Instr(session("AdminPurviewFLW"),"|508.1,")=0 then
			datawhere=datawhere&" and (a.员工代号='"&UserName&"' or b.员工代号='"&UserName&"' or c.员工代号='"&UserName&"' or a.BillerID='"&UserName&"')"
		end if
		datawhere = datawhere&Session("AllMessage69")
		session.contents.remove "AllMessage69"
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
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "a.编号" 
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
  sql="select count(distinct a.编号) as idCount from "& datafrom &" " & datawhere
'	response.Write(sql)
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,0,1
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
    sql="select distinct a.编号 from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid="'"&rs("编号")&"'"
	  else
	    sqlid=sqlid &",'"&rs("编号")&"'"
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
    sql="select 编号,员工代号,员工姓名,部门名称,签呈日期,主题,影响程度,类别,ProjectType,状态,case when 确认=1 then '已确认' else '未确认' end as confirm,审核,BillDate,Biller from [N-签呈表] a where 编号 in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    do until rs.eof'填充数据到表格'
%>		
		{"id":"<%=rs("编号")%>",
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
		set rs = server.createobject("adodb.recordset")
		sql="select * from [N-签呈表]"
		rs.open sql,conn,1,3
		rs.addnew
		dim billymd,billid
		billymd=(YEAR(request.Form("签呈日期"))*10000+MONTH(request.Form("签呈日期"))*100+DAY(request.Form("签呈日期"))) mod 1000000
		set rs2=conn.Execute("select max(编号) as ids from [N-签呈表] where CAST(编号 as varchar) like '"&billymd&"%'")
		if cdbl(rs2("ids"))=0 then
			billid=billymd*10000+1
		else
		billid=cdbl(rs2("ids"))+1
		end if
		rs("编号")=billid
		rs("签呈日期")=request.Form("签呈日期")
		rs("员工代号")=request.Form("员工代号")
		rs("员工姓名")=request.Form("员工姓名")
		rs("部门代号")=request.Form("部门代号")
		rs("部门名称")=request.Form("部门名称")
		rs("影响程度")=request.Form("影响程度")
		rs("性质")=request.Form("性质")
		rs("类别")=request.Form("类别")
		rs("主题")=request.Form("主题")
		rs("签呈内容")=request.Form("签呈内容")
		rs("签呈内容1")=request.Form("签呈内容1")
		rs("状态")="待审核"
		rs("LossMoney")=request.Form("LossMoney")
		rs("ProjectType")=request.Form("ProjectType")
		rs("Biller")=AdminName
		rs("BillerID")=UserName
		rs("BillDate")=now()
		rs.update
		for   i=2   to   Request.form("d1_no").count
			if request.Form("d1_员工代号")(i)<>"" and request.Form("d1_员工姓名")(i)<>"" then
			set rs = server.createobject("adodb.recordset")
			sql="select * from [N-签合表单身]"
			rs.open sql,conn,1,3
			rs.addnew
			rs("项次")=i-1
			rs("员工代号")=request.Form("d1_员工代号")(i)
			rs("员工姓名")=request.Form("d1_员工姓名")(i)
			if request.Form("d1_序号")(i)="" then
				rs("序号")=1
			else
				rs("序号")=request.Form("d1_序号")(i)
			end if
			rs("编号")=billid
			rs("可修改")=request.Form("d1_可修改")(i)
			rs("审核")=request.Form("d1_审核")(i)
			rs("职等")=request.Form("d1_职等")(i)
			rs.update
			end if
		next
		for   i=2   to   Request.form("d2_no").count
			if request.Form("d2_员工代号")(i)<>"" and request.Form("d2_姓名")(i)<>"" then
			set rs = server.createobject("adodb.recordset")
			sql="select * from [N-签呈奖惩明细]"
			rs.open sql,conn,1,3
			rs.addnew
			rs("项次")=i-1
			rs("员工代号")=request.Form("d2_员工代号")(i)
			rs("姓名")=request.Form("d2_姓名")(i)
			rs("奖点数")=request.Form("d2_奖点数")(i)
			rs("预奖点")=request.Form("d2_奖点数")(i)
			rs("惩点数")=request.Form("d2_惩点数")(i)
			rs("预奖点")=request.Form("d2_惩点数")(i)
			rs("编号")=billid
			rs("奖惩项目")=request.Form("d2_奖惩项目")(i)
			rs("事由")=request.Form("d2_事由")(i)
			rs("按照规定")=request.Form("d2_按照规定")(i)
			rs("工作岗位")=request.Form("d2_工作岗位")(i)
			rs("职等")=request.Form("d2_职等")(i)
			rs.update
			end if
		next
		response.write "保存成功！"
    SerialNum=billid
		sql="select * from [N-签呈表] where 编号 ="&SerialNum
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,3
			set rs2=conn.Execute("select 1 from [N-签合表单身] where (序号 is null or 员工代号 is null or 序号='' or 员工代号='') and 编号 ="&SerialNum)
			if (not rs2.eof) then
				response.Write("你输入的会签序号或员工代号不能为空,确认失败！")
				response.End()
			end if
			set rs2=conn.Execute("select count(1) as c from [N-签合表单身] where 审核=1 and 编号 ="&SerialNum)
			if rs2("c")<>1 then
				response.Write("审核人只能并且只能是一个人,确认失败！")
				response.End()
				set rs2=conn.Execute("select 1 from [N-签合表单身] a where exists (select 1 from [N-签合表单身] where 审核=1 and 编号 ="&SerialNum&" and a.序号>序号 and a.编号=编号") 
				if not rs2.eof then
					response.Write("请选择最后一个序号为审核人,确认失败！")
					response.End()
				end if
			end if
			rs("确认")=1
			rs("确认时间")=now()
			rs.update
		rs.close
		set rs=nothing 
'		response.write "保存成功！提示：需确认后才生效！"
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
		sql="select * from [N-签呈表] where 编号="&request.Form("编号")
		rs.open sql,conn,1,3
		billid=rs("编号")
		set rs2=conn.Execute("select 是否已签合 from [N-签合表单身] where 可修改=1 and 员工代号='"&UserName&"' and 是否已签合=0 and 编号="&rs("编号"))
		if not rs2.eof then
			if rs2("是否已签合") then
			response.Write(rs2("是否已签合")&"##")
			response.Write("你已经签合，不允许修改！")
			response.End()
			end if
		end if
		if ((not rs("确认")) and rs("BillerID")=UserName) or (not rs2.eof) then 
		rs("签呈日期")=request.Form("签呈日期")
		rs("员工代号")=request.Form("员工代号")
		rs("员工姓名")=request.Form("员工姓名")
		rs("部门代号")=request.Form("部门代号")
		rs("部门名称")=request.Form("部门名称")
		rs("影响程度")=request.Form("影响程度")
		rs("性质")=request.Form("性质")
		rs("类别")=request.Form("类别")
		rs("主题")=request.Form("主题")
		rs("签呈内容")=request.Form("签呈内容")
		rs("签呈内容1")=request.Form("签呈内容1")
		rs("LossMoney")=request.Form("LossMoney")
		rs("ProjectType")=request.Form("ProjectType")
		rs("ChangeDate")=now()
		rs.update
		for   i=2   to   Request.form("d1_no").count
			if request.Form("d1_员工代号")(i)<>"" and request.Form("d1_员工姓名")(i)<>"" then
			set rs = server.createobject("adodb.recordset")
			sql="select * from [N-签合表单身]"
			if Request.form("d1_no")(i)<>"" then sql=sql&" where no="&Request.form("d1_no")(i)
			rs.open sql,conn,1,3
			if Request.form("d1_no")(i)="" then rs.addnew
			rs("项次")=i-1
			rs("员工代号")=request.Form("d1_员工代号")(i)
			rs("员工姓名")=request.Form("d1_员工姓名")(i)
			rs("序号")=request.Form("d1_序号")(i)
			rs("编号")=billid
			rs("可修改")=request.Form("d1_可修改")(i)
			rs("审核")=request.Form("d1_审核")(i)
			rs("职等")=request.Form("d1_职等")(i)
			rs.update
			end if
		next
		for   i=2   to   Request.form("d2_no").count
			if request.Form("d2_员工代号")(i)<>"" and request.Form("d2_姓名")(i)<>"" then
			set rs = server.createobject("adodb.recordset")
			sql="select * from [N-签呈奖惩明细]"
			if Request.form("d2_no")(i)<>"" then sql=sql&" where no="&Request.form("d2_no")(i)
			rs.open sql,conn,1,3
			if Request.form("d2_no")(i)="" then rs.addnew
			rs("项次")=i-1
			rs("员工代号")=request.Form("d2_员工代号")(i)
			rs("姓名")=request.Form("d2_姓名")(i)
			rs("奖点数")=request.Form("d2_奖点数")(i)
			rs("预奖点")=request.Form("d2_奖点数")(i)
			rs("惩点数")=request.Form("d2_惩点数")(i)
			rs("预奖点")=request.Form("d2_惩点数")(i)
			rs("编号")=billid
			rs("奖惩项目")=request.Form("d2_奖惩项目")(i)
			rs("事由")=request.Form("d2_事由")(i)
			rs("按照规定")=request.Form("d2_按照规定")(i)
			rs("工作岗位")=request.Form("d2_工作岗位")(i)
			rs("职等")=request.Form("d2_职等")(i)
			rs.update
			end if
		next
		end if
		rs.close
		set rs=nothing 
		response.write "保存成功！"
  elseif detailType="Confirm"  then
    SerialNum=request("SerialNum")
		sql="select * from [N-签呈表] where 编号 ="&SerialNum
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,3
		If rs("BillerID")<>UserName then
			response.Write("只能确认自己登记的单据，请检查！")
			response.End()
		end if
		if request("firmtag")="1" then
			if rs("确认") then
				response.Write("单据已经确认不需要重复确认！")
				response.End()
			end if
			set rs2=conn.Execute("select 1 from [N-签合表单身] where (序号 is null or 员工代号 is null or 序号='' or 员工代号='') and 编号 ="&SerialNum)
			if (not rs2.eof) then
				response.Write("你输入的会签序号或员工代号不能为空,确认失败！")
				response.End()
			end if
			set rs2=conn.Execute("select count(1) as c from [N-签合表单身] where 审核=1 and 编号 ="&SerialNum)
			if rs2("c")<>1 then
				response.Write("审核人只能并且只能是一个人,确认失败！")
				response.End()
				set rs2=conn.Execute("select 1 from [N-签合表单身] a where exists (select 1 from [N-签合表单身] where 审核=1 and 编号 ="&SerialNum&" and a.序号>序号 and a.编号=编号") 
				if not rs2.eof then
					response.Write("请选择最后一个序号为审核人,确认失败！")
					response.End()
				end if
			end if
			rs("确认")=1
			rs("确认时间")=now()
			rs.update
			rs.close
			set rs=nothing
			response.write "确认成功！"
		elseif request("firmtag")="0" then
			if (not rs("确认")) or rs("审核") then
				response.Write("单据未确认或已审核，不允许取消确认！")
				response.End()
			end if
			set rs2=conn.Execute("select 1 from [N-签合表单身] where 是否已签合=1 and 编号 ="&SerialNum)
			if not rs2.eof then
				response.Write("已进入签合流程，不允许取消确认！")
				response.End()
			end if
			rs("确认")=0
			rs("确认时间")=null
			rs.update
			rs.close
			set rs=nothing
			response.write "取消成功！"
		end if
  elseif detailType="Delete" then
    SerialNum=request("SerialNum")
		sql="select * from [N-签呈表] where 编号 ="&SerialNum
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		If rs("BillerID")<>UserName then
			response.Write("只能删除自己的单据，请检查！")
			response.End()
		end if
		If rs("确认") then
			response.Write("已确认不允许删除，请检查！")
			response.End()
		end if
		conn.Execute("Delete from [N-签呈表] where 编号 ="&SerialNum)
		response.write "删除成功！"
  elseif detailType="DeleteDetails" then
		SerialNum=request("SerialNum")
		sql="select 确认,BillerID from [N-签呈表] where 编号 ="&request("No")
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		If rs("BillerID")<>UserName then
			response.Write("只能删除自己的单据，请检查！")
			response.End()
		end if
		If rs("确认") then
			response.Write("已确认不允许删除，请检查！")
			response.End()
		end if
		if request("Type")="t1" then
		conn.Execute("Delete from [N-签合表单身] where no ="&SerialNum)
		response.write "删除成功！"
		elseif request("Type")="t2" then
		conn.Execute("Delete from [N-签呈奖惩明细] where no ="&SerialNum)
		response.write "删除成功！"
		end if
  elseif detailType="CheckDetails" then
		SerialNum=request("SerialNum")
		if request("Type")="Agree" or request("Type")="Disagree" then
			set rs=conn.Execute("select 确认,员工代号 from [N-签呈表] where 编号 ="&request("No"))
			If not rs("确认") then
				response.Write("未确认不允许审核,请检查！")
				response.End()
			end if
			dim MailPerson
			MailPerson=rs("员工代号")
			sql="select * from [N-签合表单身] where no ="&SerialNum
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			if rs("是否已签合") then
				response.Write("已签合不允许重复操作，请检查！")
				response.End()
			end if
			if rs("员工代号")<>UserName then
				response.Write("你没有权限审核此签合，请检查！")
				response.End()
			end if
			set rs2=conn.Execute("select 1 from [N-签合表单身] where 是否已签合=0 and 序号<"&rs("序号")&" and 编号 ="&request("No"))
			if not rs2.eof then
				response.Write("签合顺序未轮到本次签合，请等待其他人员签合！")
				response.End()
			end if
			if rs("审核") then
				if request("Type")="Agree" then
					rs("同意")=1
					rs("不同意")=0
					rs("意见")="同意."&request("CheckText")
					conn.Execute("update [N-签呈表] set 审核 =1 ,审核日期='" & now() & "',状态='已审核' where 编号 ="&request("No"))
				elseif request("Type")="Disagree" then
					rs("同意")=0
					rs("不同意")=1
					rs("意见")="不同意."&request("CheckText")
					conn.Execute("update [N-签呈表] set 审核 =1 ,审核日期='" & now() & "',状态='已退回' where 编号 ="&request("No"))
					set rs2=connk3.Execute("select distinct FEmail from t_Base_Emp where FNumber='"&MailPerson&"'")
					if not rs2.eof then
						SendMail rs2("FEmail"),"您的签呈被退回请查看",request("No"),request("CheckText"),""
					end if
				end if
				rs("是否已签合")=1
				rs("时间")=now()
				rs("审核人")=UserName
				rs.update
				rs.close
				set rs=nothing
			else
				if request("Type")="Agree" then
					rs("同意")=1
					rs("不同意")=0
					rs("意见")="同意."&request("CheckText")
				elseif request("Type")="Disagree" then
					rs("同意")=0
					rs("不同意")=1
					rs("意见")="不同意."&request("CheckText")
					sql="select Femail,主题,签呈内容,签呈内容1 from [N-签呈表] a, [N-签合表单身] b,AIS20081217153921.dbo.t_Base_Emp c  where a.编号=b.编号 and b.员工代号=c.FNumber and a.编号 ="&request("No")&" and b.审核 =1"
					set rst=server.createobject("adodb.recordset")
					rst.open sql,conn,1,1
					if not rst.eof then
					SendMail rst("Femail"),"签呈审核不同意",request("No"),rst("主题")&"<br>"&rst("签呈内容")&"<br>"&rst("签呈内容1")&"<br>不同意."&request("CheckText"),""
					end if
				end if
				rs("是否已签合")=1
				rs("时间")=now()
				rs("意见")=request("CheckText")
				rs("审核人")=UserName
				rs.update
				rs.close
				set rs=nothing
				set rs2=conn.Execute("select 1 from [N-签合表单身] where 是否已签合=0 and 编号 ="&request("No"))
				if rs2.eof then
					conn.Execute("update [N-签呈表] set 审核 =1 ,审核日期='" & now() & "',状态='已审核' where 审核=0 and 编号 ="&request("No"))
				end if
			end if
			response.Write("审核成功！")
		'反审核
		elseif request("Type")="Uncheck" then
			sql="select * from [N-签合表单身] where no ="&SerialNum
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			if not rs("是否已签合") then
				response.Write("未签合不需要反审核，请检查！")
				response.End()
			end if
			if rs("员工代号")<>UserName then
				response.Write("你没有权限进行此步骤反审核，请检查！")
				response.End()
			end if
			set rs2=conn.Execute("select 1 from [N-签合表单身] where 是否已签合=1 and 序号>"&rs("序号")&" and 编号 ="&request("No"))
			if not rs2.eof then
				response.Write("签合顺序中以有后续审核不允许您反审核，请检查！")
				response.End()
			end if
			if rs("审核") then
				conn.Execute("update [N-签呈表] set 审核 =0 ,审核日期=null,状态='待审核' where 编号 ="&request("No"))
			end if
			rs("同意")=0
			rs("不同意")=0
			rs("是否已签合")=0
			rs("时间")=null
			rs("意见")=request("CheckText")
			rs("审核人")=UserName
			rs.update
			rs.close
			set rs=nothing
			response.Write("反审核成功！")
		end if
  elseif detailType="EmailDetails" then
		sql="select Femail,主题,签呈内容,签呈内容1 from [N-签呈表] a, [N-签合表单身] b,AIS20081217153921.dbo.t_Base_Emp c  where a.编号=b.编号 and b.员工代号=c.FNumber and a.编号 ="&request("No")&" and b.no ="&request("SerialNum")&" and a.BillerID ='"&UserName&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.eof then
			response.Write("邮箱不能为空，或者单据号不存在，只能制单人才能发送邮件！")
			response.End()
		else
		SendMail rs("Femail"),"签呈审核通知",request("No"),rs("主题")&"<br>"&rs("签呈内容")&"<br>"&rs("签呈内容1"),""
		response.Write("邮件发送成功，请确保单据已确认！")
		end if
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
	if detailType="编号" then
			InfoID=request("InfoID")
		sql="select * from [N-签呈表] where 编号 ="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
			if rs.bof and rs.eof then
					response.write ("对应单据不存在，请检查！")
					response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":{"
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
			end if
			next
			if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""},")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""},")
			end if
			response.write ("""t1"":[")
			sql="select * from  [N-签合表单身]  where 编号 ="&InfoID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			do until rs.eof
				response.write "{"
				for i=0 to rs.fields.count-1
				if IsNull(rs.fields(i).value) then
					if rs.fields(i).type="11" then
						response.write ("""d1_"&rs.fields(i).name & """:"&LCase(cstr(rs.fields(i).value))&",")
					else
						response.write ("""d1_"&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
					end if
				else
					if rs.fields(i).type="11" then
						response.write ("""d1_"&rs.fields(i).name & """:"&LCase(cstr(rs.fields(i).value))&",")
					else
						response.write ("""d1_"&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
					end if
				end if
				next
				response.write ("""bg"":""#EBF2F9""}")
				rs.movenext
			If Not rs.eof Then
				Response.Write ","
			End If
			loop
			response.write "],""t2"":["
			sql="select * from  [N-签呈奖惩明细]  where 编号 ="&InfoID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,1
			do until rs.eof
				response.write "{"
				for i=0 to rs.fields.count-1
				if IsNull(rs.fields(i).value) then
					if rs.fields(i).type="11" then
						response.write ("""d2_"&rs.fields(i).name & """:"&LCase(cstr(rs.fields(i).value))&",")
					else
						response.write ("""d2_"&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
					end if
				else
					if rs.fields(i).type="11" then
						response.write ("""d2_"&rs.fields(i).name & """:"&LCase(cstr(rs.fields(i).value))&",")
					else
						response.write ("""d2_"&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
					end if
				end if
				next
				response.write ("""bg"":""#EBF2F9""}")
				rs.movenext
			If Not rs.eof Then
				Response.Write ","
			End If
			loop
			response.write ("]}")
		end if
		rs.close
		set rs=nothing 
	elseif detailType="员工代号" then
    InfoID=request("InfoID")
		sql="select a.员工代号,a.姓名,a.部门别,a.工作岗位,a.职等,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and (a.员工代号 like '%"&InfoID&"%' or a.姓名 like '%"&InfoID&"%')and 是否离职=0 "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
			if rs.bof and rs.eof then
					response.write ("员工编号不存在！")
					response.end
		else
			response.write(rs("员工代号")&"###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###"&rs("工作岗位")&"###"&rs("职等"))
		end if
		rs.close
		set rs=nothing 
	elseif detailType="getJd" then
		sql="select 内容,参考奖惩 from [G-奖惩基本资料] where 类别='奖点'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		response.write "["
		do until rs.eof
		Response.Write("{""nr"":"""&JsonStr(rs("内容"))&""",""ckjc"":"""&JsonStr(rs("参考奖惩"))&"""}")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
	elseif detailType="getCd" then
		sql="select 内容,参考奖惩 from [G-奖惩基本资料] where 类别='惩点'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		response.write "["
		do until rs.eof
		Response.Write("{""nr"":"""&JsonStr(rs("内容"))&""",""ckjc"":"""&JsonStr(rs("参考奖惩"))&"""}")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
	elseif detailType="getAllPerson" then
		sql="select 员工代号,姓名 from [N-基本资料单头] where 离职否='在职' "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,0,1
		response.write "["
		do until rs.eof
		Response.Write("{""FNumber"":"""&rs("员工代号")&""",""FName"":"""&JsonStr(rs("姓名"))&"""}")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
  end if
elseif showType="Export" then 
	%>
<div id="listtable" style="width:100%; height:420; overflow:scroll">
<table>
<tr>    <td height="20" width="100%" class="tablemenu" colspan="12"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="$('#listtable').hide().css('z-index','550');" >&nbsp;<strong>页面查看明细</strong></font></td>
</tr>
<tr bgcolor="#99BBE8">
  <td>序号</td>
  <td>部门</td>
  <td>员工代号</td>
  <td>姓名</td>
  <td>薪资别</td>
  <td>奖惩</td>
  <td>事由</td>
  <td><%
	if request("Lb")="4.厂服" then
		response.Write("金额")
	else
		response.Write("奖点")
	end if
	%></td>
  <td>惩点</td>
  <td>签呈日期</td>
  <td>审核日期</td>
  <td>签呈单号</td>
  </tr>
	<%
	sql="select d.部门名称,b.员工代号,b.姓名,b.奖惩项目,b.事由,b.奖点数,b.惩点数,a.签呈日期,a.审核日期,a.编号,c.薪资别 "
	sql=sql&" from [N-签呈表] a inner join [N-签呈奖惩明细] b on a.编号=b.编号 left join "
	sql=sql&"  [N-基本资料单头] c on b.员工代号=c.员工代号 left join [G-部门资料表] d on c.部门别=d.部门代号 where 1=1 "
	dim checkdatestr
	if request("ConfirmStat")<>"" then sql=sql&" and a.确认 ="&request("ConfirmStat")&" "
	if request("CheckStat")="1" then
		sql=sql&" and a.状态 = '已审核' "
		checkdatestr="a.审核日期"
	elseif request("CheckStat")="0" then
		sql=sql&" and a.状态 = '待审核' "
		checkdatestr="a.签呈日期"
	elseif request("CheckStat")="2" then
		sql=sql&" and a.状态 = '已退回' "
		checkdatestr="a.签呈日期"
	else
		checkdatestr="a.签呈日期"
	end if
	if request("SDate")<>"" then sql=sql&" and datediff(d,"&checkdatestr&",'"&request("SDate")&"')<=0 "
	if request("EDate")<>"" then sql=sql&" and datediff(d,"&checkdatestr&",'"&request("EDate")&"')>=0 "
	if request("Dpt")<>"" then sql=sql&" and d.部门名称 like '%"&request("Dpt")&"%' "
	if request("UId")<>"" then sql=sql&" and b.员工代号 like '%"&request("UId")&"%' "
	if request("Xzb")<>"" then sql=sql&" and c.薪资别='"&request("Xzb")&"' "
	if request("Lb")<>"" then
		sql=sql&" and a.类别='"&request("Lb")&"' "
	else
		sql=sql&" and a.类别<>'4.厂服' and a.类别<>'3.自离' "
	end if
		if Instr(session("AdminPurviewFLW"),"|508.1,")=0 then
			sql=sql&" and (a.员工代号='"&UserName&"' or b.员工代号='"&UserName&"' or a.BillerID='"&UserName&"')"
		end if
	sql=sql&" order by d.部门代号,b.员工代号 "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	dim zjd,zcd
	dim xuhao:xuhao=0
	zjd=0
	zcd=0
	while (not rs.eof)
		zjd=zjd+rs("奖点数")
		zcd=zcd+rs("惩点数")
		xuhao=xuhao+1
		response.Write("<tr  bgcolor=""#EBF2F9"">")
		response.Write("<td>"&xuhao&"</td>")
		response.Write("<td>"&rs("部门名称")&"</td>")
		response.Write("<td>"&rs("员工代号")&"</td>")
		response.Write("<td>"&rs("姓名")&"</td>")
		response.Write("<td>"&rs("薪资别")&"</td>")
		response.Write("<td>"&rs("奖惩项目")&"</td>")
		response.Write("<td>"&rs("事由")&"</td>")
		response.Write("<td>"&rs("奖点数")&"</td>")
		response.Write("<td>"&rs("惩点数")&"</td>")
		response.Write("<td>"&rs("签呈日期")&"</td>")
		response.Write("<td>"&rs("审核日期")&"</td>")
		response.Write("<td>"&rs("编号")&"</td>")
		response.Write("</tr>")
		rs.movenext
	wend
	response.Write("<tr>")
	response.Write("<td colspan='7'>汇总</td>")
	response.Write("<td>"&zjd&"</td>")
	response.Write("<td>"&zcd&"</td>")
	response.Write("<td colspan='3'></td>")
	response.Write("</tr>")
	response.Write("</table>")
	response.Write("</div>")
	rs.close
	set rs=nothing
elseif showType="Print" then 
		response.ContentType("application/vnd.ms-word")
		response.AddHeader "Content-disposition", "attachment; filename=erpData.doc"
	sql=sql&" select * from [N-签呈表] where 编号= "&request("SerialNum")
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if not rs.eof then
	%>
&nbsp;
<table style="MARGIN: auto auto auto -0.75pt; WIDTH: 17cm; BORDER-COLLAPSE: collapse; mso-padding-alt: 0cm 5.4pt 0cm 5.4pt" class="MsoNormalTable " border="0" cellspacing="0" cellpadding="0" width="643">
	<tbody>
		<tr style="HEIGHT: 35.25pt; mso-yfti-irow: 0; mso-yfti-firstrow: yes">
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 477.75pt; PADDING-RIGHT: 5.4pt; HEIGHT: 35.25pt; BORDER-TOP: windowtext 1pt solid; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-top-alt: solid windowtext .5pt; mso-border-bottom-alt: solid windowtext 1.0pt" width="637" colspan="17" nowrap="nowrap">
				<p style="TEXT-ALIGN: center; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="center">
					<strong><span style="font-family:宋体;FONT-SIZE: 16pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">签<span lang="EN-US"><span style="mso-spacerun: yes">&nbsp;&nbsp; </span></span>呈<span lang="EN-US"><o:p></o:p></span></span></strong>
				</p>
			</td>
		</tr>
		<tr style="HEIGHT: 14.25pt; mso-yfti-irow: 1">
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 76.15pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt" width="102" colspan="2" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">制表日期：<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 11.8pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt" width="16" nowrap="nowrap" colspan="15"><%=rs("BillDate")%>
			</td>
			
		</tr>
		<tr style="HEIGHT: 14.25pt; mso-yfti-irow: 2">
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 228.7pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-top-alt: solid windowtext .5pt; mso-border-alt: solid windowtext .5pt" rowspan="2" width="305" colspan="7" nowrap="nowrap">
				<p style="TEXT-ALIGN: center; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="center">
					<span style="font-family:宋体;FONT-SIZE: 16pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">签呈<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 72pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt; mso-border-right-alt: solid windowtext .5pt" width="96" colspan="2" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">签呈日期<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 177.05pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-top-alt: solid windowtext .5pt; mso-border-bottom-alt: solid windowtext .5pt; mso-border-right-alt: solid windowtext .5pt" width="236" colspan="8" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><%=rs("签呈日期")%>　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
		</tr>
		<tr style="HEIGHT: 14.25pt; mso-yfti-irow: 3">
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 72pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt; mso-border-right-alt: solid windowtext .5pt" width="96" colspan="2" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">签呈部门<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 177.05pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-top-alt: solid windowtext .5pt; mso-border-bottom-alt: solid windowtext .5pt; mso-border-right-alt: solid windowtext .5pt" width="236" colspan="8" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><%=rs("部门名称")%>　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
		</tr>
		<tr style="HEIGHT: 30px; ">
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 477.75pt; PADDING-RIGHT: 5.4pt;  BORDER-TOP: #ece9d8; BORDER-RIGHT: black 1pt solid; PADDING-TOP: 0cm; mso-border-top-alt: solid windowtext .5pt; mso-border-right-alt: solid black .5pt; mso-border-left-alt: solid windowtext .5pt" width="637" colspan="17" nowrap="nowrap">

					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><%=rs("签呈内容")%></span>
        
			</td>
		</tr>
		<tr style="HEIGHT: 500px; ">
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 477.75pt; PADDING-RIGHT: 5.4pt;height:500px;  BORDER-TOP: #ece9d8; BORDER-RIGHT: black 1pt solid; PADDING-TOP: 0cm; mso-border-top-alt: solid windowtext .5pt; mso-border-right-alt: solid black .5pt; mso-border-left-alt: solid windowtext .5pt" width="637" colspan="17" nowrap="nowrap">

					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><%=rs("签呈内容1")%></span>
			</td>
		</tr>
		<tr style="HEIGHT: 14.25pt; mso-yfti-irow: 5">
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 41.8pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-left-alt: solid windowtext .5pt" width="56" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 34.35pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm" width="46" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><o:p>&nbsp;</o:p></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 86.6pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm" width="115" colspan="3" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><o:p>&nbsp;</o:p></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 45pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm" width="60" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><o:p>&nbsp;</o:p></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 41.75pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm" width="56" colspan="2" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><o:p>&nbsp;</o:p></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 133.7pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm" width="178" colspan="6" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><o:p>&nbsp;</o:p></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 49.3pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm" width="66" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><o:p>&nbsp;</o:p></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 11.8pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm" width="16" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><o:p>&nbsp;</o:p></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 33.45pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-right-alt: solid windowtext .5pt" width="45" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><%=rs("员工姓名")%><span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
		</tr>
		<tr style="HEIGHT: 14.25pt; mso-yfti-irow: 6">
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 41.8pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt" width="56" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 34.35pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt" width="46" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 86.6pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt" width="115" colspan="3" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 45pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt" width="60" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 41.75pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt" width="56" colspan="2" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 133.7pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt" width="178" colspan="6" nowrap="nowrap">
				<p style="TEXT-ALIGN: left; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="left">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 49.3pt; PADDING-RIGHT: 5.4pt; HEIGHT: 14.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt" width="127" nowrap="nowrap" colspan="3">
				<p style="TEXT-ALIGN: right; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="right">
					<st1:chsdate w:st="on" year="2012" month="2" day="2" islunardate="False" isrocdate="False"><span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><%=rs("签呈日期")%></span></st1:chsdate><span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体"><o:p></o:p></span>
				</p>
			</td>
		</tr>
		<tr style="HEIGHT: 62.25pt; mso-yfti-irow: 7; mso-yfti-lastrow: yes">
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 41.8pt; LAYOUT-FLOW: vertical-ideographic; PADDING-RIGHT: 5.4pt; HEIGHT: 62.25pt; BORDER-TOP: #ece9d8 solid; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt; mso-border-right-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt" width="56" nowrap="nowrap">
				<p style="TEXT-ALIGN: center; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="center">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">批示<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 120.95pt; PADDING-RIGHT: 5.4pt; HEIGHT: 62.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-top-alt: solid windowtext .5pt; mso-border-bottom-alt: solid windowtext .5pt; mso-border-right-alt: solid windowtext .5pt" width="161" colspan="4" nowrap="nowrap">
				<p style="TEXT-ALIGN: center; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="center">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt; BORDER-LEFT: #ece9d8 solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 45pt; LAYOUT-FLOW: vertical-ideographic; PADDING-RIGHT: 5.4pt; HEIGHT: 62.25pt; BORDER-TOP: #ece9d8 solid; BORDER-RIGHT: windowtext 1pt; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt; mso-border-right-alt: solid windowtext .5pt" width="60" nowrap="nowrap">
				<p style="TEXT-ALIGN: center; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="center">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">会办部门<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 108pt; PADDING-RIGHT: 5.4pt; HEIGHT: 62.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-top-alt: solid windowtext .5pt; mso-border-bottom-alt: solid windowtext .5pt; mso-border-right-alt: solid windowtext .5pt" width="144" colspan="4" nowrap="nowrap">
				<p style="TEXT-ALIGN: center; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="center">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt; BORDER-LEFT: #ece9d8 solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 41.8pt; LAYOUT-FLOW: vertical-ideographic; PADDING-RIGHT: 5.4pt; HEIGHT: 62.25pt; BORDER-TOP: #ece9d8 solid; BORDER-RIGHT: windowtext 1pt; PADDING-TOP: 0cm; mso-border-bottom-alt: solid windowtext .5pt; mso-border-right-alt: solid windowtext .5pt" width="56" colspan="3" nowrap="nowrap">
				<p style="TEXT-ALIGN: center; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="center">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">部门主管<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
			<td style="BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: #ece9d8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 124.4pt; PADDING-RIGHT: 5.4pt; HEIGHT: 62.25pt; BORDER-TOP: #ece9d8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-top-alt: solid windowtext .5pt; mso-border-bottom-alt: solid windowtext .5pt; mso-border-right-alt: solid windowtext .5pt" width="166" colspan="4" nowrap="nowrap">
				<p style="TEXT-ALIGN: center; MARGIN: 0cm 0cm 0pt; mso-pagination: widow-orphan" class="MsoNormal" align="center">
					<span style="font-family:宋体;FONT-SIZE: 12pt; mso-font-kerning: 0pt; mso-bidi-font-family: 宋体">　<span lang="EN-US"><o:p></o:p></span></span>
				</p>
			</td>
		</tr>
		<tr height="0">
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="56">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="46">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="16">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="73">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="27">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="60">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="28">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="28">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="68">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="20">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="30">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="17">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="9">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="34">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="66">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="16">&nbsp;
				
			</td>
			<td style="BORDER-BOTTOM: #ece9d8; BORDER-LEFT: #ece9d8; BACKGROUND-COLOR: transparent; BORDER-TOP: #ece9d8; BORDER-RIGHT: #ece9d8" width="78">&nbsp;
				
			</td>
		</tr>
	</tbody>
</table>
<p style="MARGIN: 0cm 0cm 0pt" class="MsoNormal">
	<span style="font-family:Arial;FONT-SIZE: 9pt; mso-bidi-font-size: 10.0pt"><o:p>&nbsp;</o:p></span>
</p>
	<%
	end if
	rs.close
	set rs=nothing
end if
 %>
