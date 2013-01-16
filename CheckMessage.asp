<!--#include file="Include/ConnSiteData.asp" -->
<%
dim key,UserName,AdminName,Depart,Purview,PurviewFLW,DepartName
dim datawhere
Purview=session("AdminPurview")
PurviewFLW=session("AdminPurviewFLW")
UserName=session("UserName")
AdminName=session("AdminName")
Depart=session("Depart")
key=request("key")
sql="select 部门名称 from [G-部门资料表] where 部门代号='"&Depart&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
DepartName=rs("部门名称")

if key="AllNeed" then 
  dim rs,sql,idCount
  '6S执行情况回复'
'  if Instr(PurviewFLW,"|103.1,")>0 then
'  if Depart="KD01.0001.0001"  then
'	datawhere=" and FComboBox1='人资部'"
'  elseif Depart="KD01.0001.0002" then
'	datawhere=" and FComboBox1='工程部'"
'  elseif Depart="KD01.0001.0003" then
'	datawhere=" and FComboBox1='采购部'"
'  elseif Depart="KD01.0005.0004" then
'	datawhere=" and FComboBox1='营销部'"
'  elseif Depart="KD01.0001.0005" then
'	datawhere=" and FComboBox1='生技部'"
'  elseif Depart="KD01.0001.0006" then
'	datawhere=" and FComboBox1='仓储科'"
'  elseif Depart="KD01.0001.0007" then
'	datawhere=" and FComboBox1='二分厂'"
'  elseif Depart="KD01.0001.0008" then
'	datawhere=" and FComboBox1='三分厂'"
'  elseif Depart="KD01.0001.0009" then
'	datawhere=" and FComboBox1='财务部'"
'  elseif Depart="KD01.0001.0010" then
'	datawhere=" and FComboBox1='一分厂'"
'  elseif Depart="KD01.0001.0011" then
'	datawhere=" and FComboBox1='品保部'"
'  elseif Depart="KD01.0001.0012" then
'	datawhere=" and FComboBox1='总经办'"
'  elseif Depart="KD01.0004.0001" or Depart="KD01.0004.0002" then
'	datawhere=" and FComboBox1='娄桥办'"
'  elseif Depart="KD01.0001.0017" then
'	datawhere=" and FComboBox1='生管部'"
'  elseif Depart="KD01.0001.0018" then
'	datawhere=" and FComboBox1='眼镜布绳'"
'  end if
'  sql="select count(1) as idCount from t_6s a,t_5sEntry b where a.fid=b.fid and a.fuser>0 and Len(FText7)=0 "&datawhere
'  set rs=server.createobject("adodb.recordset")
'  rs.open sql,connk3,1,1
'  if rs("idCount")>0 then
%>
<!--<a href="javascript:parent.SetSession(1,'FLW/6sExecute.asp')" target="mainFrame">(<font color="#FF0000">0</font>)条6S扣分待改善</a><br>-->
<%
'  end if
'  end if
  '薪资福利津贴变动审核'
  if Instr(PurviewFLW,"|201.2,")>0 then
  sql="select count(1) as idCount from Bill_WelfareAdjust where CheckFlag=0 and Department='"&Depart&"'"
	if Left(Depart,9)="KD01.0004" then sql="select count(1) as idCount from Bill_WelfareAdjust where CheckFlag=0 and Department like 'KD01.0004%'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(2,'Bill/WelfareAdjust.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条薪资福利调整单待审核</a><br>
<%
  end if
  end if
  '薪资福利津贴变动确认'
  if Instr(PurviewFLW,"|201.3,")>0 then
  datawhere=" and 1=2"
  if Depart="KD01.0001.0012" then datawhere=" and (shenqxm='话费补贴' or shenqxm='住房补贴' )"
  if Depart="KD01.0001.0001" then datawhere=" and (shenqxm='工资调薪' or shenqxm='岗位补贴' or shenqxm='其他补贴' or shenqxm='职等调整' or shenqxm='工龄恢复' )"
  sql="select count(1) as idCount from Bill_WelfareAdjust where CheckFlag=1 "&datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(3,'Bill/WelfareAdjust.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条薪资福利调整单待确认</a><br>
<%
  end if
  end if
  '薪资福利津贴变动审批'
  if Instr(PurviewFLW,"|201.4,")>0 then
  sql="select count(1) as idCount from Bill_WelfareAdjust where (CheckFlag=2 and (shenqxm='工资调薪' or (shenqxm='职等调整' and ApplicEmployment>4) or shenqxm='其他补贴') ) and (left(Department,9)=left('"&Depart&"',9) or (left('"&Depart&"',9)='KD01.0005' and left(Department,9)<>'KD01.0001') ) "'or (CheckFlag=0 and left(Department,9)='KD01.0004')
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(4,'Bill/WelfareAdjust.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条薪资福利调整单待审批</a><br>
<%
  end if
  end if
  '打样异常反馈'
  if Instr(PurviewFLW,"|202.2,")>0 then
  sql="select count(1) as idCount from Bill_ProofingAbnormal where len(ReplyText)=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(5,'Bill/ProofingAbnormal.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条打样异常反馈待回复</a><br>
<%
  end if
  end if
  '订单异常反馈处理'
  if Instr(PurviewFLW,"|203.2,")>0 then
  sql="select count(1) as idCount from Bill_OrderAbnormal where len(ReplyText)=0 and CheckFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(6,'Bill/OrderAbnormal.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条订单异常反馈待回复</a><br>
<%
  end if
  end if
  '维修申请处理'
  if Instr(PurviewFLW,"|204.2,")>0 then
  sql="select count(1) as idCount from Bill_RepairApplication where CheckFlag=0 and (Biller='"&UserName&"' or Register='"&UserName&"')"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(7,'Bill/RepairApplication.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条维修申请待审核</a><br>
<%
  end if
  end if
  '维修申请处理'
  if Instr(PurviewFLW,"|204.3,")>0 then
  sql="select count(1) as idCount from Bill_RepairApplication where CheckFlag=1 and ReceivDepartment='总经办'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(8,'Bill/RepairApplication.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条维修申请待接收</a><br>
<%
  end if
  end if
  '维修申请处理'
  if Instr(PurviewFLW,"|204.3,")>0 then
  sql="select count(1) as idCount from Bill_RepairApplication where CheckFlag=2 and ReceivDepartment='总经办'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(9,'Bill/RepairApplication.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条维修申请待处理</a><br>
<%
  end if
  end if
  '联络函审核'
  if Instr(PurviewFLW,"|206.2,")>0 then
  sql="select count(1) as idCount from Bill_InternalWorkLetter where CheckFlag=0 and Department='"&Depart&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(10,'Bill/InternalWorkLetter.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条内部工作联络函待审核</a><br>
<%
  end if
  end if
  '联络函审批'
  if Instr(PurviewFLW,"|206.3,")>0 then
  sql="select count(1) as idCount from Bill_InternalWorkLetter where CheckFlag=1"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(11,'Bill/InternalWorkLetter.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条内部工作联络函待审批</a><br>
<%
  end if
  end if
  '联络函会签'
  if Instr(PurviewFLW,"|206.4,")>0 then
  if Depart="KD01.0001.0001"  then
	datawhere=" and ReceivDepartment like '%人资部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0002" then
	datawhere=" and ReceivDepartment like '%工程部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0003" then
	datawhere=" and ReceivDepartment like '%采购部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0005.0004" then
	datawhere=" and ReceivDepartment like '%营销部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0005" then
	datawhere=" and ReceivDepartment like '%生技部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0006" then
	datawhere=" and ReceivDepartment like '%仓储科%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0007" then
	datawhere=" and ReceivDepartment like '%二分厂%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0008" then
	datawhere=" and ReceivDepartment like '%三分厂%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0009" then
	datawhere=" and ReceivDepartment like '%财务部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0010" then
	datawhere=" and ReceivDepartment like '%一分厂%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0011" then
	datawhere=" and ReceivDepartment like '%品保部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0012" then
	datawhere=" and ReceivDepartment like '%总经办%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0017" then
	datawhere=" and ReceivDepartment like '%生管部%' and Signman not like '%"&AdminName&"%'"
  end if
  sql="select count(1) as idCount from Bill_InternalWorkLetter where CheckFlag=2 "&datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(12,'Bill/InternalWorkLetter.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条内部工作联络函待会签</a><br>
<%
  end if
  end if
  '生管物料看板品保接收'
  if Instr(PurviewFLW,"|207.2,")>0 then
  sql="select count(1) as idCount from Bill_PMMTRboard where QCFlag<2"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(13,'Bill/PMMTRboard.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条紧急物料待处理</a><br>
<%
  end if
  end if
  '生管物料看板仓库接收'
  if Instr(PurviewFLW,"|207.3,")>0 then
  sql="select count(1) as idCount from Bill_PMMTRboard where STFlag<2"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(14,'Bill/PMMTRboard.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条紧急物料待处理</a><br>
<%
  end if
  end if
  '发货异常异常回复'
  if Instr(PurviewFLW,"|208.2,")>0 then
  sql="select count(1) as idCount from Bill_FinishDelivAbnormal where len(ReplyText)=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(15,'Bill/FinishDelivAbnormal.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条发货异常待回复</a><br>
<%
  end if
  end if
  '工伤事故部门审核'
  if Instr(PurviewFLW,"|209.2,")>0 then
  sql="select count(1) as idCount from Bill_WorkInjury where CheckFlag=0 and Department='"&Depart&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(16,'Bill/WorkInjury.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条工伤事故待审核</a><br>
<%
  end if
  end if
  '工伤事故人资审核'
  if Instr(PurviewFLW,"|209.3,")>0 then
  sql="select count(1) as idCount from Bill_WorkInjury where CheckFlag=1"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(17,'Bill/WorkInjury.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条工伤事故待确认</a><br>
<%
  end if
  end if
  '工伤事故总监审批'
  if Instr(PurviewFLW,"|209.4,")>0 then
  sql="select count(1) as idCount from Bill_WorkInjury where CheckFlag=4"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(18,'Bill/WorkInjury.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条工伤事故待审批</a><br>
<%
  end if
  end if
  '工伤事故总监审批'
  if Instr(PurviewFLW,"|209.5,")>0 then
  sql="select count(1) as idCount from Bill_WorkInjury where CheckFlag=2"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(19,'Bill/WorkInjury.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条工伤事故待实施</a><br>
<%
  end if
  end if
  '派车单审核'
  if Instr(Purview,"|1003.5,")>0 then
  	datawhere=""
	  if Depart="KD01.0001.0001"  then
		datawhere=" and left(a.fnumber,2)='06' "
	  elseif Depart="KD01.0001.0002" then
		datawhere=" and left(a.fnumber,2)='03' "
	  elseif Depart="KD01.0001.0003" then
		datawhere=" and left(a.fnumber,2)='05' "
	  elseif Depart="KD01.0005.0004" then
		datawhere=" and left(a.fnumber,2)='02' "
	  elseif Depart="KD01.0001.0005" then
		datawhere=" and left(a.fnumber,2)='08' "
	  elseif Depart="KD01.0001.0006" then
		datawhere=" and left(a.fnumber,2)='07' "
	  elseif Depart="KD01.0001.0007" then
		datawhere=" and left(a.fnumber,2)='11' "
	  elseif Depart="KD01.0001.0008" then
		datawhere=" and left(a.fnumber,2)='12' "
	  elseif Depart="KD01.0001.0009" then
		datawhere=" and left(a.fnumber,2)='04' "
	  elseif Depart="KD01.0001.0010" then
		datawhere=" and left(a.fnumber,2)='10' "
	  elseif Depart="KD01.0001.0011" then
		datawhere=" and left(a.fnumber,2)='09' "
	  elseif Depart="KD01.0001.0012" then
		datawhere=" and left(a.fnumber,2)='01' "
	  elseif Depart="KD01.0001.0018" then
		datawhere=" and left(a.fnumber,5)='10.04' "
	  elseif Depart="KD01.0001.0019" then
		datawhere=" and left(a.fnumber,2)='23' "
	  end if
  sql="select count(1) as idCount from t_item a,z_SendCar where RegistDepartment=a.fitemid and DeleteFlag<1 and RejecteFlag!=1 and CheckFlag=0 and not exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber like OutPeron) "&datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(20,'managesys/SendCarMana.asp?Result=Search&Page=1&queryFlag=none')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条外出单待审核</a><br>
<%
  end if
  end if
  '生产任务单反结案审核'
  if Instr(PurviewFLW,"|212.2,")>0 then
  sql="select count(1) as idCount from Bill_ICMOUnEnd where CheckFlag=0 and Department='"&Depart&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(21,'Bill/ICMOUnEnd.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条生产任务反结案申请待审核</a><br>
<%
  end if
  end if
  '生产任务单反结案审批'
  if Instr(PurviewFLW,"|212.3,")>0 then
  sql="select count(1) as idCount from Bill_ICMOUnEnd where CheckFlag=1 and CancelFlag=0 "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(22,'Bill/ICMOUnEnd.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条生产任务反结案申请待审批</a><br>
<%
  end if
  end if
  '生产任务单反结案执行'
  if Instr(PurviewFLW,"|212.4,")>0 then
  sql="select count(1) as idCount from Bill_ICMOUnEnd where CheckFlag=2 and CancelFlag=0 "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(23,'Bill/ICMOUnEnd.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条生产任务反结案申请待执行</a><br>
<%
  end if
  end if
  '生产任务单反结案结案'
  if Instr(PurviewFLW,"|212.4,")>0 then
  sql="select count(1) as idCount from Bill_ICMOUnEnd where CheckFlag=3 and datediff(d,PlanEndDate,getdate())>=0 and CancelFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(24,'Bill/ICMOUnEnd.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条生产任务反结案申请待结案</a><br>
<%
  end if
  end if
  rs.close
  set rs=nothing 
  '人员管制表审核'
  if Instr(PurviewFLW,"|211.2,")>0 then
  sql="select count(1) as idCount from Bill_AllPersonalCtrl where CheckFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(25,'Bill/AllPersonalCtrl.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条人员管制表待审核</a><br>
<%
  end if
  end if
  '加班申请单审核'
  if Instr(PurviewFLW,"|214.2,")>0 then
  sql="select count(1) as idCount from Bill_Overtime where CheckFlag=0 and (Department='"&Depart&"' or (Department='KD01.0001.0005' and '"&Depart&"'='KD01.0001.0012' ))"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(26,'Bill/Overtime.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条加班申请待审核</a><br>
<%
  end if
  end if
  '加班申请单相关部门签核'
  if Instr(PurviewFLW,"|214.3,")>0 then
		if Depart="KD01.0001.0017" then
			sql="select count(1) as idCount from Bill_Overtime where CancelFlag=0 and CheckFlag=1 and (OverType='分厂加班' or OverType='外协加班')"
		elseif Depart="KD01.0001.0003" then
			sql="select count(1) as idCount from Bill_Overtime where CancelFlag=0 and CheckFlag=1 and OverType='收料加班'"
		end if
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(27,'Bill/Overtime.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条加班申请待签核</a><br>
<%
  end if
  end if
  '加班申请单营销总监审批'
  if Instr(PurviewFLW,"|214.4,")>0 then
  sql="select count(1) as idCount from Bill_Overtime where CancelFlag=0 and CheckFlag=1 and (Department='KD01.0005.0004' or OverType='接送客人')"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(28,'Bill/Overtime.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条加班申请待审批</a><br>
<%
  end if
  end if
  '加班申请单营销总监审批'
  if Instr(PurviewFLW,"|214.5,")>0 then
  sql="select count(1) as idCount from Bill_Overtime where CancelFlag=0 and ((CheckFlag=1 and OverType<>'收料加班' and Department<>'KD01.0005.0004') or CheckFlag=2)  and left(Department,9)='"&left(Depart,9)&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(29,'Bill/Overtime.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条加班申请待审批</a><br>
<%
  end if
  end if
  '加班申请单执行'
  if Instr(PurviewFLW,"|214.6,")>0 then
  sql="select count(1) as idCount from Bill_Overtime where CancelFlag=0 and CheckFlag=3 and left(Department,9)='"&left(Depart,9)&"'"
'	if left(Depart,9)="KD01.0005" then
'	sql=sql&" and (left(Department,9)='"&left(Depart,9)&"' or Department='KD01.0001.0009') "
'	else
'	sql=sql&" and left(Department,9)='"&left(Depart,9)&"' "
'	end if
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(30,'Bill/Overtime.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条加班申请待执行</a><br>
<%
  end if
  end if
  '异物料起订量审核'
  if Instr(PurviewFLW,"|107.1,")>0 then
  sql="select count(1) as idCount from z_AbnomalOrder where CheckFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(31,'FLW/AbnomalOrder.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条异物料起订量待审核</a><br>
<%
  end if
  end if
  '异物料起订量回复'
  if Instr(PurviewFLW,"|107.2,")>0 then
  sql="select count(1) as idCount from z_AbnomalOrder where CheckFlag=1 and PCer='' and PCFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(32,'FLW/AbnomalOrder.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条异物料起订量待回复</a><br>
<%
  end if
  end if
  '异物料起订量回复'
  if Instr(PurviewFLW,"|107.4,")>0 then
  sql="select count(1) as idCount from z_AbnomalOrder where CheckFlag=1 and PUer='' and PUFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(33,'FLW/AbnomalOrder.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条异物料起订量待回复</a><br>
<%
  end if
  end if
  '异物料起订量回复'
  if Instr(PurviewFLW,"|107.6,")>0 then
  sql="select count(1) as idCount from z_AbnomalOrder where CheckFlag=1 and SAer='' and SAFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(34,'FLW/AbnomalOrder.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条异物料起订量待回复</a><br>
<%
  end if
  end if
  '异物料起订量回复'
  if Instr(PurviewFLW,"|107.8,")>0 then
  sql="select count(1) as idCount from z_AbnomalOrder where CheckFlag=1 and ENer='' and ENFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(35,'FLW/AbnomalOrder.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条异物料起订量待回复</a><br>
<%
  end if
  end if
  '异物料起订量签核'
  if Instr(PurviewFLW,"|107.3,")>0 then
  sql="select count(1) as idCount from z_AbnomalOrder where CheckFlag=1 and PCFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(36,'FLW/AbnomalOrder.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条异物料起订量待签核</a><br>
<%
  end if
  end if
  '异物料起订量签核'
  if Instr(PurviewFLW,"|107.5,")>0 then
  sql="select count(1) as idCount from z_AbnomalOrder where CheckFlag=1 and PUFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(37,'FLW/AbnomalOrder.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条异物料起订量待签核</a><br>
<%
  end if
  end if
  '异物料起订量签核'
  if Instr(PurviewFLW,"|107.7,")>0 then
  sql="select count(1) as idCount from z_AbnomalOrder where CheckFlag=1 and SAFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(38,'FLW/AbnomalOrder.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条异物料起订量待签核</a><br>
<%
  end if
  end if
  '异物料起订量签核'
  if Instr(PurviewFLW,"|107.9,")>0 then
  sql="select count(1) as idCount from z_AbnomalOrder where CheckFlag=1 and ENFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(39,'FLW/AbnomalOrder.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条异物料起订量待签核</a><br>
<%
  end if
  end if
  '异物料起订量签核'
  if Instr(PurviewFLW,"|107.10,")>0 then
  sql="select count(1) as idCount from z_AbnomalOrder where CheckFlag=1 and VPFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(40,'FLW/AbnomalOrder.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条异物料起订量待签核</a><br>
<%
  end if
  end if
  '薪资福利津贴变动执行'
  if Instr(PurviewFLW,"|201.6,")>0 and (Depart="KD01.0001.0012" or Depart="KD01.0001.0001") then
		sql="select count(1) as idCount from Bill_WelfareAdjust where (CheckFlag>1 and CheckFlag<99 and (1=2 "
		if Instr(session("AdminPurviewFLW"),"|201.7,")>0 then    sql=sql&" or (shenqxm='工资调薪' and CheckFlag>2)"
		if Instr(session("AdminPurviewFLW"),"|201.8,")>0 then    sql=sql&" or shenqxm='话费补贴'"
		if Instr(session("AdminPurviewFLW"),"|201.9,")>0 then    sql=sql&" or shenqxm='住房补贴'"
		if Instr(session("AdminPurviewFLW"),"|201.10,")>0 then    sql=sql&" or shenqxm='岗位补贴'"
		if Instr(session("AdminPurviewFLW"),"|201.11,")>0 then    sql=sql&" or (shenqxm='其他补贴' and CheckFlag>2)"
		if Instr(session("AdminPurviewFLW"),"|201.12,")>0 then    sql=sql&" or (shenqxm='职等调整' and ApplicEmployment<5) or (shenqxm='职等调整' and ApplicEmployment>4 and CheckFlag>2)"
		if Instr(session("AdminPurviewFLW"),"|201.13,")>0 then    sql=sql&" or shenqxm='工龄恢复'"
		sql=sql&"))"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(41,'Bill/WelfareAdjust.asp')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条薪资福利调整单待执行</a><br>
<%
  end if
  end if
  '未打卡证明单审核'
  if Instr(PurviewFLW,"|215.2,")>0 then
  sql="select count(1) as idCount from Bill_UnCardProof where CheckFlag=0 and Department='"&Depart&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(42,'Bill/UnCardProof.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条未打卡证明单待审核</a><br>
<%
  end if
  end if
  '未打卡证明单实施'
  if Instr(PurviewFLW,"|215.3,")>0 or Instr(PurviewFLW,"|215.4,")>0 then
	if Instr(PurviewFLW,"|215.3,")>0 then
		sql="select count(1) as idCount from Bill_UnCardProof where CancelFlag=0 and CheckFlag=1 "
		sql=sql&" and Department='"&Depart&"'"
	end if
	if Instr(PurviewFLW,"|215.4,")>0 then
		sql="select count(1) as idCount from Bill_UnCardProof where CancelFlag=0 and CheckFlag=1 "
		if Depart="KD01.0005.0001" then
			sql=sql&" and (Department='KD01.0001.0009' or left(Department,9)='"&left(Depart,9)&"')"
		else
			sql=sql&" and left(Department,9)='"&left(Depart,9)&"' "
		end if
	end if
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(43,'Bill/UnCardProof.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条未打卡证明单待执行</a><br>
<%
  end if
  end if
  '客户验货品保确认'
  if Instr(Purview,"|411.2,")>0 then
	sql="select count(1) as idCount from qcsys_CustomInspect where CheckFlag=0 "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(44,'qcsys/CustomInspect.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条客户验货待确认结果</a><br>
<%
  end if
  end if
  '调休单待部门主管审核'
  if Instr(PurviewFLW,"|216.1,")>0 then
	sql="select count(1) as idCount from Bill_Annualleave where CheckFlag=0 and Department='"&Depart&"' and not exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber=RegisterID)"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(45,'Bill/Annualleave.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条调休单待审核</a><br>
<%
  end if
  end if
  '调休单待总监审核'
  if Instr(PurviewFLW,"|216.2,")>0 then
	sql="select count(1) as idCount from Bill_Annualleave where CheckFlag=0 and left(Department,9)='"&left(Depart,2)&"' and exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber=RegisterID)"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(46,'Bill/Annualleave.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条调休单待审核</a><br>
<%
  end if
  end if
  '调休单待执行'
  if Instr(PurviewFLW,"|216.3,")>0 then
	sql="select count(1) as idCount from Bill_Annualleave where CheckFlag=1 and CancelFlag=0 "
	if Depart="KD01.0005.0001" then
		sql=sql&" and (Department='KD01.0001.0009' or left(Department,9)='"&left(Depart,9)&"')"
	elseif Depart="KD01.0001.0001" then
		sql=sql&" and left(Department,9)='"&left(Depart,9)&"' "
	else
		sql=sql&" and Department='"&Depart&"' "
	end if
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(47,'Bill/Annualleave.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条调休单待执行</a><br>
<%
  end if
  end if
  '请假单待部门主管审核'
  if Instr(PurviewFLW,"|217.1,")>0 then
	sql="select count(1) as idCount from Bill_Leave where CancelFlag=0 and ((CheckFlag=1 and (SalaryType='1.行政月薪' or (TotalDay>6 and (SalaryType='2.分厂月薪' or SalaryType='3.计件')))) or (CheckFlag=0 and (SalaryType='1.行政月薪' or Grade>4))) and Department='"&Depart&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(48,'Bill/Leave.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条请假单待审核</a><br>
<%
  end if
  end if
  '请假单待总监审核'
  if Instr(PurviewFLW,"|217.2,")>0 then
	sql="select count(1) as idCount from Bill_Leave where CheckFlag=2 and left(Department,9)='"&left(Depart,9)&"' and CancelFlag=0 and (exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber=RegisterID) or (TotalDay>6 and SalaryType='1.行政月薪') or ((SalaryType='1.行政月薪' or SalaryType='3.计件') and TotalDay>14))"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(49,'Bill/Leave.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条请假单待审批</a><br>
<%
  end if
  end if
  '请假单待执行'
  if Instr(PurviewFLW,"|217.3,")>0 then
	sql="select count(1) as idCount from Bill_Leave where CancelFlag=0 and (CheckFlag=3 or (CheckFlag=2 and not exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber=RegisterID) and TotalDay<7) or (CheckFlag=1 and TotalDay<7 and (SalaryType='2.分厂月薪' or SalaryType='3.计件')))"
	if Depart="KD01.0005.0001" then
		sql=sql&" and (Department='KD01.0001.0009' or left(Department,9)='"&left(Depart,9)&"') "
	elseif Depart="KD01.0001.0001" then
		sql=sql&" and left(Department,9)='"&left(Depart,9)&"' "
	else
		sql=sql&" and Department='"&Depart&"' "
	end if
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(50,'Bill/Leave.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条请假单待执行</a><br>
<%
  end if
  end if
  '派车单总监审核'
  if Instr(Purview,"|1003.9,")>0 then
  sql="select count(1) as idCount from z_SendCar where DeleteFlag<1 and RejecteFlag!=1 and CheckFlag=0 and len(OutPeron)>0 and exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber like OutPeron)"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(51,'managesys/SendCarMana.asp?Result=Search&Page=1&queryFlag=none')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条外出单待审批</a><br>
<%
  end if
  end if
  '供应商评估表工程意见'
  if Instr(Purview,"|207.2,")>0 then
  sql="select count(1) as idCount from purchasesys_SupplierEvaluat where CheckFlag=0 and EnFlag=0 "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(52,'purchasesys/SupplierEvaluat.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条供应商评估表待评价</a><br>
<%
  end if
  end if
  '供应商评估表工程意见'
  if Instr(Purview,"|207.3,")>0 then
  sql="select count(1) as idCount from purchasesys_SupplierEvaluat where CheckFlag=0 and QcFlag=0 "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(53,'purchasesys/SupplierEvaluat.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条供应商评估表待评价</a><br>
<%
  end if
  end if
  '供应商评估表工程意见'
  if Instr(Purview,"|207.4,")>0 then
  sql="select count(1) as idCount from purchasesys_SupplierEvaluat where CheckFlag=0 and PoFlag=0 "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(54,'purchasesys/SupplierEvaluat.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条供应商评估表待评价</a><br>
<%
  end if
  end if
  '供应商评估表工程意见'
  if Instr(Purview,"|207.5,")>0 then
  sql="select count(1) as idCount from purchasesys_SupplierEvaluat where CheckFlag=0 and EnFlag=1 and QcFlag=1 and PoFlag=1 "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(55,'purchasesys/SupplierEvaluat.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条供应商评估表待审批</a><br>
<%
  end if
  end if
  '请假单待部门工段长审核'
  if Instr(PurviewFLW,"|217.4,")>0 then
	sql="select count(1) as idCount from Bill_Leave where CheckFlag=0 and Department='"&Depart&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(56,'Bill/Leave.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条请假单待审核</a><br>
<%
  end if
  end if
  '印章使用申请审核'
  if Instr(PurviewFLW,"|218.2,")>0 then
	sql="select count(1) as idCount from Bill_StampUse where CheckFlag=0 and Department='"&Depart&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(57,'Bill/StampUse.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条印章使用申请待审核</a><br>
<%
  end if
  end if
  '印章使用申请审批'
  if Instr(PurviewFLW,"|218.3,")>0 then
	sql="select count(1) as idCount from Bill_StampUse where CheckFlag=1 and CancelFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(58,'Bill/StampUse.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条印章使用申请待审批</a><br>
<%
  end if
  end if
  '虚拟网申请'
  if Instr(PurviewFLW,"|219.1,")>0 then
	sql="select count(1) as idCount from Bill_VirtualNetwork where CheckFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(59,'Bill/VirtualNetwork.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条虚拟网申请待执行</a><br>
<%
  end if
  end if
  '出差申请审核'
  if Instr(PurviewFLW,"|701.1,")>0 then
	sql="select count(1) as idCount from Attendance_Travel where CheckFlag=0 and Department='"&Depart&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(60,'Attendance/Travel.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条出差申请待审核</a><br>
<%
  end if
  end if
  '出差申请审核'
  if Instr(PurviewFLW,"|701.2,")>0 then
	sql="select count(1) as idCount from Attendance_Travel where CheckFlag=1 and CancelFlag=0 and Left(Department,9)='"&left(Depart,9)&"' "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(61,'Attendance/Travel.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条出差申请待审批</a><br>
<%
  end if
  end if
  '出差申请执行'
  if Instr(PurviewFLW,"|701.3,")>0 then
	sql="select count(1) as idCount from Attendance_Travel where CheckFlag=2 and CancelFlag=0 and Left(Department,9)='"&left(Depart,9)&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(62,'Attendance/Travel.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条出差申请待执行</a><br>
<%
  end if
  end if
  '工伤事故财务确认'
  if Instr(PurviewFLW,"|209.6,")>0 then
  sql="select count(1) as idCount from Bill_WorkInjury where CheckFlag=3"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(63,'Bill/WorkInjury.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条工伤事故待财务确认</a><br>
<%
  end if
  end if
  '产前试做'
  if Instr(PurviewFLW,"|109.2,")>0 then
  sql="select count(1) as idCount from Flw_PrenatalTest where CheckFlag=0"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(64,'Flw/PrenatalTest.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条产前试做待审核</a><br>
<%
  end if
  end if
  '产前试做'
  if Instr(PurviewFLW,"|109.3,")>0 then
  sql="select count(1) as idCount from Flw_PrenatalTest where CheckFlag=1"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(65,'Flw/PrenatalTest.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条产前试做待确认</a><br>
<%
  end if
  end if
  '产前试做'
  if Instr(PurviewFLW,"|109.4,")>0 then
  sql="select count(1) as idCount from Flw_PrenatalTest where CheckFlag=2"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(66,'Flw/PrenatalTest.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条产前试做待追踪</a><br>
<%
  end if
  end if
  '产前试做'
  if Instr(PurviewFLW,"|109.5,")>0 then
  sql="select count(1) as idCount from Flw_PrenatalTest where CheckFlag=3"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(67,'Flw/PrenatalTest.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条产前试做待确定</a><br>
<%
  end if
  end if
  '产前试做'
  if Instr(PurviewFLW,"|109.6,")>0 then
  sql="select count(1) as idCount from Flw_PrenatalTest where CheckFlag=4"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(68,'Flw/PrenatalTest.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条产前试做待确定</a><br>
<%
  end if
  end if
  '签呈会签'
  sql="select count(distinct a.编号) as idCount from [N-签呈表] a,[N-签合表单身] b where a.编号=b.编号 and a.确认=1 and b.员工代号='"&UserName&"' and b.是否已签合=0  and not exists (select 1 from [N-签合表单身] where a.编号=编号 and 序号<b.序号 and 是否已签合=0 ) "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(69,'OA/Signwas.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条签呈待签合</a><br>
<%
  end if 
  '记账或认证'
  if Instr(Purview,"|808.3,")>0 then
  sql="select count(1) as idCount from Financesys_SupplyBill where (AccountFlag=0 or ApprovFlag=0) and Invoice<>''"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(70,'financesys/SupplyBill.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)供方开票信息待记账或认证</a><br>
<%
  end if
  end if
  '虚拟网结案'
	sql="select count(1) as idCount from Bill_VirtualNetwork where CheckFlag=1 and CancelFlag=0 and (BillerID='"&UserName&"' or RegisterID='"&UserName&"')"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(71,'Bill/VirtualNetwork.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条虚拟网申请待结案</a><br>
<%
  end if
  '计时单段组长审核'
  if Instr(Purview,"|311.2,")>0 then
  sql="select count(1) as idCount from manusys_Jishi where CheckFlag=0 and Bumen='"&DepartName&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(72,'manusys/Jishi.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条员工计时单段组长待审</a><br>
<%
  end if
  end if
  '计时单厂长审核'
  if Instr(Purview,"|311.3,")>0 then
  sql="select count(1) as idCount from manusys_Jishi where CheckFlag=1 and Bumen='"&DepartName&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(73,'manusys/Jishi.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条员工计时单厂长待审</a><br>
<%
  end if
  end if
  '计时单责任审核'
  sql="select count(distinct a.SerialNum) as idCount from manusys_Jishi a,manusys_JishiDetails2 b where a.SerialNum=b.SNum and a.CheckFlag=2 and b.CheckerID='"&UserName&"' and b.CheckFlag=0  "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(74,'manusys/Jishi.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条员工计时单待责任会签</a><br>
<%
  end if
  '计时单执行'
  if Instr(Purview,"|311.6,")>0 then
  sql="select count(1) as idCount from manusys_Jishi where CheckFlag=3 and Bumen='"&DepartName&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
%>
<a href="javascript:parent.SetSession(75,'manusys/Jishi.html')" target="mainFrame">(<font color="#FF0000"><%= rs("idCount") %></font>)条员工计时单待执行</a><br>
<%
  end if
  end if
  rs.close
  set rs=nothing 
elseif key="PartnerState" then
	i=1
  sql="select '外出' as statetype,a.FName,max(b.StarteDate) as outTime from t_emp a,z_SendCar b  "&_
"where b.checkFlag=3 and datediff(d,b.StarteDate,getdate())=0   "&_
"and b.CarryMans>0 and b.OutPeron like '%'+a.fnumber+'%' group by a.FName   "&_
"union all "&_
"select '请假' as statetype,b.FName,CtrlDate as outTime  "&_
"from zxpt.dbo.bill_AllPersonalCtrl as a, "&_
"t_emp as b "&_
" where a.LeaveTo>0 and a.LeaveToMan like '%'+b.fnumber+'%' "&_
" and checkFlag=1 and datediff(d,ctrlDate,getdate())=0 "&_
"union all "&_
"select '旷工' as statetype,b.FName,CtrlDate as outTime  "&_
"from zxpt.dbo.bill_AllPersonalCtrl as a, "&_
"t_emp as b "&_
" where a.DesertTo>0 and a.DesertToMan like '%'+b.fnumber+'%' "&_
" and checkFlag=1 and datediff(d,ctrlDate,getdate())=0 "&_
"order by outtime desc "

  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
	while (not rs.eof)
		response.write ("▲"&i&"."&rs("FName"))
		dim nnn
		for nnn=len(rs("FName")) to 4
			response.write "　"
		next
		response.write rs("statetype")&"　"&rs("outTime")&"<br>"
		rs.movenext
		i=i+1
	wend
  rs.close
  set rs=nothing 
elseif key="longinUser" then
	if request("SerialNum")="" then
		sql="select 部门代号 as SerialNum,部门代号 as UserName,部门名称 as AdminName,1 as flag,3 as o_state,0 as 性别 from [G-部门资料表] where len(部门代号)=14 order by 部门代号 asc"
	else
  sql="select c.SerialNum,c.UserName,c.AdminName,0 as flag,c.o_state,a.性别 from [N-基本资料单头] a,[G-部门资料表] b,zxpt.dbo.smmsys_online as c "&_
" where c.UserName=a.员工代号 and a.部门别=b.部门代号 and b.部门代号='"&request("SerialNum")&"' "&_
"order by c.o_state desc "
	end if
	response.Write("[")
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
	do until rs.eof
	%>
	{"SerialNum": "<%=rs("SerialNum")%>", "UserName": "<%=rs("UserName")%>", "name": "<%=rs("AdminName")%>", "state": "<%=rs("o_state")%>", "gender": "<%=rs("性别")%>"
	<%
		if rs("o_state")=1 then response.write ",iconSkin : ""onlineUser"""
		if rs("flag")=1 then response.write ",isParent:true"
		Response.Write "}"
		rs.movenext
		If Not rs.eof Then
			Response.Write ","
		End If
	loop
	Response.Write "]"
  rs.close
  set rs=nothing 
elseif key="messagelist" then
  dim page'页码
      page=clng(request("page"))
  dim pages'每页条数
      pages=request("rp")
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" smmsys_Message "
  dim i'用于循环的整数
    datawhere=" where ((inceptuserid='"&UserName&"' and delR=0) or (senderuserid='"&UserName&"' and delS=0)) "
	
	if Request.Form("undo") = 1 then
		datawhere=datawhere&" and inceptuserid='"&UserName&"' "
	elseif Request.Form("undo") = 2 then
		datawhere=datawhere&" and senderuserid='"&UserName&"' "
	end if
	
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
		dim temsend,temsendid
		dim temcept,temceptid,newflag,clstrs,clstre
		clstrs="<font color=#000000>"
		clstre="</font>"
		newflag=""
		if rs("senderuserid")=UserName then
			temsendid=""
			temsend="我"
			clstrs="<font color=#42B475>"
			if rs("flag")="0" then
				newflag="未读"
			else
				newflag="已读"
			end if
		else
			temsendid=rs("senderuserid")
			temsend=rs("sender")
		end if
		if rs("inceptuserid")=UserName then
			temceptid=""
			temcept="我"
			clstrs="<font color=#006EFE>"
			if rs("flag")="0" then
				newflag="新"
			end if
		else
			temceptid=rs("inceptuserid")
			temcept=rs("incept")
		end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=clstrs&newflag&clstre%>","<%=temsend%>","<%=temcept%>","<%=clstrs&rs("sendtime")&clstre%>","<%=clstrs&JsonStr(rs("title"))&clstre%>","<%=clstrs&JsonStr(rs("content"))&clstre%>","<%=rs("SerialNum")%>","<%=temceptid%>","<%=temsendid%>"]}
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
elseif key="messSend" then
  Login()

	sql="select * from smmsys_Message"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,3
	rs.addnew
	rs("incept")=Request.Form("incept")
	rs("title")=Request.Form("title")
	rs("content")=Request.Form("content")
	rs("inceptuserid")=Request.Form("inceptuserid")
	rs("sendtime")=now()
	rs("sender")=AdminName
	rs("senderuserid")=UserName
	rs.update
	rs.close
	set rs=nothing
elseif key="messRead" then
	Login()
	sql="update smmsys_Message set flag=1 where SerialNum="&Request.Form("SerialNum")
	connzxpt.Execute (sql)
elseif key="messDel" then
	Login()
	sql="update smmsys_Message set delR=1 where SerialNum in ("&Request.Form("SerialNum")&") and inceptuserid='"&UserName&"' "
	connzxpt.Execute (sql)
	sql="update smmsys_Message set delS=1 where SerialNum in ("&Request.Form("SerialNum")&") and senderuserid='"&UserName&"' "
	connzxpt.Execute (sql)
	response.write "###"
elseif key="CheckNewMessage" then
	sql="select count(1) as idCount from smmsys_Message where flag=0 and inceptuserid='"&UserName&"' and delR=0 "
	set rs = server.createobject("adodb.recordset")
  rs.open sql,connzxpt,1,1
  if rs("idCount")>0 then
		response.write rs("idCount")
  end if
	rs.close
	set rs=nothing
elseif key="showNotice" then
	sql="select * from oa_Announce where SerialNum="&request("SerialNum")
	set rs = server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,3
	if rs.eof and rs.bof then
		response.Write("对应公告已经不存在，请刷新再查看！")
		response.End()
	else
		if Instr(rs("Reader"),UserName)=0 then rs("Reader")=rs("Reader")&"("&UserName&")"
		response.Write(rs("Content"))
	end if
	rs.update
	rs.close
	set rs=nothing
elseif key="SetSession" then
  if Depart="KD01.0001.0001"  then
	datawhere=" and FComboBox1='人资部'"
  elseif Depart="KD01.0001.0002" then
	datawhere=" and FComboBox1='工程部'"
  elseif Depart="KD01.0001.0003" then
	datawhere=" and FComboBox1='采购部'"
  elseif Depart="KD01.0005.0004" then
	datawhere=" and FComboBox1='营销部'"
  elseif Depart="KD01.0001.0005" then
	datawhere=" and FComboBox1='生技部'"
  elseif Depart="KD01.0001.0006" then
	datawhere=" and FComboBox1='仓储科'"
  elseif Depart="KD01.0001.0007" then
	datawhere=" and FComboBox1='二分厂'"
  elseif Depart="KD01.0001.0008" then
	datawhere=" and FComboBox1='三分厂'"
  elseif Depart="KD01.0001.0009" then
	datawhere=" and FComboBox1='财务部'"
  elseif Depart="KD01.0001.0010" then
	datawhere=" and FComboBox1='一分厂'"
  elseif Depart="KD01.0001.0011" then
	datawhere=" and FComboBox1='品保部'"
  elseif Depart="KD01.0001.0012" then
	datawhere=" and FComboBox1='总经办'"
  elseif Depart="KD01.0004.0001" or Depart="KD01.0004.0002" then
	datawhere=" and FComboBox1='娄桥办'"
  elseif Depart="KD01.0001.0017" then
	datawhere=" and FComboBox1='生管部'"
  end if
	select case Request("SessionNum")
	case 1
	Session("AllMessage1")=" and Len(FText7)=0 "&datawhere
	case 2
	Session("AllMessage2")=" and CheckFlag=0 and Department='"&Depart&"'"
	if left(Depart,9)="KD01.0004" then Session("AllMessage2")=" and CheckFlag=0 and Department like 'KD01.0004%'"
	case 3
  if Depart="KD01.0001.0012" then datawhere=" and (shenqxm='话费补贴' or shenqxm='住房补贴' )"
  if Depart="KD01.0001.0001" then datawhere=" and (shenqxm='工资调薪' or shenqxm='岗位补贴' or shenqxm='其他补贴' or shenqxm='职等调整' or shenqxm='工龄恢复' )"
	Session("AllMessage3")=" and CheckFlag=1 "&datawhere
	case 4
	Session("AllMessage4")=" and (CheckFlag=2 and (shenqxm='工资调薪' or shenqxm='其他补贴' or (shenqxm='职等调整' and ApplicEmployment>4))) and (left(Department,9)=left('"&Depart&"',9) or (left('"&Depart&"',9)='KD01.0005' and left(Department,9)<>'KD01.0001') )  "' or (CheckFlag=0 and left(Department,9)='KD01.0004')
	case 5
	Session("AllMessage5")=" and len(ReplyText)=0"
	case 6
	Session("AllMessage6")=" and len(ReplyText)=0 and CheckFlag=0"
	case 7
	Session("AllMessage7")=" and CheckFlag=0 and (Biller='"&UserName&"' or Register='"&UserName&"')"
	case 8
	Session("AllMessage8")=" and CheckFlag=1 and ReceivDepartment='总经办'"
	case 9
	Session("AllMessage9")=" and CheckFlag=2 and ReceivDepartment='总经办'"
	case 10
	Session("AllMessage10")=" and CheckFlag=0 and Department='"&Depart&"'"
	case 11
	Session("AllMessage11")=" and CheckFlag=1"
	case 12
  if Depart="KD01.0001.0001"  then
	datawhere=" and ReceivDepartment like '%人资部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0002" then
	datawhere=" and ReceivDepartment like '%工程部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0003" then
	datawhere=" and ReceivDepartment like '%采购部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0005.0004" then
	datawhere=" and ReceivDepartment like '%营销部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0005" then
	datawhere=" and ReceivDepartment like '%生技部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0006" then
	datawhere=" and ReceivDepartment like '%仓储科%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0007" then
	datawhere=" and ReceivDepartment like '%二分厂%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0008" then
	datawhere=" and ReceivDepartment like '%三分厂%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0009" then
	datawhere=" and ReceivDepartment like '%财务部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0010" then
	datawhere=" and ReceivDepartment like '%一分厂%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0011" then
	datawhere=" and ReceivDepartment like '%品保部%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0012" then
	datawhere=" and ReceivDepartment like '%总经办%' and Signman not like '%"&AdminName&"%'"
  elseif Depart="KD01.0001.0017" then
	datawhere=" and ReceivDepartment like '%生管部%' and Signman not like '%"&AdminName&"%'"
  end if
	Session("AllMessage12")=" and CheckFlag=2 "&datawhere
	case 13
	Session("AllMessage13")=" and QCFlag<2"
	case 14
	Session("AllMessage14")=" and STFlag<2"
	case 15
	Session("AllMessage15")=" and len(ReplyText)=0"
	case 16
	Session("AllMessage16")=" and CheckFlag=0 and Department='"&Depart&"'"
	case 17
	Session("AllMessage17")=" and CheckFlag=1"
	case 18
	Session("AllMessage18")=" and CheckFlag=4"
	case 19
	Session("AllMessage19")=" and CheckFlag=2"
	case 20
	  if Depart="KD01.0001.0001"  then
		datawhere=" and left(a.fnumber,2)='06' "
	  elseif Depart="KD01.0001.0002" then
		datawhere=" and left(a.fnumber,2)='03' "
	  elseif Depart="KD01.0001.0003" then
		datawhere=" and left(a.fnumber,2)='05' "
	  elseif Depart="KD01.0005.0004" then
		datawhere=" and left(a.fnumber,2)='02' "
	  elseif Depart="KD01.0001.0005" then
		datawhere=" and left(a.fnumber,2)='08' "
	  elseif Depart="KD01.0001.0006" then
		datawhere=" and left(a.fnumber,2)='07' "
	  elseif Depart="KD01.0001.0007" then
		datawhere=" and left(a.fnumber,2)='11' "
	  elseif Depart="KD01.0001.0008" then
		datawhere=" and left(a.fnumber,2)='12' "
	  elseif Depart="KD01.0001.0009" then
		datawhere=" and left(a.fnumber,2)='04' "
	  elseif Depart="KD01.0001.0010" then
		datawhere=" and left(a.fnumber,2)='10' "
	  elseif Depart="KD01.0001.0011" then
		datawhere=" and left(a.fnumber,2)='09' "
	  elseif Depart="KD01.0001.0012" then
		datawhere=" and left(a.fnumber,2)='01' "
	  elseif Depart="KD01.0001.0018" then
		datawhere=" and left(a.fnumber,5)='10.04' "
	  elseif Depart="KD01.0001.0019" then
		datawhere=" and left(a.fnumber,2)='23' "
	  end if
	Session("AllMessage20")=" and RejecteFlag!=1 and not exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber like OutPeron) "&datawhere
	case 21
	Session("AllMessage21")=" and CheckFlag=0 and Department='"&Depart&"'"
	case 22
	Session("AllMessage22")=" and CheckFlag=1 and CancelFlag=0"
	case 23
	Session("AllMessage23")=" and CheckFlag=2 and CancelFlag=0"
	case 24
	Session("AllMessage24")=" and CheckFlag=3 and datediff(d,PlanEndDate,getdate())>=0 and CancelFlag=0 "
	case 25
	Session("AllMessage25")=" and CheckFlag=0"
	case 26
	Session("AllMessage26")=" and CheckFlag=0 and (Department='"&Depart&"' or (Department='KD01.0001.0005' and '"&Depart&"'='KD01.0001.0012' )) "
	case 27
		if Depart="KD01.0001.0017" then
	Session("AllMessage27")=" and CheckFlag=1 and (OverType='分厂加班' or OverType='外协加班') "
		elseif Depart="KD01.0001.0003" then
	Session("AllMessage27")=" and CheckFlag=1 and OverType='收料加班' "
		end if
	case 28
	Session("AllMessage28")=""'" and CheckFlag=1 and (Department='KD01.0005.0004' or OverType='成品出货' or OverType='接送客人') "
	case 29
	Session("AllMessage29")=" and ((CheckFlag=1 and OverType<>'收料加班' and Department<>'KD01.0005.0004') or CheckFlag=2)"
	case 30
	Session("AllMessage30")=" and CheckFlag=3 "
	case 31
	Session("AllMessage31")=" and CheckFlag=0 "
	case 32
	Session("AllMessage32")=" and CheckFlag=1 and PCer='' and PCFlag=0 "
	case 33
	Session("AllMessage33")=" and CheckFlag=1 and PUer='' and PUFlag=0 "
	case 34
	Session("AllMessage34")=" and CheckFlag=1 and SAer='' and SAFlag=0 "
	case 35
	Session("AllMessage35")=" and CheckFlag=1 and ENer='' and ENFlag=0 "
	case 36
	Session("AllMessage36")=" and CheckFlag=1 and PCFlag=0 "
	case 37
	Session("AllMessage37")=" and CheckFlag=1 and PUFlag=0 "
	case 38
	Session("AllMessage38")=" and CheckFlag=1 and SAFlag=0 "
	case 39
	Session("AllMessage39")=" and CheckFlag=1 and ENFlag=0 "
	case 40
	Session("AllMessage40")=" and CheckFlag=1 and VPFlag=0 "
	case 41
	
  Session("AllMessage41")=" and (CheckFlag>1 and CheckFlag<99 and ( 1=2 "
		if Instr(session("AdminPurviewFLW"),"|201.7,")>0 then Session("AllMessage41")=Session("AllMessage41")&" or (shenqxm='工资调薪' and CheckFlag>2)"
		if Instr(session("AdminPurviewFLW"),"|201.8,")>0 then  Session("AllMessage41")=Session("AllMessage41")&" or shenqxm='话费补贴'"
		if Instr(session("AdminPurviewFLW"),"|201.9,")>0 then  Session("AllMessage41")=Session("AllMessage41")&" or shenqxm='住房补贴'"
		if Instr(session("AdminPurviewFLW"),"|201.10,")>0 then Session("AllMessage41")=Session("AllMessage41")&" or shenqxm='岗位补贴'"
		if Instr(session("AdminPurviewFLW"),"|201.11,")>0 then Session("AllMessage41")=Session("AllMessage41")&" or (shenqxm='其他补贴' and CheckFlag>2)"
		if Instr(session("AdminPurviewFLW"),"|201.12,")>0 then Session("AllMessage41")=Session("AllMessage41")&" or (shenqxm='职等调整' and ApplicEmployment<5) or (shenqxm='职等调整' and ApplicEmployment>4 and CheckFlag>2)"
		if Instr(session("AdminPurviewFLW"),"|201.13,")>0 then Session("AllMessage41")=Session("AllMessage41")&" or shenqxm='工龄恢复'"
		Session("AllMessage41")=Session("AllMessage41")&"))"
	case 42
	Session("AllMessage42")=" and CheckFlag=0 and Department='"&Depart&"' "
	case 43
	if Instr(PurviewFLW,"|215.3,")>0 then
		Session("AllMessage43")=" and CheckFlag=1 and Department='"&Depart&"'"
	end if
	if Instr(PurviewFLW,"|215.4,")>0 then
		Session("AllMessage43")=" and CheckFlag=1 "
		if Depart="KD01.0005.0001" then
		Session("AllMessage43")=" and CheckFlag=1 and (left(Department,9)='"&left(Depart,9)&"' or Department='KD01.0001.0009')"
		else
		Session("AllMessage43")=" and CheckFlag=1 and left(Department,9)='"&left(Depart,9)&"'"
		end if
	end if
	case 44
	Session("AllMessage44")=" and CheckFlag=0 "
	case 45
	Session("AllMessage45")=" and CheckFlag=0 and Department='"&Depart&"' and not exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber=RegisterID) "
	case 46
	Session("AllMessage46")=" and CheckFlag=0 and exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber=RegisterID) "
	case 47
	Session("AllMessage47")=" and CheckFlag=1 and CancelFlag=0 "
	if Depart="KD01.0005.0001" then Session("AllMessage47")=Session("AllMessage47")&" and (Department='KD01.0001.0009' or Department='"&Depart&"') "
	case 48
	Session("AllMessage48")=" and ((CheckFlag=1 and (SalaryType='1.行政月薪' or (TotalDay>6 and (SalaryType='2.分厂月薪' or SalaryType='3.计件')))) or (CheckFlag=0 and (SalaryType='1.行政月薪' or Grade>4))) and Department='"&Depart&"' and CancelFlag=0 "
	case 49
	Session("AllMessage49")=" and CheckFlag=2 and CancelFlag=0 and (exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber=RegisterID) or (TotalDay>6 and SalaryType='1.行政月薪') or ((SalaryType='1.行政月薪' or SalaryType='3.计件') and TotalDay>14)) "
	case 50
	Session("AllMessage50")=" and CancelFlag=0 and (CheckFlag=3 or (CheckFlag=2 and not exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber=RegisterID) and TotalDay<7) or (CheckFlag=1 and TotalDay<7 and (SalaryType='2.分厂月薪' or SalaryType='3.计件'))) "
	if Depart="KD01.0005.0001" then
		Session("AllMessage50")=Session("AllMessage50")&" and (Department='KD01.0001.0009' or left(Department,9)='"&left(Depart,9)&"')"
	elseif Depart<>"KD01.0001.0001" then
		Session("AllMessage50")=Session("AllMessage50")&" and Department='"&Depart&"' "
	end if
	case 51
	Session("AllMessage51")=" and RejecteFlag!=1 and len(OutPeron)>0 and exists (select  a.FNumber,a.Name from AIS20081217153921.dbo.HM_Employees a,AIS20081217153921.dbo.HM_EmployeesAddInfo b where b.EM_ID = a.EM_ID and a.Status=1 and hrms_userField_18=1 and a.FNumber like OutPeron) "
	case 52
	Session("AllMessage52")=" and CheckFlag=0 and EnFlag=0 "
	case 53
	Session("AllMessage53")=" and CheckFlag=0 and QcFlag=0 "
	case 54
	Session("AllMessage54")=" and CheckFlag=0 and PoFlag=0 "
	case 55
	Session("AllMessage55")=" and CheckFlag=0 and PoFlag=1 and QcFlag=1 and EnFlag=1 "
	case 56
	Session("AllMessage56")=" and CheckFlag=0 "
	case 57
	Session("AllMessage57")=" and CheckFlag=0 "
	case 58
	Session("AllMessage58")=" and CheckFlag=1 and CancelFlag=0 "
	case 59
	Session("AllMessage59")=" and CheckFlag=0 "
	case 60
	Session("AllMessage60")=" and CheckFlag=0 and Department='"&Depart&"' "
	case 61
	Session("AllMessage61")=" and CheckFlag=1 and CancelFlag=0 "
	case 62
	Session("AllMessage62")=" and CheckFlag=2 and CancelFlag=0 "
	case 63
	Session("AllMessage63")=" and CheckFlag=3 "
	case 64
	Session("AllMessage64")=" and CheckFlag=0 "
	case 65
	Session("AllMessage65")=" and CheckFlag=1 "
	case 66
	Session("AllMessage66")=" and CheckFlag=2 "
	case 67
	Session("AllMessage67")=" and CheckFlag=3 "
	case 68
	Session("AllMessage68")=" and CheckFlag=4 "
	case 69
	Session("AllMessage69")=" and a.确认=1 and b.员工代号='"&UserName&"' and b.是否已签合=0  and not exists (select 1 from [N-签合表单身] where a.编号=编号 and 序号<b.序号 and 是否已签合=0 ) "
	case 70
	Session("AllMessage70")=" and (AccountFlag=0 or ApprovFlag=0) and Invoice<>'' "
	case 71
	Session("AllMessage71")=" and CheckFlag=1 and CancelFlag=0 and (BillerID='"&UserName&"' or RegisterID='"&UserName&"') "
	case 72
	Session("AllMessage72")=" and CheckFlag=0 and Bumen='"&DepartName&"' "
	case 73
	Session("AllMessage73")=" and CheckFlag=1 and Bumen='"&DepartName&"' "
	case 74
	Session("AllMessage74")=" and a.CheckFlag=2 and SerialNum in (select Snum from manusys_JishiDetails2 where CheckerID='"&UserName&"' and CheckFlag=0) "
	case 75
	Session("AllMessage75")=" and a.CheckFlag=3 and a.Bumen='"&DepartName&"' "
 	end select
end if

sub Login()
  if session("UserName")="" or session("LoginSystem")<>"Succeed" then
     response.write "您还没有登录或登录已超时，请<a href='/u_Login.asp' target='_parent'><font color='red'>返回登录</font></a>!"
     response.end
	else
		 dim LoginIP,LoginTime,LoginSoft
		 LoginIP=Request.ServerVariables("Remote_Addr")
		 LoginSoft=Request.ServerVariables("Http_USER_AGENT")
		 LoginTime=now()
		 sql="select * from smmsys_Online where UserName='"&session("UserName")&"'"
		 set rs = server.createobject("adodb.recordset")
		 rs.open sql,connzxpt,1,1
		 if not rs.eof then
			sql="update smmsys_Online set o_ip='"&LoginIP&"',o_lasttime='"&LoginTime&"',LoginSoft='"&LoginSoft&"',AdminName='"&AdminName&"',o_state=1 where UserName='"&UserName&"'"
			connzxpt.Execute (sql)
		 else
			 sql = "insert into smmsys_Online (o_ip,UserName, o_lasttime,LoginSoft,AdminName,o_state) values ('"&LoginIP&"','" & UserName & "', '" & LoginTime & "','"&LoginSoft&"','" & AdminName & "',1)"
			 connzxpt.Execute (sql)
		 end if
  end if
end sub
 %>
