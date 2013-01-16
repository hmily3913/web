<%
Sub SelPlay(strUrl,strWidth,StrHeight)
Dim Exts,isExt
If strUrl <> "" Then
    isExt = LCase(Mid(strUrl,InStrRev(strUrl, ".")+1))
Else
    isExt = ""
End If
Exts = "avi,wmv,asf,mov,rm,ra,ram"
If Instr(Exts,isExt)=0 Then
Response.write "非法视频文件"
Else
Select Case isExt
   Case "avi","wmv","asf","mov"
    Response.write "<EMBED id=MediaPlayer src="&strUrl&" width="&strWidth&" height="&strHeight&" loop=""false"" autostart=""true""></EMBED>"
   Case "mov","rm","ra","ram"
    Response.Write "<OBJECT height="&strHeight&" width="&strWidth&" classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA>"
    Response.Write "<PARAM NAME=""_ExtentX"" VALUE=""12700"">"
    Response.Write "<PARAM NAME=""_ExtentY"" VALUE=""9525"">"
    Response.Write "<PARAM NAME=""AUTOSTART"" VALUE=""-1"">"
    Response.Write "<PARAM NAME=""SHUFFLE"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""PREFETCH"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""NOLABELS"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""SRC"" VALUE="""&strUrl&""">"
    Response.Write "<PARAM NAME=""CONTROLS"" VALUE=""ImageWindow"">"
    Response.Write "<PARAM NAME=""CONSOLE"" VALUE=""Clip"">"
    Response.Write "<PARAM NAME=""LOOP"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""NUMLOOP"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""CENTER"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""MAINTAINASPECT"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""BACKGROUNDCOLOR"" VALUE=""#000000"">"
    Response.Write "</OBJECT>"
    Response.Write "<BR>"
    Response.Write "<OBJECT height=32 width="&strWidth&" classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA>"
    Response.Write "<PARAM NAME=""_ExtentX"" VALUE=""12700"">"
    Response.Write "<PARAM NAME=""_ExtentY"" VALUE=""847"">"
    Response.Write "<PARAM NAME=""AUTOSTART"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""SHUFFLE"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""PREFETCH"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""NOLABELS"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""CONTROLS"" VALUE=""ControlPanel,StatusBar"">"
    Response.Write "<PARAM NAME=""CONSOLE"" VALUE=""Clip"">"
    Response.Write "<PARAM NAME=""LOOP"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""NUMLOOP"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""CENTER"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""MAINTAINASPECT"" VALUE=""0"">"
    Response.Write "<PARAM NAME=""BACKGROUNDCOLOR"" VALUE=""#000000"">"
    Response.Write "</OBJECT>"
End Select
End If
End Sub
'dim ImgFileName '图片保存名
'（类型，图表名称，存储文件名，共几组,标签，数据，宽，高）
Sub OwcToGif(ChartType,ImgTitle,ImgName,MaxImum,categories,values,imgwidth,imgheight) 
 '创建owc对象
 dim chart,chartShow,chartConst,fnt,strChartAbsPath,strChartRelPath,strChart,strFileName
 set chart=Server.CreateObject("OWC11.ChartSpace")
 chart.clear
 set chartShow=chart.Charts.Add
 set chartConst = chart.Constants'设定图表类型，图表类型有很多种，详见OWC帮助文件

 Select Case ChartType
  Case 1 '柱形图
  chartShow.Type = chartConst.chChartTypeColumnStacked3d
  Case 2 '饼形图
  chartShow.Type = chartConst.chChartTypePieExploded3D
  Case 3 '雷达图
  chartShow.Type = chartConst.chChartTypeRadarLineFilled
  Case Else
  chartShow.Type = chartConst.chChartTypeColumnStacked
 End Select
chartShow.HasLegend=false '图例
 '(显示图例)'以下为图表标题设定
  chartShow.Inclination=0 '上下倾斜度
  chartShow.AmbientLightIntensity=0.9 '环绕光
  chartShow.Rotation=0 '左右倾斜度
  chartShow.DirectionalLightRotation=360'光源的旋转角度
  chartShow.ExtrudeAngle=90
  chartShow.HasTitle=true '指定图表或坐标轴具有标题
  chartShow.Title.Caption=ImgTitle  
  set fnt=chartShow.title.font
  fnt.size=13
  fnt.bold=true
  fnt.color="blue" 
'  chartShow.Axes(0).HasTitle=True
'  chartShow.Axes(0).Title.Caption="关数"
'  chartShow.Axes(1).HasTitle=True
'  chartShow.Axes(1).Title.Caption="成绩"
'  chartShow.Axes(1).Title.font.size=10
  chartShow.SetData chartConst.chDimCategories, chartConst.chDataLiteral, categories
  chartShow.SeriesCollection(0).SetData chartConst.chDimValues, chartConst.chDataLiteral, values
'  chartShow.SeriesCollection(0).Interior.Color = "Red"
  'chartShow.Axes(chartConst.chAxisPositionLeft).NumberFormat = "0.00%"
  'chartShow.Axes(chartConst.chAxisPositionLeft).MajorUnit = 25
  'chartShow.Axes(chartConst.chAxisPositionLeft).GroupingUnit=1
'  chartShow.Axes(chartConst.chAxisPositionValue).Scaling.Maximum=MaxImum
  chartShow.Scalings(chartConst.chDimValues).Maximum = MaxImum 
  chartShow.Scalings(chartConst.chDimValues).Minimum = 0
  With chartShow.SeriesCollection(0).DataLabelsCollection.Add '添加图例的数据标记
  .Font.Size = 9  
  if ChartType=2 then
    .HasCategoryName = True
	.HasValue = false   
    .HasPercentage = True
  else
    .HasValue = true              
  end if
'  .HasSeriesName = True
     ' .Column.Color = RGB(255, 255, 0)
     End With
     strChartAbsPath=Server.MapPath("../temp/")
     '将图保存到当前路径下的temp文件夹。
     strChartRelPath = "temp"
	strFileName = ImgName&".gif"
	chart.ExportPicture strChartAbsPath&"/"&strFileName,"gif",imgwidth, imgheight 
'	ImgFileName = strChartRelPath &"/" & strFileName
Set chartConst = nothing
Set chart = nothing
end Sub
function StrLen(Str)
  if Str="" or isnull(Str) then 
    StrLen=0
    exit function 
  else
    dim regex
    set regex=new regexp
    regEx.Pattern ="[^\x00-\xff]"
    regex.Global =true
    Str=regEx.replace(Str,"^^")
    set regex=nothing
    StrLen=len(Str)
  end if
end function

function StrLeft(Str,StrLen)
  dim L,T,I,C
  if Str="" then
    StrLeft=""
    exit function
  end if
  Str=Replace(Replace(Replace(Replace(Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
  L=Len(Str)
  T=0
  for i=1 to L
    C=Abs(AscW(Mid(Str,i,1)))
    if C>255 then
      T=T+2
    else
      T=T+1
    end if
    if T>=StrLen then
      StrLeft=Left(Str,i) & "..."
      exit for
    else
      StrLeft=Str
    end if
  next
  StrLeft=Replace(Replace(Replace(replace(StrLeft," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

function StrReplace(Str)'表单存入替换字符
  if Str="" or isnull(Str) then 
    StrReplace=""
    exit function 
  else
    StrReplace=replace(str," ","&nbsp;") '"&nbsp;"
    StrReplace=replace(StrReplace,chr(13),"&lt;br&gt;")'"<br>"
    StrReplace=replace(StrReplace,"<","&lt;")' "&lt;"
    StrReplace=replace(StrReplace,">","&gt;")' "&gt;"
  end if
end function

function ReStrReplace(Str)'写入表单替换字符
  if Str="" or isnull(Str) then 
    ReStrReplace=""
    exit function 
  else
    ReStrReplace=replace(Str,"&nbsp;"," ") '"&nbsp;"
    ReStrReplace=replace(ReStrReplace,"<br>",chr(13))'"<br>"
    ReStrReplace=replace(ReStrReplace,"&lt;br&gt;",chr(13))'"<br>"
    ReStrReplace=replace(ReStrReplace,"&lt;","<")' "&lt;"
    ReStrReplace=replace(ReStrReplace,"&gt;",">")' "&gt;"
  end if
end function

function HtmlStrReplace(Str)'写入Html网页替换字符
  if Str="" or isnull(Str) then 
    HtmlStrReplace=""
    exit function 
  else
    HtmlStrReplace=replace(Str,"&lt;br&gt;","<br>")'"<br>"
  end if
end function

function ViewNoRight(GroupID,Exclusive)
  dim rs,sql,GroupLevel
  set rs = server.createobject("adodb.recordset")
  sql="select GroupLevel from Ameav_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupLevel=rs("GroupLevel")
  rs.close
  set rs=nothing
  ViewNoRight=true
  if session("GroupLevel")="" then session("GroupLevel")=0
  select case Exclusive
    case ">="
      if not session("GroupLevel") >= GroupLevel then
	    ViewNoRight=false
	  end if
    case "="
      if not session("GroupLevel") = GroupLevel then
	    ViewNoRight=false
      end if
  end select
end function

Function GetUrl()
  GetUrl="http://"&Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL")
  If Request.ServerVariables("QUERY_STRING")<>"" Then GetURL=GetUrl&"?"& Request.ServerVariables("QUERY_STRING")
End Function

function HtmlSmallPic(GroupID,PicPath,Exclusive)
  dim rs,sql,GroupLevel
  set rs = server.createobject("adodb.recordset")
  sql="select GroupLevel from Ameav_MemGroup where GroupID='"&GroupID&"'"
  rs.open sql,conn,1,1
  GroupLevel=rs("GroupLevel")
  rs.close
  set rs=nothing
  HtmlSmallPic=PicPath
  if session("GroupLevel")="" then session("GroupLevel")=0
  select case Exclusive
    case ">="
      if not session("GroupLevel") >= GroupLevel then HtmlSmallPic="../Images/NoRight.jpg"
    case "="
      if not session("GroupLevel") = GroupLevel then HtmlSmallPic="../Images/NoRight.jpg"
  end select
  if HtmlSmallPic="" or isnull(HtmlSmallPic) then HtmlSmallPic="../Images/NoPicture.jpg"
end function

function IsValidMemName(memname)
  dim i, c
  IsValidMemName = true
  if not (3<=len(memname) and len(memname)<=16) then
    IsValidMemName = false
    exit function
  end if  
  for i = 1 to Len(memname)
    c = Mid(memname, i, 1)
    if InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-", c) <= 0 and not IsNumeric(c) then
      IsValidMemName = false
      exit function
    end if
  next
end function

function IsValidEmail(email)
  dim names, name, i, c
  IsValidEmail = true
  names = Split(email, "@")
  if UBound(names) <> 1 then
    IsValidEmail = false
    exit function
  end if
  for each name in names
	if Len(name) <= 0 then
	  IsValidEmail = false
      exit function
    end if
    for i = 1 to Len(name)
      c = Mid(name, i, 1)
      if InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-.", c) <= 0 and not IsNumeric(c) then
        IsValidEmail = false
        exit function
      end if
	next
	if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
    end if
  next
  if InStr(names(1), ".") <= 0 then
    IsValidEmail = false
    exit function
  end if
  i = Len(names(1)) - InStrRev(names(1), ".")
  if i <> 2 and i <> 3 then
    IsValidEmail = false
    exit function
  end if
  if InStr(email, "..") > 0 then
    IsValidEmail = false
  end if
end function

'================================================
'函数名：FormatDate
'作　用：格式化日期
'参　数：DateAndTime            (原日期和时间)
'       Format                 (新日期格式)
'返回值：格式化后的日期
'================================================
Function FormatDate(DateAndTime, Format)
  On Error Resume Next
  Dim yy,y, m, d, h, mi, s, strDateTime
  FormatDate = DateAndTime
  If Not IsNumeric(Format) Then Exit Function
  If Not IsDate(DateAndTime) Then Exit Function
  yy = CStr(Year(DateAndTime))
  y = Mid(CStr(Year(DateAndTime)),3)
  m = CStr(Month(DateAndTime))
  If Len(m) = 1 Then m = "0" & m
  d = CStr(Day(DateAndTime))
  If Len(d) = 1 Then d = "0" & d
  h = CStr(Hour(DateAndTime))
  If Len(h) = 1 Then h = "0" & h
  mi = CStr(Minute(DateAndTime))
  If Len(mi) = 1 Then mi = "0" & mi
  s = CStr(Second(DateAndTime))
  If Len(s) = 1 Then s = "0" & s
   
  Select Case Format
  Case "1"
    strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
  Case "2"
    strDateTime = yy & m & d & h & mi & s
    '返回12位 直到秒 的时间字符串
  Case "3"
    strDateTime = yy & m & d & h & mi    
    '返回12位 直到分 的时间字符串
  Case "4"
    strDateTime = yy & "年 " & m & "月 " & d & "日 "
  Case "5"
    strDateTime = m & "-" & d
  Case "6"
    strDateTime = m & "/" & d
  Case "7"
    strDateTime = m & "月 " & d & "日 "
  Case "8"
    strDateTime = y & "年 " & m & "月 "
  Case "9"
    strDateTime = y & "-" & m
  Case "10"
    strDateTime = y & "/" & m
  Case "11"
    strDateTime = y & "-" & m & "-" & d
  Case "12"
    strDateTime = y & "/" & m & "/" & d
  Case "13"
    strDateTime = yy & "." & m & "." & d
  Case Else
    strDateTime = DateAndTime
  End Select
  FormatDate = strDateTime
End Function

function WriteMsg(Message)
  response.write "<table width='400' border='0' align='center' cellpadding='1' cellspacing='1' bgcolor='#FF3300'>" &_
                 "  <tr>" &_
                 "    <td bgcolor='#FFFFFF'>" &_
                 "    <table width='100%' border='0' cellpadding='0' cellspacing='0' bgcolor='#FF3300'><tr>" &_
                 "      <td align='center' style='font-family:Arial;font-size:16px;color:#FFFFFF;font-weight:bold'>MESSAGE</td>" &_
                 "    </tr></table>" &_
                 "    </td>" &_
                 "  </tr>" &_
                 "  <tr>" &_
                 "    <td bgcolor='#FFFFFF' >" &_
                 "    <table width='100%' border='0' cellspacing='0' cellpadding='4'>" &_
                 "      <tr>" &_
                 "        <td bgcolor='#FFFFFF' style='font-family:Arial;font-size:12px;line-height:18px;color:#333333;'>" &_
				 Message &_
                 "        </td>" &_
                 "      </tr>" &_
                 "    </table>" &_
                 "	  </td>" &_
                 "	</tr>" &_
                 "</table>" &_
                 "<div align='center'>" &_
                 "<br>" &_
                 "<a href='javascript:history.back()'><img src='../Images/Arrow_05.gif' width='22' height='22' border='0' /></a>" &_
                 "</div>"
end function

'获取当前时间是第几周函数：
'程序代码
Function GetWeekNo(InputDate)
dim pytY,pytNewYear,pytNewYearWeek,pytAllDay,pytBanWeek,NumWeek,tempx
NumWeek = 0
pytY = Year(InputDate)
pytNewYear=pytY &"-1-1"
pytNewYearWeek = Weekday(pytNewYear)
pytAllDay = DateDiff("d",pytNewYear,InputDate)
pytBanWeek = 8-pytNewYearWeek
if pytBanWeek<7 Then
NumWeek = 1
pytAllDay = pytAllDay - pytBanWeek
end if
tempx = pytAllDay/7
tempx = -Int(-tempx)
NumWeek = NumWeek+tempx
GetWeekNo = NumWeek
end Function 

'指定周数的日期范围函数
'程序代码
'函数名 getDateRange 
'函数 index -数值型:指定周数 
Function getDateRange(byVal Index,byVal years) 
Dim CurDate, retDate, Days, retVal 
if years="" then
CurDate = CDate(Year(Date()) & "-1-1") 
else
CurDate = CDate(years & "-1-1") 
end if
if (WeekDay(CurDate)<>1) Then Index =Index-1 
Days=Index * 7 
retDate=DateAdd("d", (Days - 1), CurDate) 
if (retDate < CurDate) Then retDate=CurDate 
retDate=DateAdd("d", 1-Weekday(retDate), retDate) 
if (retDate< CurDate) then 
retVal=CurDate & "###" & DateAdd("d", 7, retDate) 
else 
retVal=DateAdd("d", 1, retDate) & "###" & DateAdd("d", 7, retDate) 
end if 
getDateRange = retVal 
End Function
'指定月数的日期范围函数
Function getDateRangebyMonth(Index) 
dim stryear,strmonth,edaynum
if Instr(Index,"#")>0 then
stryear=Split(Index,"#")(1)
strmonth=Split(Index,"#")(0)
else
strmonth=Index
stryear=Year(now())
end if
select case strmonth
	case 2
	  if ((stryear mod 4=0) and (stryear mod 100>0)) or (stryear mod 400=0) then
	    edaynum=29
	  else
	    edaynum=28
	  end if
    case 4
	  edaynum=30
    case 6
	  edaynum=30
    case 9
	  edaynum=30
    case 11
	  edaynum=30
	case else
	  edaynum=31
end select
getDateRangebyMonth=stryear&"-"&strmonth&"-1###"&stryear&"-"&strmonth&"-"&edaynum
End Function

Function DataToRtf(FieldOne,TableOne,Idwhere)
dim rsRich,SQLRick
set rsRich=server.createobject("adodb.recordset")
SQLRick = "select "&FieldOne&" from "&TableOne&" where "&Idwhere
rsRich.open SQLRick,connk3,1,1
If Not rsRich.Eof Then
    Dim oZip        
    Dim smFile      
    Dim lOffset     
    Dim lFileSize   
    Dim bytChunk()

	dim objFSO,strChartAbsPath,sTempZipFileName,DestTempFileName,RanNum
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strChartAbsPath=objFSO.GetSpecialFolder(2)
	randomize timer
	RanNum=int(9000*rnd)+1000
	
'	strChartAbsPath=Server.MapPath("temp/")

	sTempZipFileName=strChartAbsPath&"\BigText"&RanNum
	DestTempFileName=strChartAbsPath&"\BigTextrtf"&RanNum
	set objFSO=Nothing
	
    Set smFile = CreateObject("ADODB.Stream")
    smFile.Type = 1              'adTypeBinary=1
    smFile.Open
    
    lOffset = 0
    lFileSize = rsRich(FieldOne).ActualSize
	smFile.Write rsRich(FieldOne).GetChunk(lFileSize)
    smFile.Position = 0
    smFile.SaveToFile sTempZipFileName, 2
    Set smFile = Nothing

    Set oZip = CreateObject("KDZIP.ZIP")
    oZip.DeCompress sTempZipFileName, DestTempFileName
    Set oZip = Nothing
	dim oRICHTX
	Set oRICHTX = CreateObject("RICHTEXT.RichtextCtrl")
	oRICHTX.LoadFile DestTempFileName

	response.write RichToHTML(oRICHTX)

	Set oRICHTX = Nothing
End If
rsRich.Close
Set rsRich = Nothing
End Function
Function RichToHTML(rtbRichTextBox)
'**********************************************************
'*            Rich To HTML by Joseph Huntley              *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'**********************************************************
'*   You may use this code freely as long as credit is    *
'* given to the author, and the header remains intact.    *
'**********************************************************

'--------------------- The Arguments -----------------------
'rtbRichTextBox     - The rich textbox control to convert.
'lngStartPosition   - The character position to start from.
'lngEndPosition     - The character position to end at.
'-----------------------------------------------------------
'Returns:     The rich text converted to HTML.

'Description: Converts rich text to HTML.
dim lngStartPosition , lngEndPosition
Dim blnBold , blnUnderline, blnStrikeThru
Dim blnItalic, strLastFont , lngLastFontColor
Dim strHTML , lngColor , lngRed , lngGreen 
Dim lngBlue , lngCurText , strHex , intLastAlignment 

Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2

'check for lngStartPosition ad lngEndPosition

lngStartPosition = 0
lngEndPosition = 500'Len(rtbRichTextBox.Text)

lngLastFontColor = -1 'no color

   For lngCurText = lngStartPosition To lngEndPosition
       rtbRichTextBox.SelStart = lngCurText
       rtbRichTextBox.SelLength = 1
          If intLastAlignment <> rtbRichTextBox.SelAlignment Then
             intLastAlignment = rtbRichTextBox.SelAlignment
              
                Select Case rtbRichTextBox.SelAlignment
                   Case AlignLeft: strHTML = strHTML & "<p align=left>"
                   Case AlignRight: strHTML = strHTML & "<p align=right>"
                   Case AlignCenter: strHTML = strHTML & "<p align=center>"
                End Select
                
          End If
   
          If blnBold <> rtbRichTextBox.SelBold Then
               If rtbRichTextBox.SelBold = True Then
                 strHTML = strHTML & "<b>"
               Else
                 strHTML = strHTML & "</b>"
               End If
             blnBold = rtbRichTextBox.SelBold
          End If

          If blnUnderline <> rtbRichTextBox.SelUnderline Then
               If rtbRichTextBox.SelUnderline = True Then
                 strHTML = strHTML & "<u>"
               Else
                 strHTML = strHTML & "</u>"
               End If
             blnUnderline = rtbRichTextBox.SelUnderline
          End If
   

          If blnItalic <> rtbRichTextBox.SelItalic Then
               If rtbRichTextBox.SelItalic = True Then
                 strHTML = strHTML & "<i>"
               Else
                 strHTML = strHTML & "</i>"
               End If
             blnItalic = rtbRichTextBox.SelItalic
          End If


          If blnStrikeThru <> rtbRichTextBox.SelStrikeThru Then
               If rtbRichTextBox.SelStrikeThru = True Then
                 strHTML = strHTML & "<s>"
               Else
                 strHTML = strHTML & "</s>"
               End If
             blnStrikeThru = rtbRichTextBox.SelStrikeThru
          End If

         If strLastFont <> rtbRichTextBox.SelFontName Then
            strLastFont = rtbRichTextBox.SelFontName
            strHTML = strHTML + "<font face=""" & strLastFont & """>"
         End If

         If lngLastFontColor <> rtbRichTextBox.SelColor Then
            lngLastFontColor = rtbRichTextBox.SelColor
            
            ''Get hexidecimal value of color
            strHex = Hex(rtbRichTextBox.SelColor)
            strHex = String(6 - Len(strHex), "0") & strHex

            strHex = Right(strHex, 2) & Mid(strHex, 3, 2) & Left(strHex, 2)
            
            strHTML = strHTML + "<font color=#" & strHex & ">"
        End If
     strHTML = strHTML + rtbRichTextBox.SelText
   Next

RichToHTML = strHTML

End Function


Sub SendMail(ToEml,ToSubject,ToSerialNum,ToText,strAttachmentName)
  Const cdoSendUsingMethod="http://schemas.microsoft.com/cdo/configuration/sendusing" 
  Const cdoSendUsingPort=2 
  Const cdoSMTPServer="http://schemas.microsoft.com/cdo/configuration/smtpserver" 
  Const cdoSMTPServerPort="http://schemas.microsoft.com/cdo/configuration/smtpserverport" 
  Const cdoSMTPConnectionTimeout="http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout" 
  Const cdoSMTPAuthenticate="http://schemas.microsoft.com/cdo/configuration/smtpauthenticate" 
  Const cdoBasic=1 
  Const cdoSendUserName="http://schemas.microsoft.com/cdo/configuration/sendusername" 
  Const cdoSendPassword="http://schemas.microsoft.com/cdo/configuration/sendpassword" 
  
  Dim objConfig 
  Dim objMessage  
  Dim Fields
  
  Set objConfig = Server.CreateObject("CDO.Configuration") 
  Set Fields = objConfig.Fields 
  
  With Fields 
  .Item(cdoSendUsingMethod) = cdoSendUsingPort 
  .Item(cdoSMTPServer) = "122.228.158.226"   
  .Item(cdoSMTPServerPort) = 25                  
  .Item(cdoSMTPConnectionTimeout) = 300     
  .Item(cdoSMTPAuthenticate) = 1 
  .Item(cdoSendUserName) = "admin"  
  .Item(cdoSendPassword) = "123456789"         
  .Update 
  End With 
  
  Set objMessage = Server.CreateObject("CDO.Message") 
  Set objMessage.Configuration = objConfig 
  
  With objMessage
  .BodyPart.Charset = "utf-8"                     
  .To = ToEml                                  
  .From = "administrator@loverdoor.cn"                       
  .Subject = ToSubject                
  .htmlBody = "单号为:"&ToSerialNum&",内容为："&ToText
	if strAttachmentName <> "" then 
		.AddAttachment strAttachmentName
	end if 
  .Send 
  End With 
  
  Set Fields = Nothing 
  Set objMessage = Nothing 
  Set objConfig = Nothing
End Sub

Function DataTypeName(TypeID)
    Dim DataType(), z
    ReDim DataType(2, 36)
    DataType(0, 0) = "adTinyInt"
    DataType(1, 0) = 16
    DataType(2, 0) = "数字"
    DataType(0, 1) = "adSmallInt"
    DataType(1, 1) = 2
    DataType(2, 1) = "数字"
    DataType(0, 2) = "adInteger"
    DataType(1, 2) = 3
    DataType(2, 2) = "数字"
    DataType(0, 3) = "adBigInt"
    DataType(1, 3) = 20
    DataType(2, 3) = "数字"
    DataType(0, 4) = "adUnsignedTinyInt"
    DataType(1, 4) = 17
    DataType(2, 4) = "数字"
    DataType(0, 5) = "adUnsignedSmallInt"
    DataType(1, 5) = 18
    DataType(2, 5) = "数字"
    DataType(0, 6) = "adUnsignedInt"
    DataType(1, 6) = 19
    DataType(2, 6) = "数字"
    DataType(0, 7) = "adUnsignedBigInt"
    DataType(1, 7) = 21
    DataType(2, 7) = "数字"
    DataType(0, 8) = "adSingle"
    DataType(1, 8) = 4
    DataType(2, 8) = "数字"
    DataType(0, 9) = "adDouble"
    DataType(1, 9) = 5
    DataType(2, 9) = "数字"
    DataType(0, 10) = "adCurrency"
    DataType(1, 10) = 6
    DataType(2, 10) = "数字"
    DataType(0, 11) = "adDecimal"
    DataType(1, 11) = 14
    DataType(2, 11) = "数字"
    DataType(0, 12) = "adNumeric"
    DataType(1, 12) = 131
    DataType(2, 12) = "数字"
    DataType(0, 13) = "adBoolean"
    DataType(1, 13) = 11
    DataType(2, 13) = "Bool"
    DataType(0, 14) = "adError"
    DataType(1, 14) = 10
    DataType(2, 14) = "adError"
    DataType(0, 15) = "adGUID"
    DataType(1, 15) = 72
    DataType(2, 15) = "adGUID"
    DataType(0, 16) = "adDate"
    DataType(1, 16) = 7
    DataType(2, 16) = "日期"
    DataType(0, 17) = "adDBDate"
    DataType(1, 17) = 133
    DataType(2, 17) = "日期"
    DataType(0, 18) = "adDBTime"
    DataType(1, 18) = 134
    DataType(2, 18) = "日期"
    DataType(0, 19) = "adDBTimeStamp"
    DataType(1, 19) = 135
    DataType(2, 19) = "日期"
    DataType(0, 20) = "adDBTimeStamp"
    DataType(1, 20) = 7
    DataType(2, 20) = "日期"
    DataType(0, 21) = "adBSTR"
    DataType(1, 21) = 8
    DataType(2, 21) = "文本"
    DataType(0, 22) = "adBSTR"
    DataType(1, 22) = 130
    DataType(2, 22) = "文本"
    DataType(0, 23) = "adChar"
    DataType(1, 23) = 129
    DataType(2, 23) = "文本"
    DataType(0, 24) = "adChar"
    DataType(1, 24) = 200
    DataType(2, 24) = "文本"
    DataType(0, 25) = "adVarChar"
    DataType(1, 25) = 200
    DataType(2, 25) = "文本"
    DataType(0, 26) = "adLongVarChar"
    DataType(1, 26) = 200
    DataType(2, 26) = "文本"
    DataType(0, 27) = "adLongVarChar"
    DataType(1, 27) = 201
    DataType(2, 27) = "文本"
    DataType(0, 28) = "adWChar"
    DataType(1, 28) = 130
    DataType(2, 28) = "文本"
    DataType(0, 29) = "adVarWChar"
    DataType(1, 29) = 130
    DataType(2, 29) = "文本"
    DataType(0, 30) = "adVarWChar"
    DataType(1, 30) = 202
    DataType(2, 30) = "文本"
    DataType(0, 31) = "adLongVarWChar"
    DataType(1, 31) = 203
    DataType(2, 31) = "文本"
    DataType(0, 32) = "adLongVarWChar"
    DataType(1, 32) = 130
    DataType(2, 32) = "文本"
    DataType(0, 33) = "adBinary"
    DataType(1, 33) = 128
    DataType(2, 33) = "adBinary"
    DataType(0, 34) = "adVarBinary"
    DataType(1, 34) = 204
    DataType(2, 34) = "adBinary"
    DataType(0, 35) = "adLongVarBinary"
    DataType(1, 35) = 204
    DataType(2, 35) = "adBinary"
    DataType(0, 36) = "adLongVarBinary"
    DataType(1, 36) = 205
    DataType(2, 36) = "adBinary"
    For z = 0 To 36
        If DataType(1, z) = TypeID Then
            DataTypeName = DataType(2, z)
            Exit Function
        End If
    Next
End Function

Function JsonStr(valueStr)
if isnull(valueStr) then
JsonStr=""
else
JsonStr=replace(replace(replace(replace(replace(replace(replace(valueStr,chr(92),"\\"),chr(10),"\n"),chr(13),"\r"),chr(34),"\"""),chr(9),"\t"),chr(39),"\"""),chr(47),"\/")
end if
End Function
Function parseJSON(str)
	Dim scriptCtrl
	If Not IsObject(scriptCtrl) Then
		Set scriptCtrl = Server.CreateObject("MSScriptControl.ScriptControl")
		scriptCtrl.Language = "JScript"
		scriptCtrl.AddCode "function ActiveXObject() {}" ' 覆盖 ActiveXObject
		scriptCtrl.AddCode "function GetObject() {}" ' 覆盖 ActiveXObject
		scriptCtrl.AddCode "Array.prototype.get = function(x) { return this[x]; }; var result = null;"
	End If
  On Error Resume Next
	scriptCtrl.ExecuteStatement "result = " & str & ";"
	Set parseJSON = scriptCtrl.CodeObject.result
  If Err Then
	Err.Clear
	Set parseJSON = Nothing
  End If
	If IsObject(scriptCtrl) Then Set scriptCtrl = Nothing
End Function
'Dim json
'json = "{a:""aaa"", b:{ name:""bb"", value:""text"" }, c:[""item0"", ""item1"", ""item2""]}"
'Set obj = parseJSON(json)
'Response.Write obj.a & "<br />"
'Response.Write obj.b.name & "<br />"
'Response.Write obj.c.length & "<br />"
'Response.Write obj.c.get(0) & "<br />"
'Set obj = Nothing

function getBillNo(TbName,NumBit,BillDt)
	dim Zb,Zc,Zd
	Zb=year(now())
	Zc=month(now())
	if Zc<10 then
		Zc="0"+cstr(Zc)
	end if
	Zd=day(now())
	if Zd<10 then
		Zd="0"+cstr(Zd)
	end if
	dim ZBillNo,rsBillno
	ZBillNo=right(cstr(Zb),2)+cstr(Zc)+cstr(Zd)
  set rsBillno = server.createobject("adodb.recordset")
  sql="select top 1 SerialNum from "&TbName&" where SerialNum like '"&ZBillNo&"%' order by SerialNum desc"
  rsBillno.open sql,connzxpt,1,1
	if rsBillno.eof then
		getBillNo=ZBillNo+RIGHT("000000000000",NumBit-1)+"1"
	else
		getBillNo=rsBillno("SerialNum")+1
	end if
	rsBillno.close
	set rsBillno=nothing 
end function
%>
