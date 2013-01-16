<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Server.scriptTimeout=300 %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>保存文件</title>
<link rel="stylesheet" href="Images/CssAdmin.css">
</head>
<!--#include file="UpFileClass.asp"-->
<body >
<table width=400 border=0 align="center" cellpadding="12" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td width=100% height=100% align="center" bgcolor="#EBF2F9" class=tablebody1 >
<%
dim Upload,File,FormName,SaveToPath,FileName,FileExt
dim RanNum
call UpFile()
'===========无组件上传(upload_0)====================
sub UpFile()
  set Upload=new UpFile_Class '建立上传对象
  Upload.GetData (1024*1024*15) '取得上传数据,此处即为15M

  if Upload.err > 0 then
    select case Upload.err
      case 1
        Response.Write "请先选择您要上传的文件，<a href=# onclick=history.go(-1)>返回</a>&nbsp;！"
      case 2
        Response.Write "文件大小超过了限制15M，<a href=# onclick=history.go(-1)>返回</a>&nbsp;！"
    end select
    exit sub
  else
    SaveToPath=Upload.form("SaveToPath")'"../Upload/" '文件保存目录,此目录必须为程序可读写
    if SaveToPath="" then
      SaveToPath="../Upload/"
    end if
    '在目录后加(/)
    if right(SaveToPath,1)<>"/" then 
      SaveToPath=SaveToPath&"/"
    end if 
    for each FormName in Upload.file '列出所有上传了的文件
      set file=Upload.file(FormName) '生成一个文件对象
      if file.Filesize<100 then
        response.write "请先选择您要上传的文件&nbsp;！"
        response.end
      end if

      FileExt=lcase(File.FileExt)
      if CheckFileExt(FileEXT)=false then
        response.write File.FilePath&file.FileName&"文件格式不允许上传&nbsp;！"
        response.end
      end if

      randomize timer
      RanNum=int(9000*rnd)+1000
      Filename=SaveToPath&year(now)&"."&month(now)&"."&day(now)&"_"&hour(now)&"."&minute(now)&"."&Second(now)&"_"&RanNum&"."&fileExt
      if file.FileSize>0 then '如果 FileSize > 0 说明有文件数据
        Result=file.SaveToFile(Server.mappath(FileName)) '保存文件
        if Result="ok" then
		
		  response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
          response.write "<tr>"
          response.write "<td width='60' height='30'>上传成功：</td>"
          response.write "<td nowrap><font color='#ff0000'>"&File.FilePath&file.FileName&"</font></td>"
          response.write "</tr>"
          response.write "<tr>"
          response.write "<td nowrap height='30'>保存路径：</td>"
          response.write "<td nowrap><input type='text' size='56' class='textfield' value='"&right(FileName,len(FileName))&"' readonly></td>"
          response.write "</tr>"
          response.write "<tr>"
          response.write "<td nowrap height='30'>文件大小：</td>"
          response.write "<td nowrap><input type='text' size='56' class='textfield' value='"&GainFileSize(file.Filesize)&"' readonly></td>"
          response.write "</tr>"		  
          response.write "<tr>"
          response.write "<td height='36' colspan='2' valign='bottom' align='center'><input name='CopyPath' type='button' class='button' value='拷贝文件路径'  onclick=""CopyPath2Pic('"&right(FileName,len(FileName))&"','"&GainFileSize(file.Filesize)&"')""></td>"
          response.write "</tr>"
          response.write "</table>"
        else
          response.write File.FilePath&file.FileName&"上传失败&nbsp;！"&Result&"<br>"
        end if
      end if
      set file=nothing
    next
    set Upload=nothing
  end if
end sub

'判断文件类型是否合格
Private Function CheckFileExt (FileEXT)
  dim ForumUpload
  ForumUpload="exe,gif,jpg,jpeg,rar,zip,doc,xls,bmp"
  ForumUpload=split(ForumUpload,",")
  for i=0 to ubound(ForumUpload)
    if lcase(FileEXT)=lcase(trim(ForumUpload(i))) then
      CheckFileExt=true
      exit Function
    else
      CheckFileExt=false
    end if
  next
End Function

Private Function GainFileSize (SizeByte)
  if SizeByte < 1024*1024 then
    GainFileSize=round(SizeByte/1024,2) & "&nbsp;KB"
  else  
    GainFileSize=round(SizeByte/1024/1024,2) & "&nbsp;MB"
  end if
End Function

%>
    </td>
  </tr>
</table>
</body>
</html>