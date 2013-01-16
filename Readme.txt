1.asp语言
2.utf-8编码
3.sqlserver2005数据库
4.OWC11分析图(regsvr32 OWC11.DLL)
5.jquery实现ajax
6.jquery.MultiFile插件多文件上传
7.jquery.form插件表单ajax提交
8.xheditor插件网页编辑器
9.iis启用上级目录
10.富文本在线打开RICHTX32.OCX(regsvr32 RICHTX32.OCX),控件：RICHTEXT.RichtextCtrl
11.金蝶k3压缩控件KDZIP.ZIP
12.版本控制使用svn,控件：SubWCRev.object
13.增加EXTJS4插件，grid应用
14.增加Flexigrid插件，考虑到EXTJS太庞大，考虑用JQUERY+FLEXIGRID实现grid效果
15.增加jquery.datepick插件，日期显示
16.增加jquery.anythingslider插件，图片显示
17.增加jquery.messager插件，弹窗显示预警信号
18.上传文件时报“Request 对象 错误 'ASP 0104 : 80004005' ”，
   对应代码Request.BinaryRead(Request.TotalBytes)时解决方法：
   2003增加了最大上传文件不能超过200K的限制,所以会报错,小于200K则是正常的. 找到文件c:\windows\system32\inetsrv\metabase.xml，用“记事本”打开该文件，用记事本中的“查找”功能搜索关键词ASPMaxRequestEntityAllowed”,搜索到结果如下图所示，ASPMaxRequestEntityAllowed="204800" 是win 2003用于限制最大上传文件大小的，默认是204800即200KB，可以根据您的具体情况，修改该值，1MB对应1024000，10MB对应10240000,依此类推，设置完该值，保存文件即可。注意修改的时候先停掉IIS服务,不然不允许改的.
