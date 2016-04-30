<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>课程基本信息删除</title>
<!--连接数据库-->
<!--#include file="connCourseInfo.asp"-->
</head>
<body>
<%
'删除数据
sql= "delete from CourseInfo where cno='"&request.QueryString&"' "
conn.execute sql
conn.close
response.redirect("CImaintenance.asp")
%>
</body>
</html>
