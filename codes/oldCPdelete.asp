<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ʷ�γ̰�����Ϣɾ������</title>
<!--#include file="connOldCoursePlan.asp"-->
</head>

<body>
<%
'response.Write(request("course_number"))
'response.Write(request("term_number"))
sql="delete from oldCourseArrange where cno in ('"&trim(request("course_number"))&"') and TermNum in ('"&request("term_number")&"')"
'response.Write(sql)
conn.execute sql
conn.close
response.redirect("oldCoursePlan.asp")
%>
</body>
</html>
