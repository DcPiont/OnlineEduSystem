<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ѧ�ڿγ̰�����Ϣɾ��</title>
<!--#include file="connNextCoursePlan.asp"-->
</head>
<%
sql="delete from nextCourseArrange where cno in ('"&trim(request("course_number"))&"') and TermNum in ('"&request("term_number")&"')"
conn.execute sql
conn.close
response.redirect("nextCoursePlan.asp")
%>
<body>
</body>
</html>
