<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
<!--连接数据库-->
<!--#include file="connCoursePlan.asp"-->
</head>
<%
'添加数据
course_number=request.Form("course_number")
term_number=request.Form("term_number")
major_number=request.Form("major_number")
max_student=request.Form("max_student")
present_student=request.Form("present_student")
week_course=request.Form("week_course")
day_course=request.Form("day_course")
opt1=request.Form("opt1")
opt2=request.Form("opt2")
rs.addnew
rs("cno")=course_number
rs("TermNum")=term_number
rs("MajorNum")=course_name
rs("MaxStu")=course_type
rs("PresentStu")=course_score
rs("week")=week_course
rs("day")=day_course
rs("hour")=day_hour
rs("wtow")=opt1&"-"&opt2
rs.update
response.Redirect("CoursePlan.asp")
%>
<body>
</body>
</html>
