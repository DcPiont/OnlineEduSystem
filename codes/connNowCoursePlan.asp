<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>���ӱ�ѧ�ڿγ̰��ű�</title>
</head>
<%
'�������ݿ�
 sql  ="select nowCourseArrange.cno,cname,TermNum,MajorNum,MaxStu,PresentStu,week,day,hour,wtow,place from nowCourseArrange,CourseInfo where nowCourseArrange.cno=CourseInfo.cno "
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
%>
<body>
</body>
</html>