<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>历史课程安排信息录入</title>
</head>

<body>
<%
xueqihao = trim(request.form("xueqi"))
zhuanye = trim(request.form("zhuanye"))
sql="select * from oldCourseArrange"
set conn=server.createobject("adodb.connection")
conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
cid=request.form("course_number")
arr = split(cid,",")
m=ubound(arr)
for i=0 to m
'sql="insert into nextCourseArrange(cno,TermNum,MajorNum) values('"&trim(arr(i))&"','"&xueqihao&"','"&zhuanye&"')"
'conn.execute sql
rs.addnew
rs("cno")=trim(arr(i))
rs("TermNum")=xueqihao
rs("MajorNum")=zhuanye
rs.update
next
rs.close
set rs=nothing
conn.Close
Set conn=Nothing
response.write "<script>alert('选课完成!');location.href='oldCoursePlan.asp'</script>"
%>
</body>
</html>
