<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>教师信息录入</title>
</head>

<body>
<%
'连接数据库
sql  ="select * from teacher "
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")	
rs.open sql,conn,1,3
'添加数据
tno=request.Form("tno")
tname=request.Form("tname")
ttel=request.Form("ttel")
ID=request.Form("ID")
sex=request.Form("sex")
rs.addnew
rs("tno")=tno
rs("tname")=tname
rs("ttel")=ttel
rs("tid")=ID
rs("sex")=sex
rs.update
response.write("<script type=text/javascript>alert('添加成功！');window.location.href='teacherup.asp';</script>")
%>
</body>
</html>
