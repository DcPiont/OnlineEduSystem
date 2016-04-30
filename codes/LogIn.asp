<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>用户登录</title>
</head>

<body>
 <%
 thename = Request.form("userId")
 themima = Request.form("userPass")
 session("userid")=thename
 session("password")=themima
sql  ="select * from UserID where uid='"&thename&"'"
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")
 rs.open sql,conn,1,3
 if not rs.EOF then
  rs.movefirst
 if  trim(themima) <> trim(rs("pwd")) then
 response.write("<script type=text/javascript>alert('用户名或密码错误！');history.back();</script>")
 else
 response.write("<script type=text/javascript>alert('登录成功！');window.location.href='index.html';</script>")
 end if
 else
 response.write("<script type=text/javascript>alert('用户不存在！');history.back();</script>")
 end if
 rs.close
 conn.close 
 %>
</body>
</html>
