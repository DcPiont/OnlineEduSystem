<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>密码修改</title>
</head>

<body>
<%
 oldpwd= Request.form("jiumima")
 newpwd = Request.form("newuserPass1")
 input_again = Request.form("newuserPass2")
 sql  ="select * from UserID where uid ='"&session("userid")&"'"
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,3
  rs.movefirst
  if  trim(oldpwd) <> trim(rs("pwd")) then
 	response.write("<script>alert('旧密码验证失败，请重新输入！');history.back();</script>")
 else
	if trim(newpwd) <> trim(input_again) then
		response.write("<script>alert('两次输入的密码不符，请重新输入！');history.back();</script>")
 	else
 		sqlupdate= "update UserID set pwd ='"&newpwd&"' where uid ='"&session("userid")&"'"
		conn.execute sqlupdate
 		response.write("<script type=text/javascript>alert('密码修改成功！');window.location.href='PWDupdate.html';</script>")
 	end if
 end if  
 rs.update
 rs.close
 conn.close  
%>
</body>
</html>
