<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ѧ����Ϣ���</title>
</head>
<%
'�������ݿ�
sql  ="select * from student "
 set conn=server.createobject("adodb.connection")
 conn.open "provider=sqloledb;data source=localhost\SQLEXPRESS;uid=sa;pwd=11111;database=web"
 set rs=server.createobject("adodb.recordset")	
rs.open sql,conn,1,3
'�������
student_number=request.Form("xuehao")
sname=request.Form("xuesname")
banji=request.Form("class")
ID=request.Form("ID")
sex=request.Form("sex")
major=request.Form("zhuaniye")
rs.addnew
rs("sno")=student_number
rs("sname")=sname
rs("class")=banji
rs("id")=ID
rs("sex")=sex
rs("major")=major
rs.update
response.write("<script type=text/javascript>alert('��ӳɹ���');window.location.href='studentup.asp';</script>")
%>
<body>
</body>
</html>
