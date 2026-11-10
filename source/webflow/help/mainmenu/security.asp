<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/form/inc/system.inc"-->

<%
Const FORM_NAME = "Network calendar user registration"
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link href="/<%=SYSTEM_FOLDER%>/help/css/common.css" rel="stylesheet" type="text/css">
</head>

<body text="#000000" leftMargin="0" topMargin="0" marginWidth="0" marginHeight="0">
<table width="100%">
  <tr> 
    <td colspan="2"><b><font color="#000000">Security Manager menu</font></b></td>
  </tr>
  <tr> 
    <td colspan="2">
      <hr noshade hight="1">
    </td>
  </tr>
  <tr> 
    <td colspan="2"><font color="#000000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;เป็นเมนูสำหรับแสดงข้อมูลของท่านในระบบ 
      เช่น การกำหนดรหัสผ่าน, การกำหนดให้มีการแจ้งเตือนด้วย E-Mail เป็นต้น<br>
      ซึ่งการที่ท่านจะสามารถใช้เมนูย่อยใดได้บ้างนั้นขึ้นกับสิทธิ์ที่ท่านมีอยู่ในขณะนั้น</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">การใช้งานเมนูมีขั้นตอนดังนี้</font></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <ol>
        <li><font color="#000000">เข้าสู่ระบบ <%=SYSTEM_TITLE%></font></li>
        <li><font color="#000000">เลือกเมนู Security Manager จะปรากฏเมนูย่อยดังนี้<br>
          2.1 Set representative สำหรับกำหนดผู้ทำงานแทน<br>
          2.2 Set E-Mail acknowledge สำหรับกำหนดให้มีการแจ้งเตือนด้วย E-Mail<br>
          2.3 User authentication สำหรับกำหนดสิทธิ์ผู้ใช้ระบบ</font></li>
        <li><font color="#000000">เลือกเมนูย่อยที่ท่านต้องการจะปรากฏหน้าจอสำหรับเมนูนั้น</font></li>
      </ol>
    </td>
  </tr>
</table>
</body>
</html>
