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
    <td colspan="2"><b><font color="#0000FF">System process overview</font></b></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <hr noshade height="1">
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <p><font color="#000000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ในส่วนนี้จะอธิบายขั้นตอนการทำงานในระบบ 
        เพื่อให้ท่านเข้าใจในลำดับ<br>
        ขั้นตอนต่างๆ ซึ่งจะช่วยให้ท่านสามารถใช้ระบบได้อย่างมีประสิทธิภาพ</font><font color="#000000"><br>
        การทำงานของระบบมีขั้นตอนดังนี้</font></p>
      </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <ol>
        <li>เข้าสู่ระบบ <%=SYSTEM_TITLE%></li>
        <li><font color="#000000">สร้างคำร้องจากเมนู Create</font></li>
        <li><font color="#000000">ตรวจสอบคำร้องจากเมนู Mine</font></li>
        <li><font color="#000000">ตรวจสอบคำร้องที่กำลังจะมาถึงท่านจากเมนู Coming 
          soon </font></li>
        <li><font color="#000000">ตรวจสอบคำร้องที่กำลังรอการอนุมัติจากท่านจากเมนู 
          Waiting for your approval</font></li>
        <li><font color="#000000">ตรวจสอบคำร้องที่ท่านได้อนุมัติหรือยกเลิกจากเมนู 
          Approve/Reject Status</font></li>
        <li><font color="#000000">ตรวจสอบคำร้องที่ตัวแทนของท่านได้อนุมัติหรือยกเลิกจากเมนู 
          Representative Approve/Reject Status</font></li>
        <li><font color="#000000">ตรวจสอบคำร้องที่ท่านได้อนุมัติหรือยกเลิกจากเมนู 
          Finish</font></li>
        <li><font color="#000000">ค้นหารายการที่ท่านได้ขอหรือขอแทนพนักงานท่านอื่นที่ได้รับการอนุมัติหรือยกเลิกแล้วเป็นเวลานานกว่า<br>
          ที่กำหนด (ปรกติรายการที่ได้รับการอนุมัติหรือยกเลิกแล้วจะแสดงอยู่ในเมนู 
          Finish, Approve/Reject Status, Representative<br>
          Approve/Reject Status 6 วัน)<br>
          </font></li>
      </ol>
    </td>
  </tr>
</table>
</body>
</html>
