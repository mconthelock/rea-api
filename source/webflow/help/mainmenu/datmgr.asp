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
    <td colspan="2"><font color="#000000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;เป็นเมนูสำหรับแสดงข้อมูลสำหรับการทำงานของระบบ 
      เช่น การสร้างฟอร์ม, การสร้างขั้นตอนการไหล เป็นต้น<br>
      ซึ่งการที่ท่านจะสามารถใช้เมนูย่อยใดได้บ้างนั้นขึ้นกับสิทธิ์ที่ท่านมีอยู่ในขณะนั้น</font></td>
  </tr>
  <tr> 
    <td><font color="#000000">การใช้งานเมนูมีขั้นตอนดังนี้</font></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <ol>
        <li><font color="#000000">เข้าสู่ระบบ <%=SYSTEM_TITLE%></font></li>
        <li><font color="#000000">เลือกเมนู Data Manager จะปรากฏเมนูย่อยดังนี้<br>
          2.1 Form Master สำหรับจัดการข้อมูลของแบบฟอร์มทั้งหมดในระบบ<br>
          2.2 Step Master สำหรับจัดการข้อมูลของขั้นตอนทั้งหมดในระบบที่จะนำไปใช้ในการสร้างโฟล์ใน 
          Form Master<br>
          2.3 Organize Master สำหรับจัดการข้อมูลความสัมพันธ์ของพนักงาน, ตำแหน่ง 
          และหน่วยงานเพื่อใช้ในการเลือกผู้อนุมัติ<br>
          2.4 Manager Master สำหรับกำหนดความสัมพันธ์ของพนักงานและหัวหน้าในระดับต่างๆ 
          เพื่อ<br>
          ใช้ในการเลือกผู้อนุมัติในกรณีที่ไม่สามารถหาต้นสังกัดจากข้อมูลในข้อ 2.3 
          ของพนักงานที่เป็นผู้ขอได้<br>
          2.5 Product Master สำหรับจัดการข้อมูลของผลิตภัณฑ์ที่ต้องการให้แสดงในแบบฟอร์ม 
          Computer Accessory Requisition <br>
          2.6 Server Master สำหรับจัดการข้อมูลของเซิร์ฟเวอร์และไดเร็กทอรี่ที่ต้องการให้แสดงในแบบฟอร์ม 
          File Server User Requisition<br>
          2.7 Set E-Mail to MELINA สำหรับจัดการ E-Mail ของผู้ที่รับผิดชอบ E-Mail 
          ที่ MELINA<br>
          2.8 Main Approver Master สำหรับจัดการข้อมูลของ Main Approver ในแบบฟอร์ม 
          Intranet Content Input<br>
          </font></li>
        <li><font color="#000000">เลือกเมนูย่อยที่ท่านต้องการจะปรากฏหน้าจอสำหรับเมนูนั้น</font></li>
      </ol>
    </td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
</table>
</body>
</html>
