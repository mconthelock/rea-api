<html>
<head>
<link href="/sys/css/help.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftMargin="0" topMargin="0" marginWidth="0" marginHeight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2">
  <tr> 
    <td align="center" width="1%"><img src="/sys/imgs/quest_paper.gif" width="14" height="17"></td>
    <td><font color="#000000">File Server: Server Master</font></td>
  </tr>
  <tr> 
    <td align="center" width="1%">&nbsp;</td>
    <td nowrap valign="top"> 
      <p><font color="#000000"> เป็นส่วนสำหรับจัดการข้อมูลของเซิร์ฟเวอร์และไดเร็กทอรี่ที่จะปรากฏในฟอร์ม 
        File Server User Requisition <br>
        เพื่อให้ผู้ใช้ได้เลือกไดเร็กทอรี่ที่ต้องการขอใช้งาน และใช้ในการกำหนดผู้ที่มีหน้าที่รับผิดชอบไดเร็กทอรี่ที่จะปรากฏอยู่<br>
        ในขั้น Directory Owner ในโฟล์</font><font color="#000000"><br>
        ขั้นตอนในการทำงานมีดังนี้</font></p>
      </td>
  </tr>
  <tr valign="top"> 
    <td align="center" colspan="2"><img src="/sys/imgs/spacer.gif" height="4"></td>
  </tr>
  <tr> 
    <td align="center" width="1%">&nbsp;</td>
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="1" cellpadding="2">
              <tr bgcolor="#FFFFFF"> 
                <td><font color="#000000"></font></td>
                <td nowrap bgcolor="#FFFFFF"><font color="#000000"><b>Column</b></font></td>
                <td width="100%" bgcolor="#FFFFFF"><font color="#000000"><b>Description</b></font></td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap><a href="addSvr.asp">เพิ่มเซิร์ฟเวอร์</a></td>
                <td nowrap> 
                  <p><font color="#000000">เป็นการสร้างเซิร์ฟเวอร์เพื่อบรรจุไดเร็กทอรี่สำหรับผู้ใช้เลือกในการ 
                    request <br>
                    ซึ่งท่านสามารถแก้ไขข้อมูลของเซิร์ฟเวอร์ในภายหลังได้จากเมนู 
                    <a href="editSvr.asp">Edit..</a></font></p>
                </td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap><font color="#000000"><a href="addRootDir.asp">เพิ่มไดเร็กทอรี่รากและกำหนดผู้รับผิดชอบ</a></font></td>
                <td nowrap><font color="#000000">เป็นการสร้างไดเร็กทอรี่ชั้นแรกสุดและกำหนดผู้ที่รับผิดชอบที่มีหน้าที่อนุมัติคำร้อง<br>
                  ขอใช้ไดเร็กทอรี่นั้น <br>
                  ซึ่งท่านสามารถแก้ไขข้อมูลของเซิร์ฟเวอร์ในภายหลังได้จากเมนู <a href="editDir.asp">Edit...</a></font></td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap><a href="addSubDir.asp">เพิ่มไดเร็กทอรี่ย่อยและกำหนดผู้รับผิดชอบ</a></td>
                <td nowrap><font color="#000000">เป็นการสร้างไดเร็กทอรี่เช่นเดียวกับไดเร็กทอรี่รากแต่เป็นไดเร็กทอรี่ที่อยู่ซ้อนกัน<br>
                  เป็นลำดับชั้นโดยไม่จำเป็นต้องเป็นไดเร็กทอรี่รากเท่าน้นและกำหนดผู้ที่มีหน้าที่รับ<br>
                  ผิดชอบในไดเร็กทอรี่นั้น <br>
                  ซึ่งท่านสามารถแก้ไขข้อมูลของเซิร์ฟเวอร์ในภายหลังได้จากเมนู <a href="editDir.asp">Edit...</a></font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="right"><a href="../../toc.asp">Back</a></td>
  </tr>
</table>
<br><!--#include virtual="/sys/inc/copyRight.inc"--><br>&nbsp;
</body>
</html>
