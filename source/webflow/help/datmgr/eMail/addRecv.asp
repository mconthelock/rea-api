<!--#include virtual="/form/inc/system.inc"-->

<html>
<head>
<link href="http://<%=request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER%>/help/css/common.css" style="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftMargin="0" topMargin="0" marginWidth="0" marginHeight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2">
  <tr> 
    <td align="center" width="1%"><img src="/sys/imgs/quest_paper.gif" width="14" height="17"></td>
    <td><font color="#000000">Add receiver</font></td>
  </tr>
  <tr> 
    <td align="center" width="1%">&nbsp;</td>
    <td nowrap valign="top"><font color="#000000"> เป็นส่วนสำหรับเพิ่ม E-Mail 
      ของผู้รับ <br>
      ในหน้าจอ Add receiver ประกอบไปด้วยส่วนต่างๆ ดังนี้</font></td>
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
                <td nowrap>Name</td>
                <td nowrap>แสดงชื่อผู้รับ</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap> E-Mail</td>
                <td nowrap>แสดง E-Mail ผู้รับ</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>ปุ่ม <font color="#000000"><a href="addRecv.asp"><img src="/form/imgs/help/datmgr/eMail/addBut.gif" border="0"></a></font></td>
                <td nowrap>สำหรับเพิ่ม E-Mail ของผู้รับ เมื่อคลิ๊กปุ่มนี้จะปรากฏหน้าจอสำหรับป้อน 
                  E-Mail และชื่อของผู้รับ</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>ปุ่ม <font color="#000000"><img src="/form/imgs/help/datmgr/eMail/back.gif"  border="0"></font></td>
                <td nowrap>สำหรับย้อนกลับไปหน้าสารบัญ</td>
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
    <td align="center"><img src="/form/imgs/help/datmgr/eMail/addRecv.gif"></td>
  </tr>
  <tr> 
    <td align="center"><img src="/sys/imgs/arrowup.gif"><font color="#000000">ภาพแสดงหน้าจอ 
      Add receiver</font></td>
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
