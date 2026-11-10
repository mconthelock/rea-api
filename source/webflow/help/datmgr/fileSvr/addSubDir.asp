<!--#include virtual="/form/inc/system.inc"-->

<html>
<head>
<link href="http://<%=request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER%>/help/css/common.css" style="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftMargin="0" topMargin="0" marginWidth="0" marginHeight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2">
  <tr> 
    <td align="center" width="1%"><img src="/sys/imgs/quest_paper.gif" width="14" height="17"></td>
    <td><font color="#000000">Add Sub Directory</font></td>
  </tr>
  <tr> 
    <td align="center" width="1%">&nbsp;</td>
    <td nowrap valign="top"><font color="#000000"> เป็นส่วนสำหรับเพิ่ม Sub directory<br>
      ในหน้าจอ Add Sub Directory ประกอบไปด้วยส่วนต่างๆ ดังนี้</font></td>
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
                <td nowrap>Sub Directory Name</td>
                <td nowrap>สำหรับป้อนชื่อ Sub directory ที่ต้องการเพิ่ม</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>Owner</td>
                <td nowrap>สำหรับป้อนรหัสประจำตัวพนักงานของ Owner</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>ปุ่ม <font color="#000000"><a href="addRecv.asp"><img src="/form/imgs/help/datmgr/fileSvr/addBut.gif" border="0"></a></font></td>
                <td nowrap>สำหรับเพิ่ม Directory</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>ปุ่ม <font color="#000000"><img src="/form/imgs/help/datmgr/fileSvr/back.gif"  border="0"></font></td>
                <td nowrap>สำหรับย้อนกลับไปหน้าจอ Edit Server</td>
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
    <td align="center"><img src="/form/imgs/help/datmgr/fileSvr/addSubDir.gif"></td>
  </tr>
  <tr> 
    <td align="center"><img src="/sys/imgs/arrowup.gif"><font color="#000000">ภาพแสดงหน้าจอ 
      Add Sub Directory</font></td>
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
