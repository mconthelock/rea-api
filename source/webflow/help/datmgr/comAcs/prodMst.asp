<!--#include virtual="/form/inc/system.inc"-->

<html>
<head>
<link href="http://<%=request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER%>/help/css/common.css" style="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftMargin="0" topMargin="0" marginWidth="0" marginHeight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2">
  <tr> 
    <td align="center" width="1%"><img src="/sys/imgs/quest_paper.gif" width="14" height="17"></td>
    <td><font color="#000000">Product master</font></td>
  </tr>
  <tr> 
    <td align="center" width="1%">&nbsp;</td>
    <td nowrap valign="top"><font color="#000000"> เป็นส่วนแสดงรายชื่อ Product 
      ที่อยู่ใน Expense type ประเภท Computer accessory <br>
      ที่ต้องการให้แสดงในฟอร์ม Computer accessory การกำหนดให้แสดง Product ใดสามารถทำได้<br>
      โดยการคลิ๊กที่ Checkbox ให้ปรากฏเครื่องหมายถูกและกดปุ่มบันทึก<br>
      ในหน้าจอ Product master ประกอบไปด้วยส่วนต่างๆ ดังนี้</font></td>
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
                <td nowrap><img src="/form/imgs/help/datMgr/comAcs/checkBox.gif"></td>
                <td nowrap>สำหรับเลือก Product ที่ต้องการให้แสดงในฟอร์ม Computer 
                  accessory </td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap><font color="#000000">Organize</font></td>
                <td nowrap> 
                  <p><font color="#000000">แสดงลำดับรายการ</font></p>
                </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap><font color="#000000">Product Code</font></td>
                <td nowrap><font color="#000000">แสดงรหัส Product ที่ใช้ในระบบ 
                  E-Procurement</font></td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>Product Name</td>
                <td nowrap>แสดงชื่อ Product</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>Expense Type</td>
                <td nowrap>แสดงรหัส Expense type ของ Product</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>Group in Expense Type</td>
                <td nowrap>แสดงรหัสกลุ่มใน Expense type ของ Product</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>Sup-group in Expense Type</td>
                <td nowrap>แสดงรหัสกลุ่มย่อยใน Expense type ของ Product</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>UPrice</td>
                <td nowrap>แสดงราคาต่อหน่วยของ Product</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>ปุ่ม <font color="#000000"><img src="/form/imgs/help/datmgr/comAcs/saveChange.gif" border="0"></font></td>
                <td nowrap>สำหรับบันทึกรายชื่อ Product ที่ต้องการให้แสดงในฟอร์ม 
                  Computer accessory</td>
              </tr>
              <tr bgcolor="#FFFFFF" valign="top"> 
                <td><font color="#000000"><img src="/form/imgs/handpointer.gif" ></font></td>
                <td nowrap>ปุ่ม <font color="#000000"><img src="/form/imgs/help/datmgr/comAcs/back.gif"  border="0"></font></td>
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
    <td align="center"><img src="/form/imgs/help/datmgr/comAcs/prodMst.gif"></td>
  </tr>
  <tr> 
    <td align="center"><img src="/sys/imgs/arrowup.gif"><font color="#000000">ภาพแสดงหน้าจอ 
      Product master</font></td>
  </tr>
</table>
<br>
<dd>ในกรณีที่มี Product จำนวนมากท่านสามารถเลือกแสดงเฉพาะ Product ที่ท่านต้องการได้โดยการพิมพ์รหัสหรือชื่อของ 
Product ที่ท่านต้องการและกดปุ่มค้นหาในหน้าจอส่วนค้นหาที่ปรากฏอยู่ด้านล่างของหน้าจอ Product 
master<br></dd>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="center"><img src="/form/imgs/help/datmgr/comAcs/findProd.gif"></td>
  </tr>
  <tr> 
    <td align="center"><img src="/sys/imgs/arrowup.gif"><font color="#000000">ภาพแสดงหน้าจอ 
      Product search</font></td>
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
