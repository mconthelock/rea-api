<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/authen.inc"-->
<!--#include virtual="/sys/inc/userUtil.inc"-->
<!--#include virtual="/form/inc/formUtil.inc"-->
<!--#include virtual="/sys/inc/date.inc"-->
<!--#include virtual="/sys/inc/date2.inc"-->
<!--#include virtual="/sys/inc/misc.inc"-->

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link href="/sys/css/common.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/<%=SYSTEM_FOLDER%>/js/date.js"></script>
<style>
.spfont /*Search Panel Font*/
{
	font-family: ms sans serif; 
	font-size: 8pt;
}
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">
<%
Dim rsOra, sqlOra
set rsOra = createObject("adodb.recordset")
%>

<!--########## REPORT QUERY ##########-->
<form name="frmMain" action="finResultSchFX.asp">
	<input type="hidden" name="s" value="1">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><img src="/<%=SYSTEM_FOLDER %>/imgs/border/line/tlc.gif" width="17" height="18"></td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/t.gif" width="100%"><img src="/<%=SYSTEM_FOLDER %>/imgs/border/line/t.gif" width="17" height="18"></td>
    <td><img src="/<%=SYSTEM_FOLDER %>/imgs/border/line/trc.gif" width="17" height="18"></td>
  </tr>
  <tr>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/l.gif">&nbsp;</td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/bg.gif" align="center"><b><font color="#000000">SEARCH ASSET (ค้นหาสินทรัพย์)</font></b></td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/r.gif">&nbsp;</td>
  </tr>
 
  <tr> 
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/l.gif"><img src="/<%=SYSTEM_FOLDER %>/imgs/border/line/l.gif" width="17" height="18"></td>
    <td bgcolor="#0000FF"> 
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font class="spfont"> A</font><font class="spfont">sset no (หมายเลขสินทรัพย์):</font></b></td>
            <td bgcolor="#FFFFCC"><input type="text" name="txtAssNo"></td>
        </tr>
        <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font class="spfont"> A</font><font class="spfont">sset name (ชื่อสินทรัพย์):</font></b></td>
            <td bgcolor="#FFFFCC"><input type="text" name="txtAssName"></td>
        </tr>		
        <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font class="spfont"> L</font><font class="spfont">ocation (พื้นที่สินทรัพย์):</font></b></td>
            <td bgcolor="#FFFFCC">
			<%
				  sqlOra = "select * from fxa_locMst order by locName"
				  'response.write sqlOra
				  rsOra.open sqlOra, connOra			
				 %> 
				<select name="selLocCode">
				<option></option>
				<%
					do while not rsOra.EoF
						response.write "<option value=" & rsOra("locCode") & ">" & rsOra("locName") & "</option>"
						rsOra.MoveNext
					loop
					rsOra.Close
				%>
				</select>			
				<input type="text" name="txtLocName">
			</td>
        </tr>
        <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font class="spfont"> T</font><font class="spfont">ype Asset (ประเภทของสินทรัพย์):</font></b></td>
            <td bgcolor="#FFFFCC">
				  <%
				  sqlOra = "select * from fxa_grpMst order by grpdesc"
				  'response.write sqlOra
				  rsOra.open sqlOra, connOra			
				 %> 
				<select name="selGrpCode">
				<option></option>
				<%
					do while not rsOra.EoF
						response.write "<option value=" & rsOra("grpCode") & ">" & rsOra("grpdesc") & "</option>"
						rsOra.MoveNext
					loop
					rsOra.Close
				%>
				</select>
			</td>
        </tr>	
        <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font class="spfont"> PO N</font><font class="spfont">o:</font></b></td>
            <td bgcolor="#FFFFCC"><input type="text" name="txtPoNo"></td>
        </tr>			
        <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font class="spfont"> PR N</font><font class="spfont">o:</font></b></td>
            <td bgcolor="#FFFFCC"><input type="text" name="txtPrNo"></td>
        </tr>			
        <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font class="spfont"> D</font><font class="spfont">ocument Date:</font></b></td>
            <td bgcolor="#FFFFCC">
			<%showDateLst "selReqDate", request.queryString("selReqDate"), "selReqMonth", request.queryString("selReqMonth"), "selReqYear", request.queryString("selReqYear"), "2010", year(date), true, "", "onChange=" & QUOTE & "javascript: showDateLst(document.frmMain.selReqDate, document.frmMain.selReqMonth.options[document.frmMain.selReqMonth.selectedIndex].value-1, document.frmMain.selReqYear.options[document.frmMain.selReqYear.selectedIndex].value-1)" & QUOTE, ""%>
		  <font color="#000000">To</font>			  
		  <%showDateLst "selApvDate2", request.queryString("selApvDate2"), "selApvMonth2", request.queryString("selApvMonth2"), "selApvYear2", request.queryString("selApvYear2"), "2010", year(date), true, "", "onChange=" & QUOTE & "javascript: showDateLst(document.frmMain.selApvDate2, document.frmMain.selApvMonth2.options[document.frmMain.selApvMonth2.selectedIndex].value-1, document.frmMain.selApvYear.options[document.frmMain.selApvYear2.selectedIndex].value-1)" & QUOTE, ""%>			  		  			
			</td>
        </tr>
      </table>
    </td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/r.gif"><img src="/<%=SYSTEM_FOLDER %>/imgs/border/line/r.gif" width="17" height="18"></td>
  </tr>
  <tr> 
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/l.gif">&nbsp;</td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/bg.gif">&nbsp;</td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/r.gif">&nbsp;</td>
  </tr>
  <tr> 
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/l.gif">&nbsp;</td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/bg.gif" align="center"> 
      <input type="submit" value="Search">
      <input type="button" value="Back" onClick="javascript: window.location.href = '/<%=SYSTEM_FOLDER %>/report/index.asp'">
    </td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/r.gif">&nbsp;</td>
  </tr>
  <tr> 
    <td><img src="/<%=SYSTEM_FOLDER %>/imgs/border/line/blc.gif" width="17" height="18"></td>
    <td background="/<%=SYSTEM_FOLDER %>/imgs/border/line/b.gif"><img src="/<%=SYSTEM_FOLDER & "/" & SUB_SYSTEM_FOLDER%>/imgs/border/line/b.gif" width="17" height="18"></td>
    <td><img src="/<%=SYSTEM_FOLDER %>/imgs/border/line/brc.gif" width="17" height="18"></td>
  </tr>
</table>
</form>
<!--########## END REPORT QUERY ##########-->



<!--#include virtual="/sys/inc/copyright.inc"-->
&nbsp;
</body>
</html>
