
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>FIXED ASSET REGISTER DETAIL</title>
<meta http-equiv="X-UA-Compatible" content="IE=Edge" />
<link href="../css/templatemo_style.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" src="../../js/window.js"></script>
</head>
<body>

<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/misc.inc"-->

<%
	Dim rsOra, sqlOra
	Dim mode, requester
	Dim totalAdv, totalExp, cntLine, i
	Dim grp, grpdesc, lc, lcdesc
	Dim rsOra2
	
	set rsOra = createObject("adodb.recordset")
	set rsOra2 = createObject("adodb.recordset")
	
%>	
<div id="templatemo_body_wrapper">
    <div id="templatemo_main">
	<%
				sqlOra = "select t.* from fxa_tempfaMst t , fxa_locemp l where t.ID = l.ID and l.nfrmno = '"&request.querystring("no")&"' " _
						& "and l.vorgno = '"&request.querystring("orgNo") _
						&"' and l.cyear='"&request.querystring("y") _
						&"' and l.cyear2='"&request.querystring("y2") _
						&"' and nrunno = '"&request.querystring("runNo")&"'"
				rsOra.open sqlOra, connOra	
				'response.write "sql = "&sqlOra&"<br/>"
				if NOT rsOra.EOF then
					grp = rsOra("GRPCODE")
					lc = rsOra("LOCCODE")
					sqlOra = "select ACCASSNAME from fxa_accassMst where ACCASSCODE='"&grp&"'"
					rsOra2.open sqlOra, connOra
					grpdesc = rsOra2("ACCASSNAME")
					rsOra2.close
					sqlOra = "select LOCNAME from fxa_locMst where LOCCODE='"&lc&"'"
					rsOra2.open sqlOra, connOra
					lcdesc = rsOra2("LOCNAME")
					rsOra2.close
				end if
				
				
			%>
			<div id="contact_form">
			<form name="frmMain" method="post" action="cgi-bin/app.asp">
			<table width="100%" border="1" bgcolor="#fbe9ef" bordercolor="#FFFFFF" cellspacing="2" cellpadding="0">
			 <tr bgcolor="#FFFFFF"><td colspan="15"><strong>ASSETS INSPECTION SHEET</strong></td></tr>
			 <tr>
				<td colspan="15">
					<table>
					<tr><td colspan="2" width="300"><strong>Group: </strong><%=grp%></td><td colspan="13" width="600"><strong>Group Description: </strong><%=grpdesc%></td></tr>
					<tr><td  colspan="2" width="300"><strong>L/C: </strong><%=lc%></td><td colspan="13" width="600"><strong>Description L/C:</strong><%=lcdesc%></td></tr>
					</table>
				</td>
			 </tr>
			  <tr bgcolor="#ec858f" align="center">
				<td><strong>ASSET NO.</strong></td>
				<td><strong>DESCRIPTION</strong></td>
				<td><strong>DOC.DATE</strong></td>
				<td><strong>INITIAL VALUE</strong></td>
				<td><strong>INV.NO</strong></td>
				<td><strong>MODEL NO</strong></td>
				<td><strong>S/N</strong></td>
				<td><strong>P/O</strong></td>
				<td><strong>REF.ASSET</strong></td>
				<td><strong>SUPPLIER NAME</strong></td>
				<td><strong>PR NO</strong></td>
				<td><strong>BUDGET NO.</strong></td>
				<td><strong>REQUEST BY</strong></td>
				<td><strong>QTY</strong></td>
				<td><strong>UNIT</strong></td>
			  </tr>
			<%
				i=0
				Do While Not rsOra.EOF
						
			%>
			  <tr>
			  <td><%=rsOra("ASSETNO")%></td>
			  <td><%=rsOra("ASSDESC")%></td>
			  <td><%=rsOra("DOCDATE")%></td>
			  <td align="right"><%=rsOra("INITIALVALUE")%></td>
			  <td><%=rsOra("INVNO")%></td>
			  <td><%=rsOra("MODELNO")%></td>
			  <td><%=rsOra("SNNO")%></td>
			  <td><%=rsOra("PONO")%></td>
			  <td><%=rsOra("REFASSET")%></td>
			  <td><%=rsOra("SUPNAME")%></td>
			  <td><%=rsOra("PRNO")%></td>
			  <td><%=rsOra("BUDGETNO")%></td>
			  <td><%=rsOra("REQBY")%></td>
			  <td align="right"><%=rsOra("QTY")%></td>
			  <td><%=rsOra("UNIT")%></td>
			  </tr>
			 <% 
				rsOra.MoveNext	
				i = i+1
			  	loop
			 
			 %>
			  <input type="hidden" name="maxs" value="<%=i%>" />
			 </table>
	</div>
	</div>
    <div id="templatemo_footer_wrapper">
        <div id="templatemo_footer">
            <!--#include file="../../inc/copyRight.inc"--> 
        </div>
</div>
</div>
</body>
</html>