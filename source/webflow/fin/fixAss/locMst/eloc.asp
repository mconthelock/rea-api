<%Option Explicit%>

<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/misc.inc"-->

<%
Dim rsOra, rsOra1, sqlOra
Dim i
Dim selected

set rsOra = createObject("adodb.recordset")
set rsOra1 = createObject("adodb.recordset")
sqlOra = "select " _
& "locCode as locCode, locName, vorgno, sposcode " _
& "from fxa_locMst where locCode ='" & request("locCode") & "' " _
& "order by locCode"
'response.write sqlOra & "<br>"
'response.end
rsOra.open sqlOra, connOra
%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title>Form list</title></head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">
<%
if rsOra.EOF then
	response.write "Data not found!"
else
%>
<form name="frmMain" method="get" action="updedloc.asp">
	<input type="hidden" name="locCodeO" value="<%=rsOra("locCode")%>">    
<!--	<input type="hidden" name="locNameO" value="<%'=rsOra("locName")%>">    -->
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
            <td align="right" valign="bottom"><img src="/form/imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="/form/imgs/border/book/t.gif" width="100%" align="center"><img src="/form/imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="/form/imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr>
            <td background="/form/imgs/border/book/l.gif" align="right">&nbsp;</td>
            
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Edit Location</font></b></td>
            <td background="/form/imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="/form/imgs/border/book/l.gif" align="right"><img src="/form/imgs/border/book/l.gif" width="17" height="18"></td> 
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr> 
            <td bgcolor="#FFCC00" nowrap width="0" valign="top"><b><font color="#000000">Loc Code</font></b></td>
            <td bgcolor="#FFFFCC" nowrap width="100%"><input type="text" name="locCode" value=" <%=rsOra("locCode")%>"></td>
          </tr>
          <tr> 
            <td bgcolor="#FFCC00" nowrap width="0" valign="top"><b><font color="#000000">Location</font></b></td>
            <td bgcolor="#FFFFCC" nowrap width="100%"><font color="#000000"> 
              <textarea name="locName" cols="60" rows="3"><%=rsOra("locName")%></textarea>
              </font></td>
          </tr>
          <tr> 
            <td bgcolor="#FFCC00" nowrap width="0" valign="top"><b><font color="#000000">Position</font></b></td>
            <td bgcolor="#FFFFCC" nowrap width="100%">
			<%
				sqlOra = "select * from amec.pposition where sposcode in ('10', '11','20','21', '30')"
				'response.write sqlOra & "<br>"
				rsOra1.open sqlOra, connOra
			%>
			<select name = "selPos">
				<option></option>
				<%
				if not rsOra1.EOF then
				do while not rsOra1.EOF 
					if rsOra1("sposcode") = rsOra("sposcode") then 
						selected = "selected"
					else
						selected = ""
					end if
					response.write "<option value=" & rsOra1("sposcode") & " " & selected & ">" & rsOra1("sposcode") & ": " & rsOra1("sPosition") & "</option>"
					rsOra1.moveNext
				loop
				end if
				rsOra1.close				
				%>
				 </select></td>
          </tr>
          <tr> 
            <td bgcolor="#FFCC00" nowrap width="0" valign="top"><b><font color="#000000">Organization</font></b></td>
            <td bgcolor="#FFFFCC" nowrap width="100%">
			<%		  			
				sqlOra = "select " _
				& "distinct orgPos.vOrgNo, vName " _
				& "from " _
				& "orgPos, amec.pPosition, vOrgMst " _
				& "where " _
				& "vOrgMst.vOrgNo = orgPos.vOrgNo and sPosCode = vPosNo and (vPosNo >= '10' and vPosNo <='30') and vname not like '%Cancel%' order by vname"
				
				rsOra1.open sqlOra, connOra

				if not rsOra.EOF then
				%>
				<select name = "selOrg">
				<option></option>
				<%
				do while not rsOra1.EOF 
					if rsOra1("vOrgNo") = rsOra("vOrgNo") then
						selected = "selected"
					else
						selected = ""
					end if
					response.write "<option value=" & rsOra1("vOrgNo") & " " & selected & ">" & rsOra1("vName") & "</option>"
					rsOra1.moveNext
				loop
				%>
				 </select>
				<%
				end if
				rsOra1.close
				%>
			</td>
          </tr>		  		  
        </table>
            </td>
            <td background="/form/imgs/border/book/r.gif" align="left"><img src="/form/imgs/border/book/r.gif" width="17" height="18"></td>
        </tr>
        <tr> 
            <td background="/form/imgs/border/book/l.gif" align="right"><img src="/form/imgs/border/book/l.gif" width="17" height="18"></td>
            <td background="/form/imgs/border/book/bg.gif">&nbsp;</td>
            <td background="/form/imgs/border/book/r.gif" align="left"><img src="/form/imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td background="/form/imgs/border/book/l.gif" align="right"><img src="/form/imgs/border/book/l.gif"></td>
            <td align="center" background="/form/imgs/border/book/bg.gif"> 
                
        <input type="submit" value="Save change" >
                
        <input type="button" value="Cancel" onClick="javascript: window.location.href = 'locMst.asp'">
            </td>
            <td background="/form/imgs/border/book/r.gif" align="left"><img src="/form/imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td align="right" valign="top"><img src="/form/imgs/border/book/blc.gif" width="17" height="18"></td>
            <td background="/form/imgs/border/book/b.gif"><img src="/form/imgs/border/book/b.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="/form/imgs/border/book/brc.gif" width="17" height="18"></td>
        </tr>
    </table>
</form>
<%
end if

rsOra.close
%>
<!--#include virtual="/form/inc/copyRight.inc"-->
</body>
</html>

<%
set rsOra = nothing
%>