<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/misc.inc"-->

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<%
Set rsOra=Createobject("ADODB.Recordset")
%>
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title>Form list</title></head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">
<form name="frmMain" method="get" action="aloci.asp">
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
            <td align="right" valign="bottom"><img src="/form/imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="/form/imgs/border/book/t.gif" width="100%" align="center"><img src="/form/imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="/form/imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr>
            <td background="/form/imgs/border/book/l.gif" align="right">&nbsp;</td>
            
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Add location</font></b></td>
            <td background="/form/imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="/form/imgs/border/book/l.gif" align="right"><img src="/form/imgs/border/book/l.gif" width="17" height="18"></td>
            
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr class="hdrRow"> 
            <td nowrap valign="top" width="0" bgcolor="#FFCC00"> <b><font color="#000000">LC Code</font></b></td>
            <td nowrap width="100%" bgcolor="#FFFFCC"><b> <input type="text" name="locCode">
              </b></td>
          </tr>
          <tr> 
            <td nowrap width="0" bgcolor="#FFCC00" valign="top"><b><font color="#000000">Location Name</font></b></td>
            <td nowrap width="100%" bgcolor="#FFFFCC"> 
              <textarea name="locName" cols="60" rows="5"></textarea>
            </td>
          </tr>
			  <%
				sqlOra = "select * from amec.pposition where sposcode in('10', '11','20','21', '30')"
				'response.write sqlOra & "<br>"
				rsOra.open sqlOra, connOra
				
				if not rsOra.EOF then
				%>
          <tr> 
			<td bgcolor="#FFCC00" width="1%" nowrap><b><font color="#000000">Position</font></b></td>
			<td bgcolor="#FFFFCC"> 				
			<select name = "selPos">
				<%
				do while not rsOra.EOF 
					if posNo = rsOra("sposcode") then 
						selected = "selected"
					else
						selected = ""
					end if
					response.write "<option value=" & rsOra("sposcode") & " " & selected & ">" & rsOra("sposcode") & ": " & rsOra("sPosition") & "</option>"
					rsOra.moveNext
				loop
				%>
				 </select><br>
				<%
				end if
				rsOra.close
				%>
				</td>
			  </tr>		 	  
			<%		  			
				sqlOra = "select " _
				& "distinct orgPos.vOrgNo, vName " _
				& "from " _
				& "orgPos, amec.pPosition, vOrgMst " _
				& "where " _
				& "vOrgMst.vOrgNo = orgPos.vOrgNo and sPosCode = vPosNo and (vPosNo >= '10' and vPosNo <='30') and vname not like '%Cancel%' order by vname"
				
				rsOra.open sqlOra, connOra
				%>
	          <tr> 
			<td bgcolor="#FFCC00" width="1%" nowrap><b><font color="#000000">Organize</font></b></td>
			<td bgcolor="#FFFFCC"> 
				<%
				if not rsOra.EOF then
				%>
				<select name = "selOrg">
				<%
				do while not rsOra.EOF 
					if ownorgNo = rsOra("vOrgNo") then
						selected = "selected"
					else
						selected = ""
					end if
					response.write "<option value=" & rsOra("vOrgNo") & " " & selected & ">" & rsOra("vName") & "</option>"
					rsOra.moveNext
				loop
				%>
				 </select>
				<%
				end if
				rsOra.close
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
                
        <input type="submit" value="Add">
                
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
<!--#include virtual="/form/inc/copyRight.inc"-->
</body>
</html>

<%
set rsOra = nothing
%>