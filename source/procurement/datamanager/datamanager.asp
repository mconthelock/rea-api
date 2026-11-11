<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/Program.asp" -->


<%
	' Author: Ekkapan Kongchoo
	
    Dim  Rs,sSQL,sSQL2,RS2
	Dim sScriptPath,sexptype,sPoscode
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM pexpenstype  ORDER BY sexptype "
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%>
				<%Session("ListProduct")=""%>
				<%Session("ListCattype")=""%>
				<%Session("ListCatalog")=""%>
				<%Session("SearchProduct")=""%>
				<%Session("Pageno")=""%>
<html>
<head>
<title>Data Manager</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#DB4B37" vlink="#CC3333" alink="#DB4B37" text="#CC3333">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="1%" height="2"><img src="images/ccb.gif" width="10" height="10"></td>
    <td width="63%" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="3"></font></td>
    <td width="36%" height="2" align="center"><img src="images/spacer.gif" width="8" height="8"></td>
  </tr>
  <tr>
    <td width="1%">&nbsp;</td>
    <td width="63%" valign="top" height="30"><font face="Verdana, Arial, Helvetica, sans-serif" size="3"><b><font size="+1" face="Verdana, Arial, Helvetica, sans-serif">Data 
      Manager</font> </b></font></td>
    <td width="36%" align="center" valign="top">&nbsp;</td>
  </tr>
  <tr> 
    <td width="1%"><img src="../security/images/spacer.gif" width="1" height="1"></td>
    <td width="63%" align="center" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="98%" height="2" align="left" valign="top" bgcolor="#FFD553"><img src="../images/y01.jpg" width="12" height="12"></td>
          <td align="right" valign="top" width="2%" height="2">&nbsp;&nbsp;</td>
        </tr>
        <tr> 
          <td width="98%" height="15" bgcolor="#FFD553" align="center"><img src="../images/okpic005.gif" width="300" height="250"></td>
          <td width="2%" align="right" valign="bottom" height="15"><img src="../security/images/spacer.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td width="98%" height="15" bgcolor="#FFD553">&nbsp;</td>
          <td width="2%" align="right" valign="bottom" height="15">&nbsp;</td>
        </tr>
        <tr> 
          <td width="98%" height="15" bgcolor="#FFD553"> 
<%if mid(Session("sAuthor"),10,3) = "AED" Or mid(Session("sAuthor"),10,3) = "VIW" Then%>
            <table width="100%" border="0" cellspacing="2" cellpadding="0">
              <tr> 
                <td width="1%"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
                <td bgcolor="#DB4B37" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Products</font></b></font></td>
                <td width="4%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="1%">&nbsp; </td>
                <td width="61%"> 
<!--                  <form method="post" action="../datamanager/vpproduct4.asp">
                    Search For 
                    <input type="text" name="SearchProduct" size="20" maxlength="20">
                    <input type="submit" name="Submit22" value="Go !!">
                    <br>
                    <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37"> 
                    </font> 
                  </form> -->
                  <form method="post" action="../datamanager/products.asp">
                    Search For 
                    <input type="text" name="SearchProduct" size="20" maxlength="20">
                    <input type="submit" name="Submit22" value="Go !!">
                    <br>
                     By Expense Type : 
                    <select name="ListProduct">
                      <option value="0" selected>&lt;-- Search List --&gt;</option>
                      <%Do While Not RS.EOF	%> 
                      <option value="<%=RS("sexptype")%>"><%=RS("sexptype")%>-<%=RS("sexpdesc")%></option>
                      <%
							RS.MoveNext
							Loop
							Rs.Close
							sSQL="Select * from pCattype Order By scattype"
 							Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
							%> 
                    </select>
                    <br>
                    Group &nbsp;&nbsp;: 
                    <select name="ListCattype">
                      <option value="0" selected>&lt;-- Search List --&gt;</option>
                      <%Do While Not RS.EOF	%> 
                      <option value="<%=RS("sCattype")%>"><%=RS("sCattype")%>-<%=RS("sCatdesc")%></option>
                      <%
							RS.MoveNext
							Loop
							Rs.Close
							sSQL = "Select * from pCatalog Order By scatID"
							sSQL = "select * from pCatalog where sCatalog like '%(" & year(date) & ")%' order by sCatID"
 							Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%> 
                    </select>
                    <br>
                    Sub Group &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: 
                    <select name="ListCatalog">
                      <option value="0" selected>&lt;-- Search List --&gt;</option>
                      <%Do While Not RS.EOF	%> 
                      <option value="<%=RS("sCatID")%>"><%=RS("sCatID")%>-<%=RS("sCatalog")%></option>
                      <%
							RS.MoveNext
							Loop
							Rs.Close
%> 
                    </select>
                    <br>
                    Sort By &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: 
                    <select name="Sortby">
                      <option value="0" selected>Product Code</option>
                      <option value="1">Expense Type</option>
                    </select>
                  </form>
<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37"><img src="../images/plus.gif" width="9" height="9"></font> 
                  <font color="#CC3333"><b><a href="../datamanager/vpproduct.asp">View 
                  Product</a></b></font>
                </td>
<!--                <td width="34%" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37"><img src="../images/plus.gif" width="9" height="9"></font> 
                  <font color="#CC3333"><b><a href="../datamanager/vpproduct.asp">View 
                  Product</a></b></font><font color="#FF3333"><b> </b></font></td>
                <td width="4%">&nbsp;</td> -->
              </tr>
            </table>
<%else%>
	<img src="../images/dot.gif">
<%End if%>
          </td>
          <td width="2%" align="right" valign="bottom" height="15">&nbsp;</td>
        </tr>
        <tr> 
          <td width="98%" height="2" bgcolor="#FFD553"> 
<%if mid(Session("sAuthor"),7,3) = "AED" Or mid(Session("sAuthor"),7,3) = "VIW" Then%>
            <table width="100%" border="0" cellspacing="2" cellpadding="0">
              <tr> 
                <td width="4%"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Vendors</font></b></font></td>
                <td width="4%">&nbsp;</td>
              </tr>
              <tr valign="top"> 
                <td width="4%">&nbsp; </td>
                <td width="92%"> 

                  <form method="post" action="../datamanager/pvender.asp">
                    Search For 
                    <input type="text" name="SearchVender" size="20" maxlength="20">
                    <input type="submit" name="Submit2" value="Go !!">
                    <br>
                  </form>
                </td>
                <td width="4%">&nbsp;</td>
              </tr>
            </table>
<%else%>
<img src="../images/dot.gif">
<%End if%>
          </td>
          <td width="2%" align="right" valign="bottom" height="2">&nbsp;</td>
        </tr>
        <tr> 
          <td width="98%" align="right" bgcolor="#FFD553" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
          <td width="2%" align="right" valign="bottom" height="15">&nbsp;</td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font></td>
    <td width="36%" align="center" valign="top">
<%if mid(Session("sAuthor"),4,3) = "AED" Or mid(Session("sAuthor"),4,3) = "VIW" Or Mid(Session("sAuthor"),34,3) = "AED" OR Mid(Session("sAuthor"),34,3) = "VIW" Then%> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#FFCC66">..</font></td>
          <td width="96%" align="center" bgcolor="#FFD553"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="3" COLOR="#666666"><B><FONT COLOR="#FFCC66">..</FONT></B></FONT><font face="Verdana, Arial, Helvetica, sans-serif" size="3"></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr> 
          <td width="2%" bgcolor="#FFD553">&nbsp;</td>
          <td width="96%" align="center" valign="top" bgcolor="#FFD553"> 
<%if mid(Session("sAuthor"),4,3) = "AED" Or mid(Session("sAuthor"),4,3) = "VIW" Then%>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Organization 
                  Data </font></b></font></td>
              </tr>
              <tr valign="top"> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37">
					<img src="../images/plus.gif" width="9" height="9"> 
                  <a href="pdivisionall.asp"><b>Division Data</b></a><br>
                  <img src="../images/plus.gif" width="9" height="9"> <a href="pdepartmentall.asp"><b>Department 
                  Data</b></a><br>
                  <img src="../images/plus.gif" width="9" height="9"> <a href="pSectionall.asp"><b>Section 
                  Data</b></a><br>
                  <img src="../images/plus.gif" width="9" height="9"> <a href="pPositionall.asp"><b>Position 
                  Data</b></a><br>
                  <img src="../images/plus.gif" width="9" height="9"> <a href="pStatusall.asp"><b>User Status 
                  Data</b></a> </font>
					</td>
              </tr>
            </table>
            <br>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#DB4B37"> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Approval 
                  Method </font></b></font></td>
              </tr>
              <tr valign="top"> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37">
					<img src="../images/plus.gif" width="9" height="9"> 
                  <b><a href="../datamanager/pflowstepall.asp">Approver Step code</a></b> 
                  <!--Approve Flow Start--><br>
                  <img src="../images/plus.gif" width="9" height="9"> <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><a href="../datamanager/pappflowall.asp">Org Code</a>
					</font></b></font> <!--Approve Flow Stop--><br>
                  <!-- Expense Type  Start--><img src="../images/plus.gif" width="9" height="9"> 
                  <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><a href="pexpstepall.asp">Exp Type Approval Step
					</a></font></b></font> <!-- Expense Type Stop--> 
                  <br>
                  <!--Approve Step Start--><img src="../images/plus.gif" width="9" height="9"> 
                  <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><a href="../datamanager/pappsteplist.asp"> Other Approval Step</a></font></b>
				<b> [ <font color="#FFFFFF" size="2"><a href="../datamanager/pappstepall.asp">All</a></font></b>
<b>                   ]</b></font></font> <!--Approve Step Stop--><br>
                  
                    <!-- Expense Type  Start-->
                    <img src="../images/plus.gif" width="9" height="9"> <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><a href="../datamanager/pexpenstypeall.asp">Expense 
                    Type</a> </font></b></font> <!-- Expense Type Stop--><br>
					</td>
              </tr>
            </table>
            <br>
<%End if%>
<%if mid(Session("sAuthor"),4,3) = "AED" Or mid(Session("sAuthor"),4,3) = "VIW" Or Mid(Session("sAuthor"),34,3) = "AED" OR Mid(Session("sAuthor"),34,3) = "VIW" Then%> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#DB4B37"> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Other Master Files</font></b></font></td>
              </tr>
              <tr valign="top"> 
                <td height="77">
<%End if%>
                  <p><% If Mid(Session("sAuthor"),34,3) = "AED" OR Mid(Session("sAuthor"),34,3) = "VIW" then%>
                    <!-- Catalog Type  Start--><img src="../images/plus.gif" width="9" height="9"> 
                    <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><a href="../datamanager/pCattypeall.asp">Group in Expense Type</a> </font></b></font> <!-- Catalog Type Stop--><br>
                    <!-- Catalog  Start--><img src="../images/plus.gif" width="9" height="9"> 
                    <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><a href="../datamanager/pCataloglist.asp">Sub-Group in Expense Type</a>
					<font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"> [ </font>
					<a href="../datamanager/pCatalogall.asp">All</a>
					<font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"> ] </font>
					 </font></b></font> <!-- Catalog Stop--><br>
					<%End if%>
<%if mid(Session("sAuthor"),4,3) = "AED" Or mid(Session("sAuthor"),4,3) = "VIW" Then%>
                    <br>
                    <!--Currency Start--><img src="../images/plus.gif" width="9" height="9"> 
                    <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><a href="../datamanager/pcurrencyall.asp">Currency</a></font></b></font> 
                    <!--Currency Stop-->  <br>
                    <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37">
					<!--Currency Start--><img src="../images/plus.gif" width="9" height="9"> 
                    <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><a href="../datamanager/pUnitall.asp">Unit</a></font></b></font><!--Currency Stop--> 
                    </font> <br>
					<!--Term code Start--><img src="../images/plus.gif" width="9" height="9"> 
                    <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><a href="../datamanager/ptermcodeall.asp">Term Code</a></font></b></font><!--Term Code Stop--> 
                    </font> 
                </td>
              </tr>
<%End if%>
            </table>
            <br>
            <font color="#006699"> </font></td>
          <td width="2%" align="right" valign="top" bgcolor="#FFD553">&nbsp;</td>
        </tr>
        <tr> 
          <td width="2%" align="left" valign="bottom" bgcolor="#FFD553"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom" bgcolor="#FFD553"><img src="../images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
<%End if%>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><br>
      </b></font></td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td width="1%"> <img src="images/spacer.gif" width="10" height="10"><br>
    </td>
    <td width="99%">
            <!--#include file="../inc/copyRight.inc"-->
  </tr>
</table>
</body>
</html>
