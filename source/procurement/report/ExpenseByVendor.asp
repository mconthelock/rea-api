<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: worakrit yodthonglert
	' Version: 1.01
%>
<%
	Dim sExptype
	Dim i
%>
<HTML>
<HEAD>
<TITLE>Expense By Vendor</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		
<Form Action="Sum5.asp" method="post">
  <table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="top" width="1%" height="2"><img src="../images/ccb.gif" width="10" height="10"></td>
      <td width="99%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="8" height="8"></td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="99%" height="30" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="+1" color="#666666"><b><b><b><b><b><b><b><font color="#BF00BF">Report 
        Manager</font></b></b></b></b></b></b></b></font></td>
    </tr>
    <tr> 
      <td width="1%"><img src="../images/spacer.gif" width="1" height="1"></td>
      <td width="99%" align="center" valign="top" height="0"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="98%" height="2" align="left" valign="top" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><img src="../images/y01.jpg" width="12" height="12"></b></font></td>
          </tr>
          <tr> 
            <td width="98%" align="left" valign="top" height="0"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr align="left" valign="top" bgcolor="#FFD553"> 
                  <td colspan="2" align="center">&nbsp;</td>
                </tr>
                <tr align="center" valign="top"> 
                  <td colspan="2" bgcolor="#FFD553"> 
                    <table border=0 width=400 align="Center" cellspacing="1" cellpadding="0">
                      <tr bgcolor="#CC00CC"> 
                        <td colspan="2" height="20"> 
                          <div align="center"><b><font size="2"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Expense 
                            By Vendor</font></font></b></div>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="30%" height="10"> 
                          <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                            Vender Code &nbsp;</font></div>
                        </td>
                        <td width="70%" height="10"> 
                          <select name="VendCode" size=1>
                            <% List_Pvender(Session("sexptype")) %> 
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="30%" height="10"> 
                          <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">From 
                            &nbsp;</font></div>
                        </td>
                        <td width="70%" height="10"> 
                          <select name="Form_day">
                            <%Create_daylist(day(date)) %> 
                            <option value= 'All' Selected>All</option>
                          </select>
                          <select name="Form_month">
                            <% Create_MonthList(month(date))%> 
                            <option value= 'All' Selected>All</option>
                          </select>
                          <select name="Form_Year">
                            <% Create_YearList(Year(date))%> 
                            <option value= 'All' Selected>All</option>
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="30%" height="10"> 
                          <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">To 
                            &nbsp;</font></div>
                        </td>
                        <td width="70%" height="10"> 
                          <select name="To_Day">
                            <%Create_daylist(day(date-7)) %> 
                            <option value= 'All' Selected>All</option>
                          </select>
	                         <select name="To_Month">
                            <% Create_MonthList(month(date-7))%> 
                            <option value= 'All' Selected>All</option>
                          </select>
                          <select name="To_Year">
                            <% Create_YearList(Year(date))%> 
                            <option value= 'All' Selected>All</option>
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="30%" height="10"> 
                          <div align="right"></div>
                        </td>
                        <td width="70%" height="10"> 
                          <input type="submit" name="Submit" value="Submit">
                          <input type="reset" name="Reset" value="Reset">
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr align="left" valign="top"> 
                  <td width="98%" bgcolor="#FFD553">&nbsp;</td>
                  <td width="2%" align="right" bgcolor="#FFD553" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font></td>
    </tr>
  </table>
  <table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr align="center"> 
      <td width="1%"> <img src="../report/images/spacer.gif" width="1" height="1"><br>
      </td>
      <td width="99%"> 
         <!--#include file="../inc/copyRight.inc"--></td>
    </tr>
  </table>
  </Form>
</BODY>
</HTML>
<%
Sub Create_Yearlist(searchyear)
Dim syear
	For i=2000 to searchyear
		Response.Write "<option value= '"& i & "'>"& i &"</option>"& vbCrLf
	Next
End Sub
%>