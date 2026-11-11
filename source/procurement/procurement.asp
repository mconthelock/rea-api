<%	
  Dim sAuthorise
	sAuthorise = Session("sAuthor")
	if Trim(sAuthorise)	= "" Then
		Response.Redirect "default.htm"
	End If
%>

<html>
<head>
<title>MITSUBISHI ELEVATOR ASIA</title>
</head>
<frameset rows="60,*" frameborder="NO" border="0" framespacing="0" cols="*"> 
  <frame name="topFrame" scrolling="NO" noresize src="head.asp" >
  <frameset cols="125,*" frameborder="NO" border="0" framespacing="0" rows="*"> 
    <frame name="leftFrame" scrolling="NO" noresize src="menubar.asp">
    <frame src="phome.asp"name="num"scrolling="auto">
  </frameset>
</frameset>
<noframes><body bgcolor="#FFFFFF">

</body></noframes>
</html>
