<%

 ' ===================================================='
 ' 		AbleOn.com Intranet Solution
 ' 		Author: Phornthep Khongsathian
 '		Version: 1.01
 ' ===================================================='
 ' This is the Connect to ORACLE7 database
 ' Set System DNS to connect ORACLE USER=ableon/dotcom;Server=2:
 ' Name of DNS is "aintra"

	dim  connas400, connas4002

	set connas400= Server.CreateObject("ADODB.Connection")
	set connas4002= Server.CreateObject("ADODB.Connection")
	'conn.Open  "oracle73"
'	connas400.Open "DSN=as400;uid=PCS;pwd=PCS"
'	connas4002.Open "DSN=as400;uid=PCS;pwd=PCS"

	
	connas400.Open "DSN=as400;uid=U02035;pwd=k14092523"
	connas4002.Open "DSN=as400;uid=U02035;pwd=k14092523"
%>

<%

'Dim sConStr 
'sConStr = "Driver=Client Access ODBC Driver (32-bit);"
'sConStr = sConStr & "System=172.20.255.243;UID=U02035;PWD=kali;"
'connas400.open sConStr
'connas4002.open sConStr
 'cm.ActiveConnection = c


'connas4002.Open "Driver={Client Access ODBC Driver (32-bit)};" & _
'           "System=172.20.255.243;" & _
'           "Uid=U02035;" & _
'           "Pwd=kali"

'connas400.Open "Driver={Client Access ODBC Driver (32-bit)};" & _
'           "System=172.20.255.243;" & _
'           "Uid=U02035;" & _
'           "Pwd=kali"

%>
