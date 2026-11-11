<%

 ' ===================================================='
 ' 		AbleOn.com Intranet Solution
 ' 		Author: Phornthep Khongsathian
 '		Version: 1.01
 ' ===================================================='
 ' This is the Connect to ORACLE7 database
 ' Set System DNS to connect ORACLE USER=ableon/dotcom;Server=2:
 ' Name of DNS is "aintra"

	dim  conn

	set conn= Server.CreateObject("ADODB.Connection")
	'conn.Open  "oracle73"
	conn.Open  "dsn=amec400;uid=INV00;pwd=INV00"
%>
