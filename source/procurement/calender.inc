<%
Dim m(12, 1), d(31, 1), y(10, 1)

'Month name
m(0, 1) = "[Month]"
m(1, 1) = "JAN"
m(2, 1) = "FEB"
m(3, 1) = "MAR"
m(4, 1) = "APR"
m(5, 1) = "MAY"
m(6, 1) = "JUN"
m(7, 1) = "JUL"
m(8, 1) = "AUG"
m(9, 1) = "SEP"
m(10, 1) = "OCT"
m(11, 1) = "NOV"
m(12, 1) = "DEC"

'Month value
for i = 0 to uBound(m)
	m(i, 0) = i
'	response.write m(i, 0) & " " & m(i, 1) & "<br>"
next

'Date value and text
d(0, 0) = 32
d(0, 1) = "[Date]"
for i = 1 to uBound(d)
	d(i, 0) = i
	d(i, 1) = d(i, 0)
'	response.write d(i, 0) & " " & d(i, 1) & "<br>"
next

'Year value and text
y(0, 0) = 0
y(0, 1) = "[Year]"
for i = 1 to 10
	y(i, 0) = year(date) - i + 1
	y(i, 1) = y(i, 0)
'	response.write y(i, 0) & " " & y(i, 1) & "<br>"
next
%>