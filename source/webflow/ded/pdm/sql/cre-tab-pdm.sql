create table pdmAuth
(
	cNo char(2), --running number
	vName varchar2(30),
	primary key (cNo)
);


create table pdmForm
(
	nFrmNo number(4,0),
	vOrgNo varchar2(5),
	cYear char(2),
	cYear2 char(4),
	nRunNo number(6, 0),
	cNo char(2),  --Authorize
	cReq char(1), --Application request method Add, Update, Delete
	vPurpose varchar2(255),
	primary key (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, cNo),
	foreign key (cNo) references pdmAuth(cNo),
	foreign key (nFrmNo, vOrgNo, cYear, cYear2, nRunNo) references form(nFrmNo, vOrgNo, cYear, cYear2, nRunNo) 
);