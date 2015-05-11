<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> WGS to PRS92 </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="JunReyes">
<META NAME="Keywords" CONTENT="WGS84, PRS92">
<META NAME="Description" CONTENT="Conversion">
</HEAD>
<STYLE TYPE="text/css">
<!--  
BODY
   {
   color:black;
   font-family:sans-serif;
   }
A:link{color:white}
A:visited{color:yellow}
-->
</STYLE>

<SCRIPT TYPE="text/javascript">
<!--
function WGSPRS92(GeoForm)
{
var theLatDeg = GeoForm.LatDeg.value/1; 
var theLatMin = GeoForm.LatMin.value/60;
var theLatSec = GeoForm.LatSec.value/3600;
var theLongDeg = GeoForm.LongDeg.value/1; 
var theLongMin = GeoForm.LongMin.value/60;
var theLongSec = GeoForm.LongSec.value/3600;
var H1 = GeoForm.EllipHt.value;
var PI = 4 * Math.atan(1);  // 3.141592653999999;     <!-- background-color:yellow; -->
var O = PI/180;
var H2, DG, MN1, MN, SC, A1, K;
var theStrg;
var chk;

DX = 1.276219999999999E+02;
DY = 6.724499999999999E+01;
DZ = 4.704300000000000E+01;
S  = 1.060019999999999;
RX = -3.068/3600;
RY = 4.902999/3600;
RZ = 1.577999/3600;
A  = 6.378137E+06;
E2 = 6.694379999999999E-03;
A2 = 6.378206400000000E+06;
E22 = 6.768657999999999E-03;

if ((theLatDeg >= 0 && theLatMin >= 0 && theLatSec >= 0 && H1 >= 0))
	{

	H2 = 0;
	P2 = 0;
	P  = P2;
	A1 = 0;
	K  = 0;
	P1 = (theLatDeg + theLatMin + theLatSec)*(PI/180);
	L1 = (theLongDeg + theLongMin + theLongSec)*(PI/180);
	N1 = A / Math.sqrt(1 - E2 * Math.pow(Math.sin(P1),2));
	X1 = ((N1/1)+(H1/1)) * Math.cos(P1) * Math.cos(L1);
	Y1 = ((N1/1)+(H1/1)) * Math.cos(P1) * Math.sin(L1);
	Z1 = ((N1/1)*(1-E2)+(H1/1))* Math.sin(P1);
	X2 = DX+(1+S* Math.pow(10,-6) )*(X1+RZ*O*Y1-RY*O*Z1);
	Y2 = DY+(1+S* Math.pow(10,-6) )*(-RZ*O*X1+Y1+RX*O*Z1);
	Z2 = DZ+(1+S* Math.pow(10,-6) )*(RY*O*X1-RX*O*Y1+Z1);
	R  = Math.sqrt( Math.pow(X2,2)+Math.pow(Y2,2));
	P = P2
	N2 = A2/ Math.sqrt(1-E22*Math.pow(Math.sin(P2),2));
	P2 = Math.atan((N2+H2)*Z2/((N2*(1-E22)+H2)*R));
	H2 = R/Math.cos(P2)-N2;
	K  = Math.abs(P-P2);
	chk = false;
	do
	{
		{
		P = P2;
		N2 = A2/ Math.sqrt(1-E22*Math.pow(Math.sin(P2),2));
		P2 = Math.atan((N2+H2)*Z2/((N2*(1-E22)+H2)*R));
		H2 = R/Math.cos(P2)-N2;
		//K  = Math.abs(P-P2);
		K = String(P-P2).substr(0,2);
		}
		if (K > 9.999999E-09)
			{
			P = P2;
			N2 = A2/ Math.sqrt(1-E22*Math.pow(Math.sin(P2),2));
			P2 = Math.atan((N2+H2)*Z2/((N2*(1-E22)+H2)*R));
			H2 = R/Math.cos(P2)-N2;
			//K  = Math.abs(P-P2);
			K = String(P-P2).substr(0,2);
			A1 = P2*180/PI;
			chk = true;
			}
	}
	while (chk = false);	//K < 9.999999E-09);

	//Latitude
	theStrg = String(A1).substr(0,2);
	DG = theStrg; 
	MN1 = (A1-DG)*60;
	MN = String(MN1).substr(0,2);
	SC= (MN1-MN)*60;
	
	GeoForm.PRS92_Latitude.value = DG+String.fromCharCode(176)+" "+MN+"' "+SC+String.fromCharCode(34);

	//Longitude
	B2 = (Math.atan(Y2/X2))*180/PI;
	L2 = B2+180;
	A1 = L2;
	theStrg = String(A1).substr(0,3);
	DG = theStrg; 
	MN1 = (A1-DG)*60;
	MN = String(MN1).substr(0,2);
	SC = (MN1-MN)*60;

	GeoForm.PRS92_Longitude.value = DG+String.fromCharCode(176)+" "+MN+"' "+SC+String.fromCharCode(34);
	}
}

function PRS92WGS(GeoForm)
{
var thewLatDeg = GeoForm.wLatDeg.value/1; 
var thewLatMin = GeoForm.wLatMin.value/60;
var thewLatSec = GeoForm.wLatSec.value/3600;
var thewLongDeg = GeoForm.wLongDeg.value/1; 
var thewLongMin = GeoForm.wLongMin.value/60;
var thewLongSec = GeoForm.wLongSec.value/3600;
var wH1 = GeoForm.wEllipHt.value;
var PI = 4 * Math.atan(1);  // 3.141592653999999;
var O = PI/180;
var H2, DG, MN1, MN, SC, A1, K;
var theStrg;
var chk;

DX = -1.276219999999999E+02;
DY = -6.724499999999999E+01;
DZ = -4.704300000000000E+01;
S  = -1.060019999999999;
RX = 3.068/3600;
RY = -4.902999/3600;
RZ = -1.577999/3600;
A  = 6.3782064+06;
E2 = 6.768657999999999E-03;
A2 = 6.378137E+06;
E22 = 6.694379999999999E-03;

if ((thewLatDeg >= 0 && thewLatMin >= 0 && thewLatSec >= 0 && wH1 >= 0))
	{
	H2 = 0;
	P2 = 0;
	P  = P2;
	A1 = 0;
	K  = 0;
	P1 = (thewLatDeg + thewLatMin + thewLatSec)*(PI/180);
	L1 = (thewLongDeg + thewLongMin + thewLongSec)*(PI/180);
	N1 = A / Math.sqrt(1 - E2 * Math.pow(Math.sin(P1),2));
	X1 = ((N1/1)+(wH1/1)) * Math.cos(P1) * Math.cos(L1);
	Y1 = ((N1/1)+(wH1/1)) * Math.cos(P1) * Math.sin(L1);
	Z1 = ((N1/1)*(1-E2)+(wH1/1))* Math.sin(P1);
	X2 = DX+(1+S* Math.pow(10,-6) )*(X1+RZ*O*Y1-RY*O*Z1);
	Y2 = DY+(1+S* Math.pow(10,-6) )*(-RZ*O*X1+Y1+RX*O*Z1);
	Z2 = DZ+(1+S* Math.pow(10,-6) )*(RY*O*X1-RX*O*Y1+Z1);
	R  = Math.sqrt( Math.pow(X2,2)+Math.pow(Y2,2));
//	P = P2;
//	N2 = A2/ Math.sqrt(1-E22*Math.pow(Math.sin(P2),2));
//	P2 = Math.atan((N2+H2)*Z2/((N2*(1-E22)+H2)*R));
//	H2 = R/Math.cos(P2)-N2;
//	K  = Math.abs(P-P2);
	chk = false;
	do
	{
		{
		P = P2;
		N2 = A2/ Math.sqrt(1-E22*Math.pow(Math.sin(P2),2));
		P2 = Math.atan((N2+H2)*Z2/((N2*(1-E22)+H2)*R));
		H2 = R/Math.cos(P2)-N2;
		K  = Math.abs(P-P2);
		//K = String(P-P2).substr(0,2);
		alert(K);
		}
		if (K > 9.999999E-09)
			{
			P = P2;
			N2 = A2/ Math.sqrt(1-E22*Math.pow(Math.sin(P2),2));
			P2 = Math.atan((N2+H2)*Z2/((N2*(1-E22)+H2)*R));
			H2 = R/Math.cos(P2)-N2;
			//**K  = Math.abs(P-P2);
			K = String(P-P2).substr(0,2);
			A1 = P2*180/PI;
			chk = true;
			alert(P2);
			}
	}
	while (chk = false);	//K < 9.999999E-09);

	//Latitude
	theStrg = String(A1).substr(0,2);
	DG = theStrg; 
	MN1 = (A1-DG)*60;
	MN = String(MN1).substr(0,2);
	SC= (MN1-MN)*60;

	GeoForm.WGS84.value = DG+String.fromCharCode(176)+" "+MN+"' "+SC+String.fromCharCode(34);

	}
}
//-->
</SCRIPT>

<BODY>
<FONT NAME="Arial" SIZE="" COLOR="">
<FORM>
<TABLE border="1" width="100%">
<TR>
	<TD bgcolor="#66CC33" ><H3><BR><CENTER>WGS to PRS92</CENTER></H3></TD>
</TR>
<TR><TD>
<%
Response.write "<CENTER><TABLE>"
Response.write "<TR>"
Response.write "	<TD colspan=""2""><FONT FACE=""Geneva, Arial"" SIZE=""2"">"
Response.write "		Please input the Latitude and Longitude"
Response.write "    </FONT></TD>"
Response.write "</TR>"
Response.write "<TR>"
Response.write "	<TD align=""right""><FONT FACE=""Geneva, Arial"" SIZE=""2"">Latitude&nbsp;&nbsp;</FONT></TD>"
Response.write "	<TD>"
Response.write "		<INPUT TYPE=""text"" NAME=""LatDeg"" size=""5"" value=""17""> "
Response.write "		<INPUT TYPE=""text"" NAME=""LatMin"" size=""5"" value=""36""> "
Response.write "		<INPUT TYPE=""text"" NAME=""LatSec"" size=""10"" value=""16.84045"">"
Response.write "    </TD>"
Response.write "</TR>"
Response.write "<TR>"
Response.write "	<TD align=""right""><FONT FACE=""Geneva, Arial"" SIZE=""2"">Longitude&nbsp;&nbsp;</FONT></TD>"
Response.write "	<TD>"
Response.write "		<INPUT TYPE=""text"" NAME=""LongDeg"" size=""5"" value=""120""> "
Response.write "		<INPUT TYPE=""text"" NAME=""LongMin"" size=""5"" value=""36""> "
Response.write "		<INPUT TYPE=""text"" NAME=""LongSec"" size=""10"" value=""49.89478"">"
Response.write "    </TD>"
Response.write "</TR>"
Response.write "<TR>"
Response.write "	<TD align=""right""><FONT FACE=""Geneva, Arial"" SIZE=""2"">Ellipsoidal Height&nbsp;&nbsp;</FONT></TD>"
Response.write "	<TD>"
Response.write "		<INPUT TYPE=""text"" NAME=""EllipHt"" size=""10"" value=""56.948""> "
Response.write "    </TD>"
Response.write "</TR>"
Response.write "<TR>"
Response.write "	<TD align=""center"">"
Response.write "		<INPUT TYPE=BUTTON OnClick=""WGSPRS92(this.form);"" VALUE=""Convert to PRS92""> "
Response.write "    </TD>"
Response.write "	<TD align=""center"">"
Response.write "		<FONT FACE=""Geneva, Arial"" SIZE=""2"">Latitude:</FONT>&nbsp;&nbsp; <INPUT NAME=""PRS92_Latitude"" SIZE=""27""> "
Response.write "    </TD>"
Response.write "</TR>"
Response.write "<TR>"
Response.write "	<TD align=""center"">"
Response.write "		&nbsp;"
Response.write "    </TD>"
Response.write "	<TD align=""center"">"
Response.write "		<FONT FACE=""Geneva, Arial"" SIZE=""2"">Longitude:</FONT> <INPUT NAME=""PRS92_Longitude"" SIZE=""27""> "
Response.write "    </TD>"
Response.write "</TR>"
Response.write "</TABLE></CENTER>"
%>
</TD></TR>
</TABLE><BR>
<TABLE border="1" width="100%">
<TR>
	<TD bgcolor="#66CC33"><H3><BR><CENTER>PRS92 to WGS84</CENTER></H3></TD>
</TR>
<TR><TD>
<FORM>
<%
Response.write "<CENTER><TABLE>"
Response.write "<TR>"
Response.write "	<TD colspan=""2""><FONT FACE=""Geneva, Arial"" SIZE=""2"">"
Response.write "		Please input the Latitude and Longitude"
Response.write "    </FONT></TD>"
Response.write "</TR>"
Response.write "<TR>"
Response.write "	<TD align=""right""><FONT FACE=""Geneva, Arial"" SIZE=""2"">Latitude&nbsp;&nbsp;</FONT></TD>"
Response.write "	<TD>"
Response.write "		<INPUT TYPE=""text"" NAME=""wLatDeg"" size=""5"" value=""17""> "
Response.write "		<INPUT TYPE=""text"" NAME=""wLatMin"" size=""5"" value=""36""> "
Response.write "		<INPUT TYPE=""text"" NAME=""wLatSec"" size=""10"" value=""22.95611"">"
Response.write "    </TD>"
Response.write "</TR>"
Response.write "<TR>"
Response.write "	<TD align=""right""><FONT FACE=""Geneva, Arial"" SIZE=""2"">Longitude&nbsp;&nbsp;</FONT></TD>"
Response.write "	<TD>"
Response.write "		<INPUT TYPE=""text"" NAME=""wLongDeg"" size=""5"" value=""120""> "
Response.write "		<INPUT TYPE=""text"" NAME=""wLongMin"" size=""5"" value=""36""> "
Response.write "		<INPUT TYPE=""text"" NAME=""wLongSec"" size=""10"" value=""45.26102"">"
Response.write "    </TD>"
Response.write "</TR>"
Response.write "<TR>"
Response.write "	<TD align=""right""><FONT FACE=""Geneva, Arial"" SIZE=""2"">Ellipsoidal Height&nbsp;&nbsp;</FONT></TD>"
Response.write "	<TD>"
Response.write "		<INPUT TYPE=""text"" NAME=""wEllipHt"" size=""10"" value=""24.08315047""> "
Response.write "    </TD>"
Response.write "</TR>"
Response.write "<TR>"
Response.write "	<TD align=""center"">"
Response.write "		<INPUT TYPE=BUTTON OnClick=""PRS92WGS(this.form);"" VALUE=""Convert to WGS84""> "
Response.write "    </TD>"
Response.write "	<TD align=""center"">"
Response.write "		<INPUT NAME=""WGS84"" SIZE=""27""> "
Response.write "    </TD>"
Response.write "</TR>"
Response.write "</TABLE></CENTER>"
%>
</FORM>
</TD></TR>
</TABLE></FONT>
</BODY>
</HTML>
