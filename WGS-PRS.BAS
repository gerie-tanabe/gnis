10 '            *******************************************************
20 '            **  THIS PROGRAM CONVERTS WGS TO PRS AND VISE VERSA  **
30 '            **        <<<< develop by Jiggy de Jesus >>>>        **
40 '            **                 13 December 1994                  **
50 '            *******************************************************
60 CLS
70 LOCATE 8: COLOR 14:PRINT "           *******************************************************                         **                                                   **                         **  PLEASE SWITCH 'ON' AND 'ON LINE' THE PRINTER!!!  **
80 LOCATE 11: COLOR 14:PRINT "           **                                                   **                         *******************************************************"
90 LOCATE 20,25: COLOR 13:PRINT " PRESS ANY KEY TO CONTINUE "
100 W$=INPUT$(1)
110 IF W$="" THEN 120
120 CLS:LOCATE 10,25: COLOR 14:PRINT "                           **********************                                                          **                  **                                                          **  ENTER PASSWORD  **"
130 LOCATE 13,25: COLOR 14:PRINT "                           **                  **                                                          **********************"
140 Q$=INPUT$(4)
150 IF Q$="boss" THEN 270 ELSE 160
160 BEEP : BEEP : BEEP
170 LOCATE 10,25 : COLOR 18:PRINT "                           **********************                                                          **                  **                                                          **  ACCESS  DENIED  **"
180 LOCATE 13,25 : COLOR 18:PRINT "                           **                  **                                                          **********************"
190 BEEP : BEEP : BEEP
200 LOCATE 10,25 : COLOR 20:PRINT "                           **********************                                                          **                  **                                                          **  ACCESS  DENIED  **
210 LOCATE 13,25 : COLOR 20:PRINT "                           **                  **                                                          **********************"
220 BEEP : BEEP : BEEP
230 LOCATE 10,25 : COLOR 19:PRINT "                           **********************                                                          **                  **                                                          **  ACCESS  DENIED  **"
240 LOCATE 13,25 : COLOR 19:PRINT "                           **                  **                                                          **********************"
250 BEEP : BEEP : BEEP
260 GOTO 120
270 CLEAR : CLS
280 DEFDBL A-H,J-Z
290 PI=3.141592653999999:O=PI/180
300 CLS:LOCATE 10,20:COLOR 14:PRINT"              TYPE  <1> CONVERSION FROM WGS 84 TO PRS 92                                                                                                                            <2> CONVERSION FROM PRS 92 TO WGS
310 A$=INPUT$(1):PRINT A$
320 IF A$="1" THEN 400
330 IF A$="2" THEN 340 ELSE 300

340 DX=-1.276219999999999E+02:DY=-6.724499999999999E+01:DZ=-4.704300000000000E+01:S=-1.060019999999999
350 RX=3.068000/3600:RY=-4.902999/3600:RZ=-1.577999/3600
360 A=6.378206400000000E+06:E2=6.768657999999999E-03
370 A2=6.378137E+06:E22=6.694379999999999E-03
380 CLS:LOCATE 5,25:COLOR 11:PRINT"* CONVERSION FROM PRS 92 TO WGS 84 *"
390 GOTO 450

400 DX=1.276219999999999E+02:DY=6.724499999999999E+01:DZ=4.704300000000000E+01:S=1.060019999999999
410 RX=-3.068000/3600:RY=4.902999/3600:RZ=1.577999/3600
420 A=6.378137E+06:E2=6.694379999999999E-03
430 A2=6.378206400000000E+06:E22=6.768657999999999E-03
440 CLS:LOCATE 5,25:COLOR 11:PRINT"* CONVERSION FROM WGS 84 TO PRS 92 *"

450 INPUT "NO. OF STATIONS TO BE CONVERTED=";I
460 FOR I=1 TO I
470 INPUT "STATION NAME=";Q$(I)
480 W(I)=I
490 INPUT "LATITUDE=";DEG,MIN,SEC:D1(I)=DEG:M1(I)=MIN:S1(I)=SEC:GOSUB 1060:P1=DERAD
500 INPUT"LONGITUDE=";DEG,MIN,SEC:D2(I)=DEG:M2(I)=MIN:S2(I)=SEC:GOSUB 1060:L1=DERAD
510 INPUT "ELLIPSOIDAL HT.=";H1:H1(I)=H1
520 N1=A/SQR(1-E2*(SIN(P1))^2)
530 X1=(N1+H1)*COS(P1)*COS(L1)
540 Y1=(N1+H1)*COS(P1)*SIN(L1)
550 Z1=(N1*(1-E2)+H1)*SIN(P1)
560 X2=DX+(1+S*10^-6)*(X1+RZ*O*Y1-RY*O*Z1)
570 Y2=DY+(1+S*10^-6)*(-RZ*O*X1+Y1+RX*O*Z1)
580 Z2=DZ+(1+S*10^-6)*(RY*O*X1-RX*O*Y1+Z1)
590 R=SQR(X2^2+Y2^2)
600 H2=0
610 P2=0
620 P=P2
630 N2=A2/SQR(1-E22*(SIN(P2))^2)
640 P2=ATN((N2+H2)*Z2/((N2*(1-E22)+H2)*R))
650 H2=R/COS(P2)-N2
660 K=ABS(P-P2)

670 IF K>9.999999E-09 THEN 620 ELSE 680
680 A1=P2*180/PI:H2(I)=H2

690 PRINT:PRINT:COLOR 10:PRINT "LATITUDE 2 =";
700 GOSUB 1080:D3(I)=DG:M3(I)=MN:S3(I)=SC
710 B2=(ATN(Y2/X2))*180/PI
720 L2=B2+180
730 A1=L2
740 COLOR 10:PRINT"LONGITUDE 2=";
750 GOSUB 1080:D4(I)=DG:M4(I)=MN:S4(I)=SC
760 COLOR 10:PRINT"HEIGHT 2=";H2(I)
770 LOCATE 20,25:COLOR 11:PRINT "PRESS ANY KEY TO CONTINUE"
780 B$=INPUT$(1)
790 IF B$=" " THEN 800
800 IF A$="1" THEN 810 ELSE 820
810 CLS:LOCATE 5,25:COLOR 11:PRINT"* CONVERSION FROM WGS 84 TO PRS 92 *":GOTO 830
820 CLS:LOCATE 5,25:COLOR 11:PRINT"* CONVERSION FROM PRS 92 TO WGS 84*"
830 NEXT I
840 IF A$="1" THEN 880 ELSE 850
850 LPRINT:LPRINT:LPRINT:LPRINT:LPRINT TAB(19)"<<<<<<CONVERSION FROM PRS 92 TO WGS 84>>>>>>":LPRINT:LPRINT
860 LPRINT TAB(29)"**PRS 92 COORDINATES**":LPRINT
870 GOTO 900
880 LPRINT:LPRINT:LPRINT:LPRINT:LPRINT TAB(19)"<<<<<<CONVERSION FROM WGS 84 TO PRS 92>>>>>>":LPRINT:LPRINT
890 LPRINT TAB(29)"**WGS 84 COORDINATES**":LPRINT
900 LPRINT "STATION NAME";TAB(24);"LATITUDE";TAB(44);"LONGITUDE";TAB(63);"HEIGHT"
910 FOR I=1 TO I-1
920 LPRINT W(I);Q$(I);TAB(20);USING"##  ##  ##.#####    ###  ##  ##.#####    ####.####";D1(I);M1(I);S1(I);D2(I);M2(I);S2(I);H1(I)
930 NEXT I
940 LPRINT:LPRINT:LPRINT
950 IF A$="1"THEN 970 ELSE 960
960 LPRINT TAB(29)"**WGS 84 COORDINATES**":LPRINT:GOTO 980
970 LPRINT TAB(29)"**PRS 92 COORDINATES**":LPRINT
980 LPRINT "STATION NAME";TAB(24);"LATITUDE";TAB(44);"LONGITUDE";TAB(63);"HEIGHT"
990 FOR I=1 TO I-1
1000 LPRINT W(I);Q$(I);TAB(20);USING"##  ##  ##.#####    ###  ##  ##.#####    ####.####";D3(I);M3(I);S3(I);D4(I);M4(I);S4(I);H2(I)
1010 NEXT I
1020 CLS:LOCATE 13,20:COLOR 14:PRINT "TYPE :<1> TO CONTINUE  <2> EXIT TO DOS :";
1030 C$=INPUT$(1):PRINT C$
1040 IF C$="1" THEN 300
1050 IF C$="2" THEN SYSTEM ELSE 1020
1060 DERAD=(DEG+MIN/60+SEC/3600)*PI/180
1070 RETURN
1080 DG=INT(A1)
1090 MN1=(A1-DG)*60
1100 MN=INT(MN1)
1110 SC=(MN1-MN)*60
1120 PRINT DG;MN;SC
1130 RETURN
