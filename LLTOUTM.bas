Attribute VB_Name = "Module2"

Public Type UTM
    
    Northing As Double
    Easting As Double
    Zone As Integer

End Type



Dim MyUTM As UTM


    Const fe = 500000 'Vars for the utmToLatlon
    Const ok = 0.9996
    Const PI = 3.14159265358979
    Const deg2rad = PI / 180
    Const rad2deg = 1 / deg2rad
    Const MajorAxis = 6378206.4
    Const Flattening = 1 / 294.978698214




Public Function LatLonToUTM(ByVal lat As Double, ByVal Lon As Double) As UTM
        On Error GoTo ErrHandler
        
        Dim a As Double
        Dim f As Double
        
        a = MajorAxis
        f = Flattening
      
        
        Dim recf As Double
        Dim b As Double
        Dim eSquared As Double
        Dim e2Squared As Double
        Dim tn As Double
        Dim ap As Double
        Dim bp As Double
        Dim cp As Double
        Dim dp As Double
        Dim ep As Double
        Dim olam As Double
        Dim dlam As Double
        Dim s As Double
        Dim c As Double
        Dim t As Double
        Dim eta As Double
        Dim sn As Double
        Dim tmd As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim t3 As Double
        Dim t6 As Double
        Dim t7 As Double
        Dim nfn As Double
        Dim Easting, Northing, utmXZone, utmYzone As Double
        If (Lon <= 0) Then
            utmXZone = 30 + Fix(Lon / 6)
        Else
            utmXZone = 31 + Fix(Lon / 6)
        End If
'        utmYzone = FutmYzone(lat)
        Dim latRad As Double
        latRad = lat * deg2rad
        Dim lonRad As Double
        lonRad = Lon * deg2rad
        'recf = 1 / f
        recf = 294.9786982
        b = a * (recf - 1#) / recf ' b is the Semiminor axis
        eSquared = CalculateESquared(a, b)
        e2Squared = CalculateE2Squared(a, b)
        tn = (a - b) / (a + b)
        ap = a * (1# - tn + 5# * ((tn * tn) - (tn * tn * tn)) / 4# + 81# * _
            ((tn * tn * tn * tn) - (tn * tn * tn * tn * tn)) / 64#)
        bp = 3# * a * (tn - (tn * tn) + 7# * ((tn * tn * tn) _
            - (tn * tn * tn * tn)) / 8# + 55# * (tn * tn * tn * tn * tn) / 64#) _
            / 2#
        cp = 15# * a * ((tn * tn) - (tn * tn * tn) + 3# * ((tn * tn * tn * tn) _
            - (tn * tn * tn * tn * tn)) / 4#) / 16#
        dp = 35# * a * ((tn * tn * tn) - (tn * tn * tn * tn) + 11# _
            * (tn * tn * tn * tn * tn) / 16#) / 48#
        ep = 315# * a * ((tn * tn * tn * tn) - (tn * tn * tn * tn * tn)) / 512#
        olam = (utmXZone * 6 - 183) * deg2rad
        dlam = lonRad - olam
        s = Sin(latRad)
        c = Cos(latRad)
        t = s / c
        eta = e2Squared * (c * c)
        sn = sphsn(a, eSquared, latRad)
        tmd = sphtmd(ap, bp, cp, dp, ep, latRad)
        t1 = tmd * ok
        t2 = sn * s * c * ok / 2#
        t3 = sn * s * (c * c * c) * ok * (5# - (t * t) + 9# * eta + 4# _
            * (eta * eta)) / 24#
        If (latRad < 0#) Then nfn = 10000000# Else nfn = 0
'        Northing = nfn + t1 + (dlam * dlam) * t2 + (dlam * dlam * dlam _
'            * dlam) * t3 + (dlam * dlam * dlam * dlam * dlam * dlam) + 0.5
Northing = nfn + t1 + (dlam * dlam) * t2 + (dlam * dlam * dlam _
            * dlam) * t3 + (dlam * dlam * dlam * dlam * dlam * dlam)
        t6 = sn * c * ok
        t7 = sn * (c * c * c) * (1# - (t * t) + eta) / 6#
'        Easting = fe + dlam * t6 + (dlam * dlam * dlam) * t7 + 0.5
Easting = fe + dlam * t6 + (dlam * dlam * dlam) * t7
        If (Northing >= 9999999#) Then Northing = 9999999#
       ' Return New Double() {utmXZone, easting, utmYzone, northing}
        LatLonToUTM.Northing = Round(Northing, 3)
        LatLonToUTM.Easting = Round(Easting, 3)
        LatLonToUTM.Zone = utmXZone
        
        Exit Function
ErrHandler:
        MsgBox ("Error Converting")
        
    End Function
    

    
    Public Function CalculateE2Squared(ByVal a As Double, ByVal b As Double)

        CalculateE2Squared = ((a * a) - (b * b)) / (b * b)

    End Function

    Public Function CalculateESquared(ByVal a As Double, ByVal b As Double)

        CalculateESquared = ((a * a) - (b * b)) / (a * a)

    End Function
    
     Public Function sphsn(ByVal a As Double, ByVal es As Double, ByVal sphi As Double)
        Dim sinSphi As Double
        sinSphi = Sin(sphi)
        sphsn = a / (1# - es * (sinSphi * sinSphi)) ^ 0.5

    End Function
    
     Public Function sphtmd(ByVal ap As Double, ByVal bp As Double, ByVal cp As Double, ByVal dp As Double, ByVal ep As Double, ByVal sphi As Double)
        sphtmd = (ap * sphi) - (bp * Sin(2# * sphi)) + (cp * Sin(4# * sphi)) - (dp * Sin(6# * sphi)) + (ep * Sin(8# * sphi))

    End Function


