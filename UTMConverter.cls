VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UTMConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const PI = 3.14159265358979
Const sm_a = 6378137#
Const sm_b = 6356752.314
Const UTMScaleFactor = 0.9996

Public Function DegToRad(deg As Double) As Double
    DegToRad = deg / 180# * PI
End Function


Public Function RadToDeg(rad As Double) As Double
    RadToDeg = rad / PI * 180#
End Function

Private Function ArcLengthOfMeridian(ByVal phi As Double) As Double '获得某一点距离中央经线弧线距离
    Dim alpha As Double, beta As Double, gamma As Double, delta As Double, epsilon As Double, n As Double
    n = (sm_a - sm_b) / (sm_a + sm_b)
    alpha = (sm_a + sm_b) / 2# * (1# + n ^ 2# / 4# + n ^ 4# / 64#)
    beta = -3# * n / 2# + 9# * n ^ 3# / 16# + (-3# * n ^ 5#) / 32#
    gamma = 15# * n ^ 2# / 16# + (-15# * n ^ 4# / 32#)
    delta = -35# * n ^ 3# / 48# + 105# * n ^ 5# / 256#
    epsilon = 315# * n ^ 4# / 512#
    ArcLengthOfMeridian = alpha * (phi + (beta * VBA.Sin(2# * phi)) + (gamma * VBA.Sin(4# * phi)) + (delta * VBA.Sin(6# * phi)) + (epsilon * VBA.Sin(8# * phi)))
End Function

Private Function UTMCentralMeridian(ByVal zone As Double) As Double  '不同带中央经线的经度
    UTMCentralMeridian = DegToRad(-183# + zone * 6#)
End Function

Private Function FootpointLatitude(ByVal y As Double) As Double
    Dim y_ As Double, alpha_ As Double, beta_ As Double, gamma_ As Double, delta_ As Double, epsilon_ As Double, n As Double
    n = (sm_a - sm_b) / (sm_a + sm_b)
    alpha_ = (sm_a + sm_b) / 2# * (1 + n ^ 2# / 4 + n ^ 4# / 64)
    y_ = y / alpha_
    beta_ = 3# * n / 2# + (-27# * n ^ 3# / 32#) + (269# * n ^ 5# / 512#)
    gamma_ = 21# * n ^ 2# / 16# + (-55# * n ^ 4# / 32#)
    delta_ = 151# * n ^ 3# / 96# + (-417# * n ^ 5#) / 128#
    epsilon_ = 1097# * n ^ 4# / 512#
    FootpointLatitude = y_ + beta_ * VBA.Sin(2# * y_) + gamma_ * VBA.Sin(4# * y_) + delta_ * VBA.Sin(6# * y_) + epsilon_ * VBA.Sin(8# * y_)
End Function

Private Sub MapLatLonToXY(ByVal phi As Double, ByVal lambda As Double, ByVal lambda0 As Double, ByRef x As Double, ByRef y As Double)
    Dim n As Double, nu2 As Double, ep2 As Double, t As Double, t2 As Double, l As Double
    Dim l3coef As Double, l4coef As Double, l5coef As Double, l6coef As Double, l7coef As Double, l8coef As Double
    Dim tmp As Double
    ep2 = (sm_a ^ 2# - sm_b ^ 2#) / sm_b ^ 2#
    nu2 = ep2 * VBA.Cos(phi) ^ 2#
    n = sm_a ^ 2# / (sm_b * VBA.Sqr(1 + nu2))
    t = VBA.Tan(phi)
    t2 = t * t
    tmp = (t2 * t2 * t2) - t ^ 6#
    l = lambda - lambda0
    l3coef = 1# - t2 + nu2
    l4coef = 5# - t2 + 9 * nu2 + 4# * (nu2 * nu2)
    l5coef = 5# - 18# * t2 + (t2 * t2) + 14# * nu2 - 58# * t2 * nu2
    l6coef = 61# - 58# * t2 + (t2 * t2) + 270# * nu2 - 330# * t2 * nu2
    l7coef = 61# - 479# * t2 + 179# * (t2 * t2) - (t2 * t2 * t2)
    l8coef = 1385# - 3111# * t2 + 543# * (t2 * t2) - (t2 * t2 * t2)
    x = n * VBA.Cos(phi) * l + (n / 6# * Math.Cos(phi) ^ 3# * l3coef * l ^ 3#) + (n / 120# * Math.Cos(phi) ^ 5# * l5coef * l ^ 5#) + (n / 5040# * Math.Cos(phi) ^ 7# * l7coef * l ^ 7#)
    y = ArcLengthOfMeridian(phi) + (t / 2# * n * Math.Cos(phi) ^ 2# * l ^ 2#) + (t / 24# * n * Math.Cos(phi) ^ 4# * l4coef * l ^ 4#) + (t / 720# * n * Math.Cos(phi) ^ 6# * l6coef * l ^ 6#) + (t / 40320# * n * Math.Cos(phi) ^ 8# * l8coef * l ^ 8#)
End Sub

Private Sub MapXYToLatLon(ByVal x As Double, ByVal y As Double, ByVal lambda0 As Double, ByRef philambda1 As Double, ByRef philambda2 As Double)
    Dim phif As Double, Nf As Double, Nfpow As Double, nuf2 As Double, ep2 As Double, tf As Double, tf2 As Double, tf4 As Double, cf As Double
    Dim x1frac As Double, x2frac As Double, x3frac As Double, x4frac As Double, x5frac As Double, x6frac As Double, x7frac As Double, x8frac As Double
    Dim x2poly As Double, x3poly As Double, x4poly As Double, x5poly As Double, x6poly As Double, x7poly As Double, x8poly As Double
    phif = FootpointLatitude(y)
    ep2 = (sm_a ^ 2# - sm_b ^ 2#) / sm_b ^ 2#
    cf = VBA.Cos(phif)
    nuf2 = ep2 * cf ^ 2#
    Nf = sm_a ^ 2# / (sm_b * VBA.Sqr(1 + nuf2))
    Nfpow = Nf
    tf = VBA.Tan(phif)
    tf2 = tf * tf
    tf4 = tf2 * tf2
    x1frac = 1# / (Nfpow * cf)
    Nfpow = Nfpow * Nf
    x2frac = tf / (2# * Nfpow)
    Nfpow = Nfpow * Nf
    x3frac = 1# / (6# * Nfpow * cf)
    Nfpow = Nfpow * Nf
    x4frac = tf / (24# * Nfpow)
    Nfpow = Nfpow * Nf
    x5frac = 1# / (120# * Nfpow * cf)
    Nfpow = Nfpow * Nf
    x6frac = tf / (720# * Nfpow)
    Nfpow = Nfpow * Nf
    x7frac = 1# / (5040# * Nfpow * cf)
    Nfpow = Nfpow * Nf
    x8frac = tf / (40320# * Nfpow)
    x2poly = -1# - nuf2
    x3poly = -1# - 2 * tf2 - nuf2
    x4poly = 5# + 3# * tf2 + 6# * nuf2 - 6# * tf2 * nuf2 - 3# * (nuf2 * nuf2) - 9# * tf2 * (nuf2 * nuf2)
    x5poly = 5# + 28# * tf2 + 24# * tf4 + 6# * nuf2 + 8# * tf2 * nuf2
    x6poly = -61# - 90# * tf2 - 45# * tf4 - 107# * nuf2 + 162# * tf2 * nuf2
    x7poly = -61# - 662# * tf2 - 1320# * tf4 - 720# * (tf4 * tf2)
    x8poly = 1385# + 3633# * tf2 + 4095# * tf4 + 1575 * (tf4 * tf2)
    philambda1 = phif + x2frac * x2poly * (x * x) + x4frac * x4poly * x ^ 4# + x6frac * x6poly * x ^ 6# + x8frac * x8poly * x ^ 8#
    philambda2 = lambda0 + x1frac * x + x3frac * x3poly * x ^ 3# + x5frac * x5poly * x ^ 5# + x7frac * x7poly * x ^ 7#
End Sub


Private Sub LatLonToUTMXY(ByVal lat As Double, ByVal lon As Double, ByRef zone As Double, ByRef x As Double, ByRef y As Double, ByRef southhemisphere As Boolean)
    zone = VBA.Int((RadToDeg(lon) + 180#) / 6) + 1
    If RadToDeg(lat) < 0 Then southhemisphere = True Else southhemisphere = False
    MapLatLonToXY lat, lon, UTMCentralMeridian(zone), x, y
    x = x * UTMScaleFactor + 500000#
    y = y * UTMScaleFactor
    If y < 0# Then y = y + 10000000#
End Sub

Private Sub UTMXYToLatLon(ByVal zone As Double, ByVal x As Double, ByVal y As Double, ByVal southhemi As Boolean, ByRef lat As Double, ByRef lon As Double)
    Dim cmeridian As Double
    x = x - 500000#
    x = x / UTMScaleFactor
    If southhemi Then y = y - 10000000#
    y = y / UTMScaleFactor
    cmeridian = UTMCentralMeridian(zone)
    MapXYToLatLon x, y, cmeridian, lat, lon
End Sub

Public Function GetCentralMeridianXAndEquatorY(ByVal lat1 As Double, ByVal lon1 As Double, ByVal x1 As Double, ByVal y1 As Double, _
    ByVal lat2 As Double, ByVal lon2 As Double, ByVal x2 As Double, ByVal y2 As Double, ByRef x As Double, ByRef y As Double) As Boolean
    
    Dim zone1 As Double, zone2 As Double, e1 As Double, n1 As Double, e2 As Double, n2 As Double, south1 As Boolean, south2 As Boolean
    LatLonToUTMXY lat1, lon1, zone1, e1, n1, south1
    LatLonToUTMXY lat2, lon2, zone2, e2, n2, south2
    
    If south1 <> south2 Then GetCentralMeridianXAndEquatorY = False: Exit Function
    If zone1 <> zone2 Then GetCentralMeridianXAndEquatorY = False: Exit Function
    If lat1 = lat2 Then GetCentralMeridianXAndEquatorY = False: Exit Function
    If lon1 = lon2 Then GetCentralMeridianXAndEquatorY = False: Exit Function
    If x1 = x2 Then GetCentralMeridianXAndEquatorY = False: Exit Function
    If y1 = y2 Then GetCentralMeridianXAndEquatorY = False: Exit Function
    If (lon1 - lon2) * (x1 - x2) < 0 Then GetCentralMeridianXAndEquatorY = False: Exit Function
    If (lat1 - lat2) * (y1 - y2) < 0 Then GetCentralMeridianXAndEquatorY = False: Exit Function
    
    x = x1 + (x2 - x1) / (e2 - e1) * (500000# - e1)

    If south1 Then
        y = y1 + (10000000# - n1) / (n1 - n2) * (y1 - y2)
    Else
        y = y1 - (y1 - y2) / (n1 - n2) * n1
    End If
    
    GetCentralMeridianXAndEquatorY = True
End Function

Public Function GetPointLatLon(ByVal lat1 As Double, ByVal lon1 As Double, ByVal x1 As Double, ByVal y1 As Double, _
    ByVal lat2 As Double, ByVal lon2 As Double, ByVal x2 As Double, ByVal y2 As Double, _
    ByVal px As Double, ByVal py As Double, ByRef plat As Double, ByRef plon As Double) As Boolean
    
    Dim zone1 As Double, zone2 As Double, e1 As Double, n1 As Double, e2 As Double, n2 As Double, south1 As Boolean, south2 As Boolean
    Dim pe As Double, pn As Double
    LatLonToUTMXY lat1, lon1, zone1, e1, n1, south1
    LatLonToUTMXY lat2, lon2, zone2, e2, n2, south2
    If south1 <> south2 Then GetPointLatLon = False: Exit Function
    If zone1 <> zone2 Then GetPointLatLon = False: Exit Function
    If lat1 = lat2 Then GetPointLatLon = False: Exit Function
    If lon1 = lon2 Then GetPointLatLon = False: Exit Function
    If x1 = x2 Then GetPointLatLon = False: Exit Function
    If y1 = y2 Then GetPointLatLon = False: Exit Function
    If (lon1 - lon2) * (x1 - x2) < 0 Then GetPointLatLon = False: Exit Function
    If (lat1 - lat2) * (y1 - y2) < 0 Then GetPointLatLon = False: Exit Function
    pe = e1 + (px - x1) / (x1 - x2) * (e1 - e2)
    pn = n1 + (py - y1) / (y1 - y2) * (n1 - n2)
    
    UTMXYToLatLon zone1, pe, pn, south1, plat, plon
    GetPointLatLon = True
End Function

Public Function GetPointXY(ByVal lat1 As Double, ByVal lon1 As Double, ByVal x1 As Double, ByVal y1 As Double, _
    ByVal lat2 As Double, ByVal lon2 As Double, ByVal x2 As Double, ByVal y2 As Double, _
    ByVal plat As Double, ByVal plon As Double, ByRef px As Double, ByRef py As Double) As Boolean
    
    Dim zone1 As Double, zone2 As Double, e1 As Double, n1 As Double, e2 As Double, n2 As Double, south1 As Boolean, south2 As Boolean
    Dim pe As Double, pn As Double, south3 As Boolean, zone3 As Double, x0 As Double, y0 As Double
    LatLonToUTMXY lat1, lon1, zone1, e1, n1, south1
    LatLonToUTMXY lat2, lon2, zone2, e2, n2, south2
    LatLonToUTMXY plat, plon, zone3, pe, pn, south3
    
    If zone1 <> zone2 Or zone1 <> zone3 Then GetPointXY = False: Exit Function
    If south1 <> south2 Then GetPointXY = False: Exit Function
    If lat1 = lat2 Then GetPointXY = False: Exit Function
    If lon1 = lon2 Then GetPointXY = False: Exit Function
    If x1 = x2 Then GetPointXY = False: Exit Function
    If y1 = y2 Then GetPointXY = False: Exit Function
    If (lon1 - lon2) * (x1 - x2) < 0 Then GetPointXY = False: Exit Function
    If (lat1 - lat2) * (y1 - y2) < 0 Then GetPointXY = False: Exit Function
    px = x1 + (x1 - x2) / (e1 - e2) * (pe - e1)
    GetCentralMeridianXAndEquatorY lat1, lon1, x1, y1, lat2, lon2, x2, y2, x0, y0
    If south3 Then
        py = y0 - (y1 - y2) / (n1 - n2) * (10000000# - pn)
    Else
        py = y0 + (y1 - y2) / (n1 - n2) * (pn)
    End If
    
    GetPointXY = True
End Function
