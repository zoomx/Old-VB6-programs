DECLARE FUNCTION Sec# (x#)
'***********
'* GEO2XYZ *
'***********

'From geographical coordinate to cartesian
'Da coordinate geografiche a coordinate cartesiane

'Program written by Roberto Maugeri June 1997

'DA CORREGGERE con Amg-agd.txt
'Realizzare la funzione per sec()
'Ci sono differenze di alcune alcuni cm con il foglio Excel Redfearn.xls


'AUSLIG GEODESY
'Redfearn's Formulae

'Redfearn's formulae were published in the "Empire Survey Review", No. 69,
'1948. They may be used to convert between latitude & longitude and easting &
'northing for a Transverse Mercator projection, as used for the Australian
'Map Grid (AMG).



COLOR 7, 1
CLS
DEFDBL A-Z
Debug = 1

'Variables
Pi = 3.141592654#
Factr = Pi / 180#
IF Debug THEN PRINT "Factr ="; Factr

'Lat            is the latitude (radians) (negative south).
Latg = -37
Latm = 39
Lats = 15.5571
Latgg = (ABS(Latg) + Latm / 60 + Lats / 3600) * SGN(Latg)
Lat = Latgg * Factr
IF Debug THEN PRINT "Latgg "; Latgg
IF Debug THEN PRINT "Lat "; Lat
'The use of ABS and SGN take in account the sign of Latg because if it
'is negative the sum for Latgg became wrong.

'Longit         is the Longitude (radians) (negative west).
Longitg = 143
Longitm = 55
Longits = 30.663
Longitgg = (ABS(Longitg) + Longitm / 60 + Longits / 3600) * SGN(Longitg)
Longit = Longitgg * Factr
IF Debug THEN PRINT "Longitgg "; Longitgg
IF Debug THEN PRINT "Longit "; Longit

'False easting  is the distance between the true origin and the false origin,
'               used to avoid negative grid references (500,000 for AMG).
FalseEasting = 500000


'False northing is the distance between the true origin and the false origin,
'               used to avoid negative grid references (10,000,000 for AMG).
FalseNorthing = 10000000


'K0             is the central scale factor (0.9996 for AMG).
K0 = .9996#

LWEZZ = -186  'Longitude of Western Edge of Zone Zero
CMZZ = -183   'Central meridian of zone zero
ZW = 6        'Zone width (degrees)

ZONE = INT((Longitgg - LWEZZ) / ZW)

'Longit0        is the Longitude of the central meridian (radians).
Longit0 = ((ZONE * ZW) + CMZZ) * Factr 'Factr is for converting in radians

IF Debug THEN PRINT "Longit0 "; Longit0 / Factr; Longit

'a              is the semi major axis of the reference ellipsoid
'               (6,378,160 metres for the AMG).
a = 6378160#

'invf=1/f       is the inverse of the flattening of the reference ellipsoid
'               (298.25 for the AMG).

invf = 298.25#
f = 1 / invf

'e^2 =e*e =e2   is the eccentricity of the reference ellipsoid
'               (0.006 694 541 855 for the AMG).
e2a = .006694541855#
e2b = (2 * f) - (f * f)
IF Debug THEN PRINT "Confronto e2a-e2b "; e2a - e2b

e2 = e2b
e = SQR(e2)


'Nu             is the radius of curvature in the prime vertical.
'Rho            is the radius of curvature in the meridian.



t = TAN(Lat)

Omega = Longit - Longit0

Rho = a * (1 - e2) / (1 - e2 * (SIN(Lat) ^ 2)) ^ (3 / 2)

Nu = a / (1 - e2 * (SIN(Lat)) ^ 2) ^ (1 / 2)

Psi = Nu / Rho

' Mdist is the Meridian distance.


A0 = 1 - (e2) / 4 - (3 * e ^ 4) / 64 - (5 * e ^ 6) / 256
A2 = (3 / 8) * (e2 + (e ^ 4) / 4 + (15 * e ^ 6) / 128)
A4 = (15 / 256) * (e ^ 4 + (3 * e ^ 6) / 4)
A6 = (35 / 3072) * e ^ 6
Mdist = a * (A0 * Lat - A2 * SIN(2 * Lat) + A4 * SIN(4 * Lat) - A6 * SIN(6 * Lat))
IF Debug THEN PRINT "MDist "; Mdist / Factr; Mdist

'For the Australian Map Grid the formula for meridian distance reduces to:
                                        
Mdist2 = 111133.348785# * (57.29577951# * Lat) - 16038.9546# * SIN(2 * Lat) + 16.8331 * SIN(4 * Lat) - .0218 * SIN(6 * Lat)
                                               
Residual = Mdist - Mdist2
IF Debug THEN PRINT
IF Debug THEN PRINT "Residual of Mdist = "; Residual

'111133.34... is the mean length of 1 degree of latitude in metres (G).
                                                
'If Lat were in degrees, then the 57.295... term, which converts radians
'to degrees, would not be needed.
                                                 

'Formulae (grid to geographical)
'-------------------------------

n = f / (2 - f)
' = 0.0016792611 for the AMG
     
G = a * (1 - n) * (1 - n ^ 2) * (1 + (9 / 4) * n ^ 2 + (225 / 64) * n ^ 4) * Factr
'= 111 133.348785 for the AMG
           
Ndash = Northing - FalseNorthing

Mdist = Ndash / K0
           
Sigma = (Mdist * Pi) / (G * 180)
           
'FPLat = Foot point latitude
FPlat1 = Sigma + ((3 * n / 2) - (27 * n ^ 3 / 32)) * SIN(2 * Sigma)
'+
FPlat2 = ((21 * n ^ 2 / 16) - (55 * n ^ 4 / 32)) * SIN(4 * Sigma)
'+
FPlat3 = (151 * n ^ 3 / 96) * SIN(6 * Sigma)
'+
FPlat4 = (1097 * n ^ 4 / 512) * SIN(8 * Sigma)

FPLat = FPlat1 + FPlat2 + FPlat3 + FPlat4

'For AMG
'FPlat1 = Sigma + .0025188877# * SIN(2 * Sigma)
'+
'FPlat2 = .0000037011# * SIN(4 * Sigma)
'+
'FPlat3 = .0000000074# * SIN(6 * Sigma)
'+
'FPlat4 = .000000000017# * SIN(8 * Sigma)
                                                                                                           
'In the following formulae t, Rho, Nu and Psi are all evaluated for the
'foot point latitude computed above.
                                                                                                           
Edash = Easting - FalseEasting
                                                                                                           
x = Edash / (K0 * Nu)
                                                                                                           

Term1 = (t / (K0 * Rho)) * x * Edash / 2
                                                                                                           
Term2a = (t / (K0 * Rho)) * ((Edash * x ^ 3) / 24)
Term2b = (-4 * Psi ^ 2 + 9 * Psi(1 - t ^ 2) + 12 * t ^ 2)
Term2 = Term1 * Term2

Term3a = (t / (K0 * Rho)) * ((Edash * x ^ 5) / 720)
'*
Term3ba=(8*Psi^4 * (11 - 24*t^2) - 12*Psi^3 * (21 - 71*t^2)
'+            
Term3bb = 15 * Psi ^ 2 * (15 - 98 * t ^ 2 + 15 * t ^ 4)
'+           
Term3bc = 180 * Psi * (5 * t ^ 2 - 3 * t ^ 4) + 360 * t ^ 4
Term3b = Term3ba + Term3bb + Term3bc
Term3 = Term3a + Term3b

Term4a = (t / (K0 * Rho)) * ((Edash * x ^ 7) / 40320)
'*
Term4b = (1385 + 3633 * t ^ 2 + 4095 * t ^ 4 + 1575 * t ^ 6)
Term4 = Term4a + Term4b

Latitude = FPLat - Term1 + Term2 - Term3 + Term4
                                                                                                                                                   

                                                                                                                                                   
Term1 = x * Sec(FPLat)
                                                                                                                                                   
Term2 = (x ^ 3 / 6) * Sec(FPLat) * (Psi + 2 * t ^ 2)
                                                                                                                                                   
Term3a = (x ^ 5 / 120) * Sec(FPLat)
'*
Term3b = (-4 * Psi ^ 3 * (1 - 6 * t ^ 2) + Psi ^ 2 * (9 - 68 * t ^ 2) + 72 * Psi * t ^ 2 + 24 * t ^ 4)
Term3 = Term3a + Term3b

Term4 = (x ^ 7 / 5040) * Sec(FPLat) * (61 + 662 * t ^ 2 + 1320 * t ^ 4 + 720 * t ^ 6)
                                                                                                                                                          
Omega = Term1 - Term2 + Term3 - Term4

'Longitude = Longitude of central meridian + Omega
Longitude = LCM + Omega



PRINT
PRINT "Latitude "; Latitude
PRINT "Longitude "; Longitude



FUNCTION Sec (x)
  Sec = 1 / COS(x)
END FUNCTION

