DECLARE FUNCTION Sin2# (x#)
'***********
'* GEO2XYZ *
'***********

'From geographical coordinate to cartesian
'Da coordinate geografiche a coordinate cartesiane

'Program written by Roberto Maugeri June 1997

'DA CORREGGERE con Amg-agd.txt
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
e3 = e2 * e
e4 = e2 * e2
e6 = e3 * e3

'Nu             is the radius of curvature in the prime vertical.
'Rho            is the radius of curvature in the meridian.



t = TAN(Lat)

Omega = Longit - Longit0

Rho = a * (1 - e2) / (1 - e2 * Sin2(Lat)) ^ (1.5)

Nu = a / (1 - e2 * Sin2(Lat)) ^ (.5)

Psi = Nu / Rho

' Mdist is the Meridian distance.


A0 = 1 - (e2) / 4 - (3 * e4) / 64 - (5 * e6) / 256
A2 = (3 / 8) * (e2 + (e4) / 4 + (15 * e6) / 128)
A4 = (15 / 256) * (e4 + (3 * e6) / 4)
A6 = (35 / 3072) * e6
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
                                                 

'Formulae (geographical to grid)
'-------------------------------

t2 = t * t
t4 = t2 * t2
t6 = t4 * t2
Psi2 = Psi * Psi
Psi3 = Psi2 * Psi
Omega2 = Omega * Omega
Omega4 = Omega2 * Omega2
Omega6 = Omega4 * Omega2
Omega8 = Omega4 * Omega4

Term1 = 1 / 6 * (Omega2 * COS(Lat) ^ 2 * (Psi - t2))

Term2 = 1 / 120 * (Omega4 * COS(Lat) ^ 4) * (4 * Psi3 * (1 - 6 * t2) + Psi2 * (1 + 8 * t2) - Psi * (2 * t2) + t4)
         
Term3 = 1 / 5040 * (Omega6 * COS(Lat) ^ 6) * ((61 - 479 * t2 + 179 * t4 - t6))

Edash = K0 * Nu * Omega * COS(Lat) * (1 + Term1 + Term2 + Term3)

EASTING = Edash + FalseEasting
                    
                     


Term1 = Nu * SIN(Lat) * ((Omega2) / 2) * COS(Lat)

Term2 = Nu * SIN(Lat) * ((Omega4) / 24) * COS(Lat) ^ 3 * (4 * Psi2 + Psi - t2)

Term3a = Nu * SIN(Lat) * ((Omega6) / 720) * COS(Lat) ^ 5
Term3ba = 8 * Psi4 * (11 - 24 * t2) - 28 * Psi3 * (1 - 6 * t2)
Term3bb = Psi2 * (1 - 32 * t2) - Psi * (2 * t2) + t4
Term3b = Term3ba + Term3bb
Term3 = Term3a * Term3b

Term4 = Nu * SIN(Lat) * ((Omega8) / 40320) * COS(Lat) ^ 7 * (1385 - 3111 * t2 + 543 * t4 - t6)
                                            
Ndash = K0 * (Mdist + Term1 + Term2 + Term3 + Term4)

NORTHING = Ndash + FalseNorthing




PRINT
PRINT "Easting "; EASTING
PRINT "Northing "; NORTHING
PRINT "Zone "; ZONE


FUNCTION Sin2 (x)
       Sin2 = SIN(x) * SIN(x)
END FUNCTION

