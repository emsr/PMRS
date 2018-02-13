Attribute VB_Name = "Fortran_DLL_Declares"
'******************************************************************************
'
' *** COPYRIGHT:  Copyright 1999
'                 IIT Research Institute
'                 US Govt retains rights in accordance
'                 with DoD FAR Supp 252.227 - 7013.
'
'-----------------------------------------------------------------------
'
' *** CODING HISTORY:
'
'   DATE      PROGRAMMER                     DESCRIPTION
' --------   ------------                   -------------
' 4/21/99    Mike Fahler                    PMRS v. 1.00
'******************************************************************************
Public status As Integer     'error code from the Dll call
Public errmess As String * 50 'error message returned by the Dlls
Public PrevUnit As String   'the previous unit of measure selected
Public NewUnit As String    'unit of measure to covert to
Public FromUnit As Integer  'numeric value cooresponding to previous unit type string
Public ToUnit As Integer   'numeric value cooresponding to new unit type string
Public FromVal As Single    'numeric value to convert
Public ToVal As Single      'numeric value after conversion
Public MeasId As Integer    'sets the measurement type in a number format (1-ant height,2-freq,3-elev,4-dist,5-ant dim)
Public CurUnits As Integer
Public UseMaxOfElev2 As Boolean
Public UserEnterElev2 As Boolean
Public UseMaxOfElev1 As Boolean
Public UserEnterElev1 As Boolean
Public UserEnteredPoints As Boolean

Public Declare Sub BoundChk Lib "BoundChk.dll" _
(ByRef CurVal As Single, ByRef MeasId As Integer, _
 ByRef CurUnits As Integer, ByRef status As Integer, _
 ByVal errmess As String)

Declare Sub Coordchk Lib "Coordchk.dll" _
(ByRef Coordtype As Long, ByRef degrees As Single, _
 ByRef minutes As Single, ByRef seconds As Single, _
 ByVal Hemisphere As String, ByRef status As Long, _
 ByVal errmess As String)


Declare Sub ConvUnits Lib "ConvUnits.dll" _
(ByRef MeasId As Integer, ByRef FromUnits As Integer, ByRef FromVal As Single, _
 ByRef ToUnits As Integer, ByRef ToVal As Single)

Declare Sub NSubS2 Lib "NSubS2.dll" _
(ByRef TXELEV As Single, ByRef RXELEV As Single, ByRef Txantht As Single, _
 ByRef Rxantht As Single, ByRef refract As Single, ByRef surfrefrac As Single)


Declare Sub CalcAptomdLoss Lib "CalcAptomdLoss.dll" _
(ByRef TxBeamw As Single, ByRef RxBeamw As Single, ByRef AlphaEf As Single, _
ByRef BetaEf As Single, ByRef Pathlen As Single, ByRef surfrefract As Single, _
ByRef diffractloss As Single, ByRef tropoloss As Single, ByVal Tiremmode As String, _
ByRef proploss As Single, ByRef CouplingLoss As Single)


Declare Sub CalcEffAnt Lib "CalcEffAnt.dll" _
(ByRef TANTHT As Single, ByRef RANTHT As Single, ByRef HORZTX As Long, _
ByRef HORZRX As Long, ByRef NumElements As Long, ByRef HPRFL As Single, _
ByRef XPRFL As Single, ByRef TxEffHeight As Single, ByRef RxEffHeight As Single)

Declare Sub CalcGrConst Lib "CalcGroundConstants.dll" _
(ByRef propfreq As Single, ByRef groundtype As Long, _
 ByRef permitivity As Single, ByRef conductivity As Single)

Declare Sub CalcPowDens Lib "CalcPowDens.dll" _
(ByRef Pathloss As Single, ByRef txpowr As Single, ByRef TXGAIN As Single, _
ByRef propfreq As Single, ByRef POWDEN As Single, ByRef RMSFS As Single)

Declare Sub CalcModVar Lib "CalcModVar.dll" _
(ByRef probability As Single, ByRef proploss As Single, ByRef MODVARArray As Single)

Declare Sub PINFCP2 Lib "pinfcp2.dll" _
(ByRef EarthRadFactor As Single, ByRef TxFreq As Single, ByRef ScaleFactor As Long, _
 ByRef pathlength As Single, ByVal units As String, ByRef Numelemt As Long, _
 ByRef XPRFL As Single, ByRef HPRFL As Single, ByRef Featurehts As Single, _
 ByRef AbsClearance As Single, ByRef avgpathht As Single, ByRef minfresclr As Single, _
 ByRef dabslcear As Single, ByRef dminfresclr As Single, ByRef txtakoffangle As Single, _
 ByRef rxtakeoffangle As Single, ByRef penangle As Single, ByRef RayPath As Single, _
 ByRef FresnelPath As Single, ByRef ERRCODE As Long)
 
Declare Sub Datum2Axis Lib "Datum2Axis.dll" _
(ByRef Datum As Long, ByRef spheroid As Long, ByRef MJAxis As Single, _
ByRef MNAxis As Single, ByRef Flat As Single, ByRef ecc As Single, _
ByRef ERRCODE As Long)

Declare Sub CalcNFDist Lib "NearFieldDist.dll" _
(ByRef Tantdim As Single, ByRef Rantdim As Single, ByRef pathlength As Single, _
ByRef PROPFQ As Single, ByRef TXNFDist As Single, ByRef RXNFDist As Single, _
ByRef comptxantdim As Single, ByRef comprxantdim As Single, _
ByRef status As Long, ByVal NFWarnmsg As String)

