Attribute VB_Name = "Tirem_Declares"
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

' Tirem Global Inputs:

Public PROPFQ As Single     ' TRANSMITTER FREQUENCY IN MHZ

' PROFILE TO EACH PROFILE POINT IN METERS
Public EXTNSN As Long       ' PROFILE INDICATOR FLAG (True or False)

' Tirem Outputs:
Public VRSION As String * 8 ' TIREM VERSION NUMBER

'

'  Here are the declares for the 32-bit FORTRAN DLL: Tirem.DLL.
'

'  Note: Variable names are identical to the FORTRAN source code
'        used to create the Tirem.DLL.

Declare Sub GetTiremVersion Lib "Tirem.dll" _
    Alias "_get_tirem_version@4" (ByVal VRSION As String)
   
Declare Sub TiremAnalysis Lib "TiremAnalysis.dll" _
    (ByRef TANTHT As Single, ByRef RANTHT As Single, ByRef PROPFQ As Single, _
    ByRef NPRFL As Long, ByRef HPRFL As Single, ByRef XPRFL As Single, _
    ByRef EXTNSN As Long, ByRef REFRAC As Single, ByRef CONDUC As Single, _
    ByRef PERMIT As Single, ByRef HUMID As Single, ByVal POLARZ As String, _
    ByVal VRSION As String, ByVal Mode As String, ByRef PRLoss As Single, _
    ByRef FSPLSS As Single, ByRef TOTTRO As Single, ByRef TOTDIF As Single, _
    ByRef ABLOSS As Single, ByRef THET00 As Single, ByRef TXANG As Single, _
    ByRef RXANG As Single, ByRef ALPHAE As Single, ByRef BETAE As Single, _
     ByRef HORZTX As Long, ByRef HORZRX As Long)
    
' Tirem Inputs:
' TANHT    TRANSMITTER STRUCTURAL ANTENNA HEIGHT IN METERS
' RANTHT   RECEIVER STRUCTURAL ANTENNA HEIGHT IN METERS
' PROPFQ  TRANSMITTER FREQUENCY IN MHZ
' NPRFL   TOTAL NUMBER OF PROFILE POINTS FOR THE ENTIRE PATH
' HPRFL   ARRAY OF PROFILE TERRAIN HEIGHTS ABOVE MEAN SEA LEVEL IN METERS
' XPRFL   ARRAY OF GREAT CIRCLE DISTANCES FROM THE BEGINNING OF THE
'         PROFILE TO EACH PROFILE POINT IN METERS
' EXTNSN  PROFILE INDICATOR FLAG (True or False)
' REFRAC  SURFACE REFRACTIVITY MEASURED IN "N UNITS"
' CONDUC  CONDUCTIVITY OF EARTH SURFACE MEASURED IN S/M
' PERMIT  RELATIVE PERMITTIVITY OF EARTH SURFACE
' HUMID   SURFACE HUMIDITY AT THE TRANSMITTER SITE G/M**3
' POLARZ  TRANSMITTER ANTENNA POLARIZATION
'
' Tirem Outputs:
' VRSION  TIREM VERSION NUMBER
' MODE    MODE INDICATOR: LINE OF SIGHT, DIFFRACTION, or TROPO SCATTER
' PRLOSS    TOTAL PATH LOSS (BASIC TRANSMISSION LOSS) IN DB
' FSPLSS  FREE SPACE LOSS IN DB
    
