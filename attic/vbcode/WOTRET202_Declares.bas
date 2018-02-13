Attribute VB_Name = "WOTRET202_Declares"
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

Option Explicit

'
'  Here are the Visual Basic v4.0 declares for the 32-bit FORTRAN DLL: WotRet202.DLL.
'       Version 1.0 Frank Cramblitt December 1996
'

'
'  Note: Variable names are identical to the FORTRAN source code
'        used to create the WotRet202.DLL.

Public Const WOTRET_MAXIMUM_OF_FOUR_CLOSEST_POINTS = 0
Public Const WOTRET_NEAREST_POINT = 1
Public Const PRFILE_FOUR_POINT_INTERPOLATION = 3
Public Const TOPGET_FOUR_POINT_INTERPOLATION = 4

Public Const MXNELV As Long = 2000 ' Maximum number of elevations returned.'
Option Base 1
Public HPRFL() As Single    ' ARRAY OF PROFILE TERRAIN HEIGHTS ABOVE MEAN SEA LEVEL IN METERS
Public XPRFL() As Single    ' ARRAY OF GREAT CIRCLE DISTANCES FROM THE BEGINNING OF THE
' GetProfile Arguments:
Public NUMELV As Long       ' Number of elevations in the profile returned by GetProfile.


'
'********************************************************************************
'

Declare Sub GetWotretNum Lib "wotret202.dll" _
   (ByVal VRSION As String)
   
'
'********************************************************************************
'
Declare Sub GetProfile Lib "wotret202.dll" _
     (ByRef LATB As Single, ByRef LONB As Single, _
     ByRef LATE As Single, ByRef LONE As Single, _
     ByRef SPACNG As Single, ByRef MJAxis As Single, _
     ByRef Flat As Single, ByRef Datum As Long, _
     ByVal TOPFIL As String, ByRef INTERP As Long, _
     ByVal ERROPT As String, ByRef DELELV As Single, _
     ByRef MXNELV As Long, ByRef DIST As Single, _
     ByRef ELEV As Single, ByRef NUMELV As Long, _
     ByRef PRFERR As Long)
'
' *** PURPOSE:
'
'     SUBROUTINE GetProfile EXTRACTS A TERRAIN PROFILE FROM TOPOGRAPHIC DATA
'     FILES GIVEN THE BEGINNING AND ENDING COORDINATES OF A PATH.
'
'-----------------------------------------------------------------------
'
' *** INPUT VARIABLES:
'
' NAME      TYPE      UNITS     DESCRIPTION
' ----      ----      -----     -----------
' LATB      REAL      RADIANS   LATITUDE OF THE BEGINNING OF PATH
'                               (NORTH IS POSITIVE, SOUTH IS NEGATIVE)
' LONB      REAL      RADIANS   LONGITUDE OF THE BEGINNING OF PATH
'                               (EAST IS POSITIVE, WEST IS NEGATIVE)
' LATE      REAL      RADIANS   LATITUDE OF THE END OF THE PATH
'                               (NORTH IS POSITIVE, SOUTH IS NEGATIVE)
' LONE      REAL      RADIANS   LONGITUDE OF THE END OF THE PATH
'                               (EAST IS POSITIVE, WEST IS NEGATIVE)
' SPACNG    REAL      METERS    SPACING BETWEEN THE POINTS OF THE
'                               PROFILE. THIS VALUE MAY BE ADJUSTED BY
'                               THIS SUBROUTINE
' MJAXIS    REAL      METERS    SEMI-MAJOR AXIS OF THE EARTH
' FLAT      REAL                FLATTENING OF THE EARTH
' DATUM     INTEGER             GEODETIC DATUM CODE
'                                  0 - WORLD GEODETIC SYSTEM 1984
'                                  1 - NORTH AMERICAN 1972
'                                  2 - EUROPEAN
'                                  3 - TOKYO
'                                  4 - GREAT BRITAIN
'                                  5 - MAUI  (OLD HAWAIIAN)
'                                  6 - OAHU  (OLD HAWAIIAN)
'                                  7 - KAUAI (OLD HAWAIIAN)
'                                  8 - KWAJALEIN(WAKE - ENIWETOK)
'                                  9 - WAKE ISLAND (WAKE-ENIWETOK)
'                                 10 - ENIWETOK ATOLL (WAKE-ENIWETOK)
'                                 11 - WAKE ISLAND ASTRO 1952
'                                 12 - GUAM 1963
'                                 13 - WORLD GEODETIC SYSTEM 1972
' TOPFIL    CHAR*12             TOPOGRAPHIC FILE NAME
' INTERP    INTEGER             INTERPOLATION OPTION
'                                  0 - MAXIMUM OF FOUR CLOSEST POINTS
'                                  1 - NEAREST POINT
'                                  4 - FOUR POINT INTERPOLATION
' ERROPT    CHAR*4              TOPO RETRIEVAL ERROR HANDLING INDICATOR
'                                 'E   ' - IF AN ERROR OCCURED IN THE
'                                          TOPO RETRIEVAL ROUTINE,
'                                          RETURN WITH THE APPROPRIATE
'                                          ERROR CONDITION SET
'                                 'D   ' - IF AN ERROR OCCURED IN THE
'                                          TOPO RETRIEVAL ROUTINE, SET
'                                          PROFILE ELEVATION TO THE
'                                          SUPPLIED DEFAULT ELEVATION
' DELELV    REAL      METERS    DEFAULT ELEVATION TO BE USED WHEN AN
'                               ELEVATION COULD NOT BE RETRIEVED FROM
'                               THE TOPOGRAPHIC DATA FILES
' MXNELV    INTEGER             MAXIMUM NUMBER OF ELEVATION POINTS IN
'                               THE PATH PROFILE
'
'-----------------------------------------------------------------------
'
' *** OUTPUT VARIABLES:
'
' NAME      TYPE      UNITS     DESCRIPTION
' ----      ----      -----     -----------
' DIST      REAL      METERS    ARRAY OF GREAT CIRCLE DISTANCES FROM THE
'                               BEGINNING OF THE PROFILE TO EACH PORFILE
'                               POINT.
' ELEV      REAL      METERS    ARRAY OF PROFILE TERRAIN HEIGHTS ABOVE
'                               MEAN SEA LEVEL
' NUMELV    INTEGER             ACTUAL NUMBER OF ELEVATION POINTS IN THE
'                               PATH PROFILE
' PRFERR    LOGICAL             PROFILE RETRIEVAL ERROR FLAG
'                                  .FALSE. - NO PROFILE RETRIEVAL ERROR
'                                  .TRUE.  - PROFILE RETRIEVAL ERROR
'
'********************************************************************************
'

Declare Sub GetElevation Lib "wotret202.dll" _
    (ByRef RLAT As Single, ByRef RLON As Single, _
    ByVal WOTLType As String, ByRef INTERP As Long, _
    ByVal TNAME As String, ByVal NAMLIST As String, _
    ByRef COUNT As Long, ByRef NSSP As Long, _
    ByRef Datum As Long, _
    ByRef WOTRER As Long, ByRef ELEV As Single)

'*** PURPOSE
'
' *** GENERAL, HIGH-LEVEL WOTL TOPO ELEVATION RETRIEVAL SUBROUTINE
'
'     THIS ROUTINE INTERFACES WITH THE WOTL TOPO RETRIEVAL SUBROUTINES:
'     TOPFND, NAMFND, WOTOPN, WOTELV, RECOVR, AND WOTCLS.  IT RETRIEVES
'     ONE ELEVATION AT A TIME AND WILL SPAN ACROSS MORE THAN THE ORIGINAL
'     DATA FILE OPENED IF IT CANNOT FIND AN ELEVATION FOR THE COORDINATE
'     REQUESTED.
'
'     THE TNAME VARIABLE IS USED TO DETERMINE WHICH ONE OF TWO MODES
'     TO SEARCH FOR A TOPOGRAPHIC DATA FILE THE USER WANTS:
'        1) TNAME = ' ' -->  GetElevation CALLS TOPFND TO SEARCH THE TOPO.INF
'                            FILE AND OPEN A TOPOGRAPHIC DATA FILE THAT
'                            CONTAINS THE SPECIFIED TRANSMITTER
'                            COORDINATES.
'        2) TNAME = 'XXXXXX'  --> GetElevation CALLS NAMFND TO SEARCH THE
'                                 TOPO.INF FILE AND OPEN A TOPOGRAPHIC
'                                 DATA FILE WITH THE SPECIFIED NAME.
'
'-----------------------------------------------------------------------------
'
' *** INPUT VARIABLES
'
' NAME   LOCATION   TYPE             DESCRIPTION
' ----   --------   ----             -----------
' RLAT   ARGUMENT   REAL/INTEGER     SIGNED FLOATING POINT (OR INTEGER)
'                                    LATITUDE OF POINT TO BE EXTRACTED.  THE
'                                    VARIABLE MAY BE RADIANS, MINUTES, OR
'                                    SECONDS DEPENDING UPON ITS DEFINITION
'                                    IN VARIABLE 'TYPE'.  POSITIVE VALUES
'                                    DENOTE NORTH, NEGATIVE VALUES DENOTE
'                                    SOUTH.
' RLON   ARGUMENT   REAL/INTEGER     SIGNED FLOATING POINT (OR INTEGER)
'                                    LONGITUDE OF POINT TO BE EXTRACTED.  THE
'                                    VARIABLE MAY BE RADIANS, MINUTES, OR
'                                    SECONDS DEPENDING UPON ITS DEFINITION
'                                    IN VARIABLE 'TYPE'.  POSITIVE VALUES
'                                    DENOTE WEST, NEGATIVE VALUES DENOTE EAST.
' TYPE   ARGUMENT   CHAR*4           UNITS AND VARIABLE TYPE FOR LAT AND LON:
'                                     'R   ' - RADIANS (REAL)
'                                     'M   ' - MINUTES (INTEGER)
'                                     'S   ' - SECONDS (INTEGER)
'                                     'MR  ' - MINUTES (REAL)
'                                     'SR  ' - SECONDS (REAL)
' INTERP ARGUMENT   INTEGER          INTERPOLATION OPTION:
'                                       0 - MAXIMUM OF FOUR CLOSEST POINTS
'                                       1 - TAKE ELEVATION OF NEAREST POINT
'                                       4 - FOUR POINT INTERPOLATION
' TNAME  ARGUMENT   CHAR*12          NAME OF USER SPECIFIED TOPOGRAPHIC
'                                    DATA FILE TO EXTRACT ELEVATIONS
'                                    FROM; IF THIS VARIABLE IS BLANK,
'                                    TNAME WILL RETURN THE NAME OF A
'                                    TOPOGRAPHIC DATA FILE CONTAINING
'                                    DATA FOR THE SPECIFIED COORDINATES
' NAMLST ARGUMENT   CHAR*12(10)      LIST OF THE NAMES OF THE TOPOGRAPHIC
'                                    DATA FILES USED TO GENERATE THE
'                                    TERRAIN PROFILE
' COUNT  ARGUMENT   INTEGER          THE NUMBER OF TOPOGRAPHIC DATA FILES
'                                    USED TO GENERATE THE TERRAIN PROFILE
' DUMMY1 ARGUMENT   INTEGER          DUMMY
'                                       - INCLUDED HERE SO THAT THE ARGUMENT
'                                         LIST IS IDENTICAL TO THAT OF THE
'                                         UNISYS VERSION OF GetElevation
'                                       - THIS DUMMY ARGUMENT IS IN PLACE
'                                         OF THE LONGITUDE SPACING ARGUMENT
'                                         TO THE UNISYS GetElevation ROUTINE
' DATUM  ARGUMENT   INTEGER          GEODETIC DATUM (0 TO 13)
'                                    0 = WGS 84 NO CONVERSION
'                                    1 = NORTH AMERICAN DATUM NAD 27
'                                    2 = EUROPEAN
'                                    3 = TOKYO
'                                    4 = GREAT BRITAIN
'                                    5 = MAUI (OLD HAWAIIAN)
'                                    6 = OAHU (OLD HAWAIIAN)
'                                    7 = KAUAI (OLD HAWAIIAN)
'                                    8 = KWAJALEIN ATOLL (WAKE-ENIWETOK)
'                                    9 = WAKE ISLAND (WAKE-ENIWETOK)
'                                   10 = ENIWETOK ATOLL (WAKE-ENIWETOK)
'                                   11 = WAKE ISLAND ASTRO 1952
'                                   12 = GUAM 1963
'                                   13 = WGS - 72
'
'-----------------------------------------------------------------------------
'
' *** OUTPUT VARIABLES
'
' NAME   LOCATION   TYPE             DESCRIPTION
' ----   --------   ----             -----------
' NSSP   ARGUMENT   INTEGER          LATITUDE SPACING (SECONDS).  IN NAMED
'                                    TOPOGRAPHIC DATA FILE AROUND REQUESTED
'                                    POINT.
' WOTRER ARGUMENT   INTEGER          ERROR RETURN CODE
'
'                WOTRER                  DESCRIPTION
'                -------------------------------------------------------
'                   0              - NO ERROR, NORMAL RETURN
'                   1              - ERROR OPENING TOPO.INF FILE
'                   2              - ERROR READING TOPO.INF FILE
'                   3              - ERROR OPENING TOPO DATA FILE
'                   4              - NO DATA FILE CONTAINS SPECIFIED
'                                    COORDINATES
'                   5              - CANNOT FIND DATA RECORD ENTRY
'                   6              - THE SPECIFIED DATA FILE DOES NOT
'                                    CONTAIN THE SPECIFIED COORDINATES
'                   7              - ERROR IN DATUM CONVERSION ROUTINE
'                   8              - INVALID LATITUDE
'                   9              - INVALID LONGITUDE
'                  10              - MISSING TOPOGRAPHIC ELEVATION DATA
'                  11              - MISSING TOPOGRAPHIC DATA RECORD
'                  12              - ERROR READING FROM DISK
'                  13              - CANNOT FIND DIRECTORY ENTRY
'
' ELEV   ARGUMENT   REAL             ELEVATION (IN METERS) OF SPECIFIED SITE
'
'-----------------------------------------------------------------------------
'

