Attribute VB_Name = "modRegexp"
Option Explicit

#Const KEEP_4_DIGIT_YEAR = True

#If Win32 Then
    Declare Function PRegExpMatch Lib "rexp32.dll" (ByVal PCode As String, ByVal nLength As Integer, ByVal Data As String) As Integer
    Declare Function SRegExpMatch Lib "rexp32.dll" (ByVal Pattern As String, ByVal Data As String) As Integer
    Declare Function PRegExpSubMatch Lib "rexp32.dll" (ByVal PCode As String, ByVal nLength As Integer, ByVal Data As String, _
                                                        ByVal nStart As Integer, ByRef nFoundIdx As Integer, _
                                                        ByRef nFoundLen As Integer) As Integer
    Declare Function SRegExpSubMatch Lib "rexp32.dll" (ByVal Pattern As String, ByVal Data As String, _
                                                        ByVal nStart As Integer, ByRef nFoundIdx As Integer, _
                                                        ByRef nFoundLen As Integer) As Integer
    Declare Function GetRegExpPCode Lib "rexp32.dll" (ByVal Pattern As String, ByVal Buff As String) As Integer
#Else
    Declare Function PRegExpMatch Lib "rexp16.dll" (ByVal PCode As String, ByVal nLength As Integer, ByVal Data As String) As Integer
    Declare Function SRegExpMatch Lib "rexp16.dll" (ByVal Pattern As String, ByVal Data As String) As Integer
    Declare Function PRegExpSubMatch Lib "rexp16.dll" (ByVal PCode As String, ByVal nLength As Integer, ByVal Data As String, _
                                                        ByVal nStart As Integer, ByRef nFoundIdx As Integer, _
                                                        ByRef nFoundLen As Integer) As Integer
    Declare Function SRegExpSubMatch Lib "rexp16.dll" (ByVal Pattern As String, ByVal Data As String, _
                                                        ByVal nStart As Integer, ByRef nFoundIdx As Integer, _
                                                        ByRef nFoundLen As Integer) As Integer
    Declare Function GetRegExpPCode Lib "rexp16.dll" (ByVal Pattern As String, ByVal Buff As String) As Integer
#End If

'Global Const szDatePattern As String = "^((\0?\1[/\-]([1-9]|\0[1-9]|[1-2]#|\3[0-1]))|(\0?\2[/\-]([1-9]|[0][1-9]|\1#|\2[0-9]))|(\0?\3[/\-]([1-9]|0[1-9]|[1-2]#|\3[0-1]))|" & _
'                          "(\0?\4[/\-]([1-9]|0[1-9]|[1-2]#|\3\0))|(\0?\5[/\-]([1-9]|0[1-9]|[1-2]#|\3[0-1]))|(\0?\6[/\-]([1-9]|0[1-9]|[1-2]#|\3\0))|" & _
'                          "(\0?\7[/\-]([1-9]|0[1-9]|[1-2]#|\3[0-1]))|(\0?\8[/\-]([1-9]|0[1-9]|[1-2]#|\3[0-1]))|(\0?\9[/\-]([1-9]|0[1-9]|[1-2]#|\3\0))|" & _
'                          "(\1\0[/\-]([1-9]|0[1-9]|[1-2]#|\3[0-1]))|(\1\1[/\-]([1-9]|0[1-9]|[1-2]#|\3\0))|(\1\2[/\-]([1-9]|0[1-9]|[1-2]#|\3[0-1])))" & _
'                          "[/\-]((\1\9##|\2[0-5]##)|([1-9]#))$"

'pattern resource constants
Public Const szPAT_DESCRIPTION10 = 100      ' any 10
Public Const szPAT_DESCRIPTION20 = 150      ' any 20
Public Const szPAT_DESCRIPTION30 = 200      ' any 30
Public Const szPAT_DESCRIPTION40 = 250      ' any 40
Public Const szPAT_ALLMINUSLOWER = 300      ' anything except for lowercase
Public Const szPAT_ANYTHINGUPPER5 = 305     ' any 5 except lowercase
Public Const szPAT_ANYTHINGUPPER10 = 310    ' any 10 except lowercase
Public Const szPAT_ANYTHINGUPPER16 = 314    ' any 16 except lowercase
Public Const szPAT_ANYTHINGUPPER20 = 315    ' any 20 except lowercase
Public Const szPAT_PUMPCLASS = 350          ' rs_pumpcls.rspc_class
Public Const szPAT_DECIMAL8_4 = 400         ' decimal 8,4
Public Const szPAT_DECIMAL8_2 = 450         ' decimal 8,2
Public Const szPAT_PHONE = 500              ' phone number
Public Const szPAT_ZIPCODE = 550            ' zipcode
Public Const szPAT_PRODUCTCODE = 600        ' inv_header.ivh_product
Public Const szPAT_TAXUSEGROUP = 650        ' tx_use_g.tu_use_g
Public Const szPAT_UOM = 305                ' rs_pump.rspf_uo_measure
Public Const szPAT_TANKNUMBER = 305         ' rs_pump.rspf_tank_nbr

Global Const szNumbers As String = "#"
Global Const szPrintable As String = "P"
Global Const szBasisChar As String = "^[GNgn]$"                   '1 char for net or gross - Lee
Global Const szSuperFundPat As String = "^(#?#)?\.##{1,4}$"     'for superfund on terminal screens - Lee
Global Const szFreightPat As String = "^#{0,3}\.##{0,5}$"       'for freight/surcharge amounts - Lee
Global Const szUOMConversionPat As String = "^#{0,6}(\.#{0,7})?$"  'for UOM to UOM conversion - Lee
Global Const szTermOperatorPat As String = "^[=+0]$"            '1 char sys_term operators - Lee
Global Const szTermRatePat As String = "^#{0,2}(\.#{0,2})?$"    'for UOM to UOM conversion - Lee
Global Const szTermDaysPat As String = "^##{0,2}$"              'for UOM to UOM conversion - Lee
Global Const szAnyFiveDigitNoPat As String = "^#{0,5}$"         'for profit center number - Lee
Global Const szYorNPat As String = "^[YN]$"                     '1 char Y or N - Lee
Global Const szPrftCtrTypPat As String = "^[WRBA]$"             '1 char profit center type - Lee
Global Const szPrftCtrFuelIndPat As String = "^[FN]$"           '1 Fuel Indicator - Lee
Global Const szPrftCtrOwnerPat As String = "^[SCPE]$"           '1 char profit center owner - Lee
Global Const szPrftCtrUnitPat As String = "^[DEN]$"             '1 char profit center unit - Lee
Global Const szPrftCtrUnitQntyPat As String = "^[SE]$"          '1 char profit center unity quantity - Lee
Global Const szPrftCtrReadingPat As String = "^[IVBN]$"         '1 char profit center reading - Lee
Global Const szPrftCtrDailyShftRptPat As String = "^(-\1|#?#)$"  'field can be -1 through 99 used in syfprftc - Lee
Global Const szPrftCtrRptFreqPat As String = "^[DWBSMC]$"       '1 char profit center Report Frequenty - Lee
Global Const szPhoneNumber As String = "^(\(#3\)[ \-]|#3-)?#3-#4$" 'standard USA phone number
Global Const szZipCode As String = "^P{0,10}$"   ' "^#5(\-#4)?$"                   'standard USA ZipCode

Public Const szDatePattern As String = "^(((\0[1-9]|\1[0-2])(\0[1-9]|[1-2]#)|(\0\1|\0[3-9]|\1[0-2])\3\0|(\0(\1|\3|\5|\7|\8)|\1\0|\1\2)\3\1)|((\0?[1-9]|\1[0-2])[/\-](\0?[1-9]|[1-2]#)[/\-]|(\0?\1|\0?[3-9]|\1[0-2])[/\-]\3\0[/\-]|\0?(\1|\3|\5|\7|\8|\1\0|\1\2)[/\-]\3\1[/\-]))((\1\8\9\9|\1\9##|\2[0-5]##)|##)$"
Public Const szIntegerPattern As String = "^(#{0,4}|[0-2]#{0,4}|\3([0-1]#{0,3}|\2([0-6]#?#?|\7([0-5]#?|\6[0-7]?)?)?)?)$"
Public Const szLongPattern As String = "^(#{0,9}|[0-1]#{0,9}|\2(\0#{0,8}|\1([0-3]#{0,7}|\4([0-6]#{0,6}|\7([0-3]#{0,5}|\4([0-7]#{0,4}|\8([0-2]#{0,3}|\3([0-5]#?#?|\6([0-3]#?|\4[0-7]?)?)?)?)?)?)?)?)?)$"

'david 01/03/2001
Public Const szDateTimeToMinutePattern = "^(((\0[1-9]|\1[0-2])(\0[1-9]|[1-2]#)|(\0\1|\0[3-9]|\1[0-2])\3\0|(\0(\1|\3|\5|\7|\8)|\1\0|\1\2)\3\1)|((\0?[1-9]|\1[0-2])[/\-](\0?[1-9]|[1-2]#)[/\-]|(\0?\1|\0?[3-9]|\1[0-2])[/\-]\3\0[/\-]|\0?(\1|\3|\5|\7|\8|\1\0|\1\2)[/\-]\3\1[/\-]))((\1\9##|\2[0-5]##)|##)(((\ )(([01][0-9])|(\2[0-3]))((:[0-5][0-9])|([0-5][0-9])))?)$"
Public Const szDateTimeToSecondPattern = "^(((\0[1-9]|\1[0-2])(\0[1-9]|[1-2]#)|(\0\1|\0[3-9]|\1[0-2])\3\0|(\0(\1|\3|\5|\7|\8)|\1\0|\1\2)\3\1)|((\0?[1-9]|\1[0-2])[/\-](\0?[1-9]|[1-2]#)[/\-]|(\0?\1|\0?[3-9]|\1[0-2])[/\-]\3\0[/\-]|\0?(\1|\3|\5|\7|\8|\1\0|\1\2)[/\-]\3\1[/\-]))((\1\9##|\2[0-5]##)|##)(((\ )(([01][0-9])|(\2[0-3]))((:[0-5][0-9])|([0-5][0-9]))((:[0-5][0-9])|([0-5][0-9])))?)$"

'valid format ( _ is space)
'yyyy-mm-dd[[_hh:mm:ss][_hh][hh:mm][hhmm][_hhmmss][_hh:mmss][_hhmm:ss]]
Public Const szLongDateTimeToSecondPattern As String = "^(((\1\9##)|(\2###))\-((\0[13578]\-((\0[1-9])|([1-2]#)|(\3[01])))|(\0[469]\-((\0[1-9])|([1-2]#)|(\3\0)))|(\1[02]\-((\0[1-9])|([1-2]#)|(\3[01])))|(\1\1\-((\0[1-9])|([1-2]#)|(\3\0)))|(\0\2\-((\0[1-9])|(\1#)|(\2[0-9]))))((((\ )((([01][0-9])|(\2[0-3]))((:[0-5][0-9])|([0-5][0-9]))?((:[0-5][0-9])|([0-5][0-9]))?)))?))$"

Public Const FMT_DATE_SHORT = "MM/DD/YY"
Public Const FMT_DATE_LONG = "MM/DD/YYYY"

Public Const DEFAULT_DECIMALS = 6
Private Function tfnFixedDecimalPattern(nNumDigits As Integer, _
                                        nNumDecimals As Integer, _
                                        Optional ByVal vNegAllowed As Variant) As String
    Dim szPattern As String
    Dim nWhole As Integer
    Dim n As Integer
    
    nWhole = nNumDigits - nNumDecimals
    szPattern = szPattern & "(#{0," & CStr(nWhole) & "}"
    If nNumDecimals > 0 Then
        szPattern = szPattern & "\.#{0," & CStr(nNumDecimals) & "}"
    End If
    If nWhole > 3 Then
        szPattern = szPattern & tfnPtnDecimals(nWhole, nNumDecimals)
        If nWhole > 6 Then
            If (nWhole Mod 3) = 1 Then
                szPattern = szPattern & tfnPtnDecimals(nWhole - 1, nNumDecimals)
            Else
                szPattern = szPattern & tfnPtnDecimals(nWhole - 2, nNumDecimals)
            End If
        End If
    End If
    szPattern = szPattern & ")"

    #If DEVELOP Then
        Clipboard.SetText szPattern, vbCFText
    #End If
    
    tfnFixedDecimalPattern = szPattern


End Function

Public Function tfnFormatDecimal(ByVal vTemp As Variant, _
                                 Optional ByVal nDecimals As Variant, _
                                 Optional ByVal nExtras As Variant, _
                                 Optional ByVal bShowComma As Boolean = True) As String
    Dim sFmt As String
    Dim sReturn As String
    Dim sChar As String
    
    If bShowComma Then
        sFmt = "###,###,###,###,##0"
    Else
        sFmt = "0"
    End If
    
    sChar = "."
    
    If Not IsMissing(nDecimals) Then
        If nDecimals > 0 Then
            sFmt = sFmt & sChar & String(nDecimals, "0")
            sChar = ""
        End If
    End If
    
    If Not IsMissing(nExtras) Then
        If nExtras > 0 Then
            sFmt = sFmt & sChar & String(nExtras, "#")
        End If
    End If
    
    If IsNumeric(vTemp) Then
        tfnFormatDecimal = Format(vTemp, sFmt)
        If Right(tfnFormatDecimal, 1) = "." Then
            If Len(tfnFormatDecimal) > 1 Then
                tfnFormatDecimal = Left(tfnFormatDecimal, Len(tfnFormatDecimal) - 1)
            Else
                tfnFormatDecimal = "0"
            End If
        End If
    Else
        tfnFormatDecimal = Format(0, sFmt)
    End If
End Function

Private Function tfnNormDate(vTemp As String) As Boolean
    'Returns true if there are 2 digit2 in year field
    Dim nPos1 As Integer
    Dim nPos2 As Integer
    Dim nTemp As Integer
    Dim sTemp As String
    
    tfnNormDate = False
    sTemp = CStr(vTemp)
    nPos1 = InStr(vTemp, "/")
    If nPos1 <= 0 Then
        nPos1 = InStr(vTemp, "-")
    End If
    If nPos1 > 0 Then
        nPos2 = InStr(nPos1 + 1, vTemp, "/")
        If nPos2 <= nPos1 Then
            nPos2 = InStr(nPos1 + 1, vTemp, "-")
        End If
        If nPos1 <= 3 Then
            If IsNumeric(Left(sTemp, nPos1 - 1)) Then
                nTemp = nPos2 - nPos1 - 1
                If nTemp = 1 Or nTemp = 2 Then
                    If IsNumeric(Mid(sTemp, nPos1 + 1, nTemp)) Then
                        nTemp = Len(sTemp) - nPos2
                        If nTemp = 2 Then
                            If IsNumeric(Mid(sTemp, nPos2 + 1, nTemp)) Then
                                tfnNormDate = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function tfnPtnDecimals(nWhole As Integer, nNumDecimals As Integer) As String
    Dim nMod3 As Integer
    Dim nTime3 As Integer
    Dim szPattern As String
'    Dim nWhole As Integer
    Dim sStart As String
    Dim sPtn() As String
    Dim n As Integer
    
    nMod3 = nWhole Mod 3
    nTime3 = nWhole \ 3
    If nMod3 = 0 Then
        nMod3 = 3
        nTime3 = nTime3 - 1
    End If
    sStart = "|#"
    While nMod3 > 1
        sStart = sStart & "#?"
        nMod3 = nMod3 - 1
    Wend
    n = nTime3
    ReDim sPtns(n)
    sPtns(n) = ",?###"
    While n > 0
        n = n - 1
        sPtns(n) = sPtns(n + 1) & sPtns(nTime3)
    Wend
    szPattern = ""
    For n = nTime3 To 1 Step -1
        szPattern = szPattern & sStart & sPtns(n)
        If nNumDecimals > 0 Then
            szPattern = szPattern & "\." & "#{0," & CStr(nNumDecimals) & "}"
        End If
    Next

    tfnPtnDecimals = szPattern
End Function

Public Function tfnFormatDate(ByVal vSource As Variant, _
                              Optional vKeepCentury As Variant) As String
    'When vKeepCentury is true. It means that if the passed date has only two digits
    '  in year, the default century is the current century. 50 years rule will not be
    '  used. This option is put here because some the the dates in the database
    '  are really old (such as those in freight rate table).
    
    Const DATE_MONTH_LEN = 2
    Const MONTH_START = 3
    Const YEAR_START = 5
    Const MIN_LEN = 6
    Const DIVIDER = "/"
    
    Dim nPos1 As Integer
    Dim nPos2 As Integer
    Dim sText As String
    Dim sMonth As String
    Dim sDay As String
    Dim sYear As String
    Dim sFmt As String
    Dim nYearLen As Integer
    Dim dTemp As Date
    Dim bKeepCentury As Boolean
    
    tfnFormatDate = ""
    If IsNull(vSource) Then
        Exit Function
    End If
    sText = Trim(vSource)
    tfnFormatDate = sText
    If Len(sText) >= MIN_LEN Then
        nPos1 = InStr(sText, DIVIDER)
        If nPos1 = 0 Then
            nPos1 = InStr(sText, "-")
        ElseIf nPos1 > MONTH_START Then
            nPos2 = nPos1
            nPos1 = InStr(sText, "-")
        End If
        If nPos1 > 0 Then
            If nPos1 <= MONTH_START Then
                nPos2 = InStr(nPos1 + 1, sText, DIVIDER)
                If nPos2 = 0 Then
                    nPos2 = InStr(nPos1 + 1, sText, "-")
                End If
            Else
                nPos2 = nPos1
                nPos1 = 0
            End If
        End If
        If nPos2 + nPos1 = 0 Then
            nYearLen = Len(sText) - 4
        ElseIf nPos2 = 0 Then
            nYearLen = Len(sText) - nPos1 - 2
        Else
            nYearLen = Len(sText) - nPos2
        End If
        sMonth = Left(sText, DATE_MONTH_LEN)
        If nPos1 = 0 Then
            If nPos2 = 0 Then
                sDay = Mid(sText, MONTH_START, DATE_MONTH_LEN)
            Else
                sDay = Mid(sText, MONTH_START, nPos2 - MONTH_START)
            End If
        ElseIf nPos1 > MONTH_START Then
            sDay = Mid(sText, MONTH_START, nPos1 - MONTH_START)
        Else
            If nPos2 > nPos1 Then
                sDay = Mid(sText, nPos1 + 1, nPos2 - nPos1 - 1)
            Else
                sDay = Mid(sText, nPos1 + 1, DATE_MONTH_LEN)
            End If
            If nPos1 < MONTH_START And nPos1 > 1 Then
                sMonth = Left(sText, nPos1 - 1)
            End If
        End If
        While Len(sDay) < 2
            sDay = "0" & sDay
        Wend
        While Len(sMonth) < 2
            sMonth = "0" & sMonth
        Wend
        If nPos2 = 0 Then
            If nPos1 > MONTH_START Then
                sYear = Mid(sText, nPos1 + 1, nYearLen)
            Else
                If nPos1 > 0 Then
                    sYear = Mid(sText, nPos1 + DATE_MONTH_LEN + 1, nYearLen)
                Else
                    sYear = Mid(sText, YEAR_START, nYearLen)
                End If
            End If
        Else
            sYear = Mid(sText, nPos2 + 1, nYearLen)
        End If
        sFmt = FMT_DATE_SHORT
        If nYearLen = 4 Then
            #If KEEP_4_DIGIT_YEAR Then
                'Always keep 4-digit-year
                sFmt = FMT_DATE_LONG
            #Else
                'Keep 4-digit-year only if the centuries are different
                nPos1 = Year(Date) \ 100     'This century
                nPos2 = val(sYear) \ 100     'Input century
                If nPos1 <> nPos2 Then
                    sFmt = FMT_DATE_LONG
                End If
            #End If
        ElseIf nYearLen <> 2 Then
            tfnFormatDate = sText
            Beep
            Exit Function
        End If
        'dTemp = Format(sMonth & DIVIDER & sDay & DIVIDER & sYear, sFmt)
        If IsMissing(vKeepCentury) Then
            bKeepCentury = False
        Else
            bKeepCentury = vKeepCentury
        End If
        If bKeepCentury Or nYearLen = 4 Then
            'Keep the century if it is against 50 years rule, otherwise, drop it
            'Check whether it is against 50 years rule.
            nPos1 = Year(Date)      'This year
            nPos2 = val(sYear)     'Input year
            If Abs(nPos1 - nPos2) >= 50 Then
                sFmt = FMT_DATE_LONG
            Else
                sFmt = FMT_DATE_SHORT
            End If
        End If
        tfnFormatDate = Trim(Format(sMonth & DIVIDER & sDay & DIVIDER & sYear, sFmt))
    End If
End Function


Public Function tfnDateString(vDate As Variant, _
                              Optional ByVal vQuote As Variant, _
                              Optional ByVal vBackYears As Variant) As String
'This function returns a string type variable for SQL statement,
'The year has been converted from the 2-digit year(input) to 4-digits,
'so that it considers the change of century. The default number of
'years going back is 50. The caller can overwrite the years going back
'by supplying the function the number of years going back as the third
'parameter

    Const YEARS_BACK_1 = 50
    Const LEN_STANDARD = 8
    Const LEN_YEAR = 2
    
    Dim nYearsLimit As Integer
    Dim nYearsDiff As Integer
    Dim nCurrYear As Integer
    Dim nTemp As Integer
    Dim nYears As Integer
    Dim sText As String
    
    nCurrYear = Year(Date)
    nYearsLimit = YEARS_BACK_1
    If Not IsMissing(vBackYears) Then
        nYearsLimit = vBackYears
    End If
    If IsNull(vDate) Then
        sText = ""
    Else
        sText = vDate
        If tfnNormDate(sText) Then
            sText = Format(sText, FMT_DATE_SHORT)
            nYears = tfnYear(sText)
            nYearsDiff = nCurrYear Mod 100 - nYears Mod 100
            If nYearsDiff < 0 Then
                nYearsLimit = 100 - nYearsLimit
                If Abs(nYearsDiff) <= nYearsLimit Then
                    nYears = nCurrYear - nYearsDiff
                Else
                    nYears = nCurrYear - nYearsDiff - 100
                End If
            Else
                If Abs(nYearsDiff) < nYearsLimit Then
                    nYears = nCurrYear - nYearsDiff
                Else
                    nYears = nCurrYear - nYearsDiff + 100
                End If
            End If
            nTemp = Len(sText) - LEN_YEAR
            sText = Left(Trim(sText), nTemp) & CStr(nYears)
        End If
    End If
    If IsMissing(vQuote) Then
        tfnDateString = sText
    Else
        If vQuote Then
            If sText = "" Then
                tfnDateString = "Null"
            Else
                tfnDateString = "'" & sText & "'"
            End If
        Else
            tfnDateString = sText
        End If
    End If
End Function


Private Function tfnYear(sText As String) As Integer
    Dim i As Integer
    Dim sChar As String * 1
    Dim sYear As String
    
    i = Len(sText)
    sYear = ""
    Do
        sChar = Mid(sText, i, 1)
        If sChar <> "/" And sChar <> "-" Then
            sYear = sChar & sYear
        Else
            Exit Do
        End If
        i = i - 1
    Loop Until i <= 1
    If Len(sYear) <= 4 Then
        tfnYear = val(sYear)
    End If
End Function

Public Function tfnDecimalPattern(nNumDigits As Integer, _
                                  nNumDecimals As Integer, _
                                  Optional ByVal vNegAllowed As Variant, _
                                  Optional ByVal vFloating As Variant) As String
'
' Generates a pattern for an nNumDigit floating point number and returns a PCode pattern

    Dim szPattern As String
    Dim sTemp As String
    Dim nWhole As Integer
    Dim bVarLen As Integer
    Dim n As Integer
    
    If nNumDigits <= 0 Then
        Exit Function
    End If
    If nNumDecimals >= nNumDigits Then
        nNumDecimals = nNumDigits - 1
    End If

    szPattern = "^("
    If Not IsMissing(vNegAllowed) Then
        If vNegAllowed Then
            szPattern = "^[\-+]?("
        End If
    End If
    
    If IsMissing(vFloating) Then
        bVarLen = False
    Else
        bVarLen = vFloating
    End If
    
    If Not bVarLen Then
        szPattern = szPattern & tfnFixedDecimalPattern(nNumDigits, nNumDecimals, vNegAllowed)
    Else
        For n = 0 To nNumDecimals - 1
            szPattern = szPattern & tfnFixedDecimalPattern(nNumDigits, n, vNegAllowed) & "|"
        Next
        szPattern = szPattern & tfnFixedDecimalPattern(nNumDigits, n, vNegAllowed)
    End If
    
    #If DEVELOP Then
        Clipboard.SetText szPattern, vbCFText
    #End If
    
    tfnDecimalPattern = szPattern & ")$"

End Function

' tfnRegExpControlKeyPress
'
' checks the text of the control against the supplied pattern
' and nullifies the keypress if the character is bad
'
' cntl      - Control to process
' KeyAscii  - The key code from the controls KeyPress event
' szPattern - The pattern string (could be PCode)
'
' Returns True if the key is good, False if bad (also nullified key)

Public Function tfnRegExpControlKeyPress(ByRef cntl As Control, ByRef KeyAscii As Integer, ByRef szPattern As String) As Boolean

    Dim nCode As Integer
    Dim szData As String
    Dim nLen As Integer
    
    If cntl Is Nothing Then
        tfnRegExpControlKeyPress = False
        Exit Function
    End If
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Then
        tfnRegExpControlKeyPress = True
        Exit Function
    End If
    
    If TypeOf cntl Is Textbox Or TypeOf cntl Is ComboBox Then
        ' check for cut/copy/paste keys
        If KeyAscii = vbKeyCancel Or KeyAscii = &H16 Or KeyAscii = &H18 Then
            tfnRegExpControlKeyPress = True
            Exit Function
        End If
        
        ' get the data up to the cursor position and tack on the key pressed
        szData = Left(cntl.Text, cntl.SelStart) & Chr(KeyAscii)
        nLen = Len(cntl.Text)
        nCode = nLen - (cntl.SelStart + cntl.SelLength)
        If nCode > 0 Then
            szData = szData & Right(cntl.Text, nCode)
        End If
        If cntl.SelLength > 0 Then
            nLen = nLen - cntl.SelLength
        End If
        If InStrB(szPattern, Chr(0)) <> 0 Then
            nCode = PRegExpMatch(szPattern, Len(szPattern), szData)
        Else
            nCode = SRegExpMatch(szPattern, szData)
        End If
        
        If nCode = 0 Or nCode = nLen + 2 Then
            tfnRegExpControlKeyPress = True
        Else
            KeyAscii = 0
            tfnRegExpControlKeyPress = False
            Beep
        End If
    Else
        tfnRegExpControlKeyPress = True
    End If

End Function


' tfnRegExpControlChange
'
' checks the text of the control against the supplied pattern
' and highlights bad input
'
' cntl      - Control to process
' szPattern - The pattern string (could be PCode)
'
' Returns True if entire text is okay - False if bad or blank
' If the text is bad, the bad portion will be highlighted

Public Function tfnRegExpControlChange(ByRef cntl As Control, ByRef szPattern As String) As Boolean

    Dim nCode As Integer
    
    If cntl Is Nothing Then
        tfnRegExpControlChange = False
        Exit Function
    End If
    
    If TypeOf cntl Is Textbox Or TypeOf cntl Is ComboBox Then
        If cntl.Text = "" Then
            tfnRegExpControlChange = True
            Exit Function
        End If
        
        If InStrB(szPattern, Chr(0)) <> 0 Then
            nCode = PRegExpMatch(szPattern, Len(szPattern), cntl.Text)
        Else
            nCode = SRegExpMatch(szPattern, cntl.Text)
        End If
        
        tfnRegExpControlChange = (nCode = 0)    ' true if text is okay
        
        ' highlight any bad text
        If nCode > 0 Then
            cntl.SelStart = nCode - 1
            cntl.SelLength = Len(cntl.Text) - cntl.SelStart
            If cntl.SelLength > 0 Then
                Beep
            End If
        End If
    
    Else
        tfnRegExpControlChange = True
    End If
    
End Function

Public Function tfnRegExpControlGotFocus(ByRef cntl As Control, ByRef szPattern As String) As Boolean
' checks the text of the control against the supplied pattern
' and highlights bad input

    tfnRegExpControlGotFocus = tfnRegExpControlChange(cntl, szPattern)

End Function

' tfnRegExpControlDateKeyPress
'
' checks the text of the control against the date pattern,
' nullifies the keypress if the character is bad, and returns zero when
' the entire pattern is complete
'
' cntl      - Control to process
' KeyAscii  - The key code from the controls KeyPress event
'
' Returns zero when the pattern has been matched, next position to be matched otherwise, and the
' nullified key if the key was bad

Public Function tfnRegExpControlDateKeyPress(ByRef cntl As Control, ByRef KeyAscii As Integer) As Integer

    Dim nCode As Integer
    Dim szData As String
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Then
        tfnRegExpControlDateKeyPress = -1
        Exit Function
    End If
    
    If TypeOf cntl Is Textbox Or TypeOf cntl Is ComboBox Then
        ' check for cut/copy/paste keys
        If KeyAscii = vbKeyCancel Or KeyAscii = &H16 Or KeyAscii = &H18 Then
            tfnRegExpControlDateKeyPress = -1
            Exit Function
        End If
        
        ' get the data up to the cursor position and tack on the key pressed
        szData = Left(cntl.Text, cntl.SelStart) & Chr(KeyAscii)
        
        If InStrB(szDatePattern, Chr(0)) <> 0 Then
            nCode = PRegExpMatch(szDatePattern, Len(szDatePattern), szData)
        Else
            nCode = SRegExpMatch(szDatePattern, szData)
        End If
        
        If nCode < 0 Or (nCode <> 0 And nCode <> cntl.SelStart + 2) Then
            KeyAscii = 0
            tfnRegExpControlDateKeyPress = nCode
            Beep
        Else
            tfnRegExpControlDateKeyPress = nCode
        End If
    End If

End Function

'david 12/28/00
'S - to second
'M - to minute
Public Function tfnFormatDateTime(sDateTime As String, sToMinuteOrSecond As String) As String
    Dim sDate As String, sTime As String, nPosi As Integer
    Dim sFormatedTime As String
    
    tfnFormatDateTime = ""
    sDateTime = Trim(sDateTime)
    
    If sDateTime = "" Then Exit Function
    
    nPosi = InStr(sDateTime, " ")
    
    If nPosi > 0 Then
        sDate = Left(sDateTime, nPosi - 1)
        
        If InStr(sDate, "-") > 3 Then
            sDate = Format(sDate, "yyyy-mm-dd")
        Else
            sDate = tfnFormatDate(sDate)
        End If
        
        If Not IsDate(tfnDateString(sDate)) Then
            tfnFormatDateTime = sDateTime
            Exit Function
        End If
        
        sTime = Mid(sDateTime, nPosi + 1)
        
        sFormatedTime = fnFormatTime(sTime, sToMinuteOrSecond)
        
        If sFormatedTime = "" Then
            tfnFormatDateTime = sDateTime
        Else
            tfnFormatDateTime = sDate & " " & sFormatedTime
        End If
    Else
        sDate = sDateTime
        tfnFormatDateTime = tfnFormatDate(sDate) + " " _
            + IIf(sToMinuteOrSecond = "M", "00:00", "00:00:00")
    End If
End Function

'sTime is in the format of (hh:mm[:ss])
'S - to second
'M - to minute
Private Function fnFormatTime(ByVal sTime As String, sToMinuteOrSecond As String) As String
    Dim nPosi As Integer
    Dim sTemp As String
    Dim sHH As String
    Dim sMM As String
    Dim sSS As String
    
    sToMinuteOrSecond = UCase(sToMinuteOrSecond)
    
    fnFormatTime = ""
    
    sTime = Trim(sTime)
    
    If sTime = "" Then Exit Function
    
    If Len(sTime) < 2 Or Len(sTime) > 8 Then
        If Len(sTime) > 8 And (UCase(Right(sTime, 2)) = "AM" Or UCase(Right(sTime, 2)) = "PM") Then
            sTime = Format(sTime, "hh:mm:ss")
        Else
            Exit Function
        End If
    End If
    
    nPosi = InStr(sTime, ":")
    
    If nPosi > 0 Then
        'parse hh:mm:ss
        sHH = Left(sTime, nPosi - 1)
        If Len(sHH) > 2 Then
            sMM = Format(Right(sHH, 2), "00")
            sHH = Format(Left(sHH, 2), "00")
        Else  '= 2
            sHH = Format(Left(sTime, nPosi - 1), "00")
        End If
        
        sTemp = Mid(sTime, nPosi + 1)
        
        nPosi = InStr(sTemp, ":")
        
        If nPosi > 0 Then
            sMM = Format(Left(sTemp, nPosi - 1), "00")
            sSS = Format(Mid(sTemp, nPosi + 1), "00")
        Else
            If Len(sTemp) > 2 Then
                sMM = Format(Left(sTemp, 2), "00")
                sSS = Format(Right(sTemp, 2), "00")
            Else  '= 2
                sMM = Format(sTemp, "00")
                sSS = "00"
            End If
        End If
    Else
        If Len(sTime) Mod 2 <> 0 Then
            Exit Function
        End If
        
        Select Case Len(sTime)
            Case 2
                sHH = Format(sTime, "00")
                sMM = "00"
                sSS = "00"
            Case 4
                sHH = Format(Left(sTime, 2), "00")
                sMM = Format(Right(sTime, 2), "00")
                sSS = "00"
            Case 6
                sHH = Format(Left(sTime, 2), "00")
                sMM = Format(Mid(sTime, 3, 2), "00")
                sSS = Format(Right(sTime, 2), "00")
        End Select
    End If

    If val(sHH) > 23 Or val(sMM) > 59 Or val(sSS) > 59 Then
        Exit Function
    Else
        fnFormatTime = sHH + ":" & sMM
        
        If sToMinuteOrSecond = "S" Then
            fnFormatTime = fnFormatTime + ":" & sSS
        End If
    End If
End Function

'sYearTo argument required only when bUsedInWhereClause = True
'Y-year to year
'M-year to month
'D-year to day
'H-year to hour
'N-year to minute
'S-year to second
Public Function tfnDateTimeString(ByVal sDateTime As String, _
                                  Optional sYearTo As String = "S", _
                                  Optional bUsedInWhereClause As Boolean = False) As String

    Dim sUsedDateTime As String
    Dim sQualifier As String
    Dim sYYYY As String
    Dim sMO As String
    Dim sDD As String
    Dim sHH As String
    Dim sMN As String
    Dim sSS As String
    
    sYearTo = UCase(sYearTo)
    
    tfnDateTimeString = "''"
    sDateTime = Trim(sDateTime)
    
    If sDateTime = "" Then
        Exit Function
    End If
    
    If Not fnParseDateTime(sDateTime, sYYYY, sMO, sDD, sHH, sMN, sSS) Then
        Exit Function
    End If
    
    Select Case sYearTo
    Case "Y"
        sUsedDateTime = fnSQLString(sYYYY)
        sQualifier = "year to year"
    Case "M"
        sUsedDateTime = fnSQLString(sYYYY + "-" + sMO)
        sQualifier = "year to month"
    Case "D"
        sUsedDateTime = fnSQLString(sYYYY + "-" + sMO + "-" + sDD)
        sQualifier = "year to day"
    Case "H"
        sUsedDateTime = fnSQLString(sYYYY + "-" + sMO + "-" + sDD + " " + sHH)
        sQualifier = "year to hour"
    Case "N"
        sUsedDateTime = fnSQLString(sYYYY + "-" + sMO + "-" + sDD + " " + sHH + ":" + sMN)
        sQualifier = "year to minute"
    Case "S"
        sUsedDateTime = fnSQLString(sYYYY + "-" + sMO + "-" + sDD + " " + sHH + ":" + sMN + ":" + sSS)
        sQualifier = "year to second"
    End Select
    
    If bUsedInWhereClause Then
        tfnDateTimeString = "extend(" + sUsedDateTime + "," + sQualifier + ")"
    Else
        tfnDateTimeString = "{ts '" + sYYYY + "-" + sMO + "-" + sDD + " " + sHH + ":" + sMN + ":" + sSS + "'}"
    End If
    
End Function

Public Function fnParseDateTime(ByVal sDateTime As String, _
                                sYYYY As String, _
                                Optional sMO As String = "", _
                                Optional sDD As String = "", _
                                Optional sHH As String = "", _
                                Optional sMN As String = "", _
                                Optional sSS As String = "") As Boolean

    Const FORMAT_DATE = "YYYY-MM-DD"
    Const FORMAT_TIME = "HH:MM:SS"
    
    Dim nPosi  As Integer
    Dim sDate As String
    Dim sTime As String
    
    sDateTime = tfnFormatDateTime(Trim(sDateTime), "S")

    If sDateTime = "" Then
        Exit Function
    End If
    
    If Not IsDate(sDateTime) Then
        Exit Function
    End If

    On Error GoTo errHandler
    
    nPosi = InStr(sDateTime, " ")
    
    If nPosi > 0 Then
        sDate = Format(Left(sDateTime, nPosi - 1), FORMAT_DATE)
        sTime = Format(Trim(Mid(sDateTime, nPosi + 1)), FORMAT_TIME)
    Else
        sDate = Format(sDateTime, FORMAT_DATE)
        sTime = "00:00:00"
    End If
    
    'parse date into yyyy,mm,dd
    nPosi = InStr(sDate, "-")
    If nPosi <= 0 Then
        Exit Function
    End If
    
    sYYYY = Left(tfnDateString(sDate), 4)
    sDate = Mid(sDate, 6)
    
    nPosi = InStr(sDate, "-")
    If nPosi <= 0 Then
        Exit Function
    End If
    
    sMO = Left(sDate, 2)
    sDD = Right(sDate, 2)
    
    sHH = Left(sTime, 2)
    sMN = Mid(sTime, 4, 2)
    sSS = Right(sTime, 2)
    
    fnParseDateTime = True
    
    Exit Function
    
errHandler:
    
End Function

Private Function fnSQLString(ByVal vTemp As Variant, Optional vNoQuotes As Variant) As String
'
' Properly quotes and formats an SQL string.  If vNoQuotes is present, the result WILL NOT BE QUOTED
' for each ' character found, insert a double ''.  Leave "%* alone
    
    Dim nIdx As Integer
    Dim nPos As Integer
    Dim szParameter As String
    
    If IsNull(vTemp) Then
        szParameter = ""
    Else
        szParameter = vTemp
    End If

    nIdx = 1
    nPos = InStr(nIdx, szParameter, "'")
    
    While nPos <> 0
        szParameter = Left(szParameter, nPos) & "'" & Right(szParameter, Len(szParameter) - nPos)
        nIdx = nPos + 2
        nPos = InStr(nIdx, szParameter, "'")
    Wend
    
    ' quote the whole string - optional
    If IsMissing(vNoQuotes) Then
        fnSQLString = "'" & szParameter & "'"
    Else
        fnSQLString = szParameter
    End If

End Function

