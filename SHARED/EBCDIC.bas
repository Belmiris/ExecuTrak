Attribute VB_Name = "modEBCDIC"
Option Explicit
'***********************************************************'
'
' Copyright (c) 2010 FACTOR, A Division of W.R.Hess Company
'
' Module name       : modEBCDIC (EBCDIC.bas)
'
' Revision          :
'
' Date              : 06/30/2010
'
' Programmer(s)     : Paul Jeanquart
'
' Specification     : Functions Handy to use when dealing with EBCDIC files
'
'Function Description()
'Translate   Converts a string from one character encoding scheme to another.
'   Requires a translation table as one of the arguments.
'ASCII_To_EBCDIC_Table   Returns a string containing the translation table for converting
'   an ASCII string to an EBCDIC string.
'EBCDIC_To_ASCII_Table   Returns a string containing the translation table for converting
'    an EBCDIC string to an ASCII string.
'HexToStr    A helper function that converts a string of hexadecimal digits into the actual
'   characters they represent.
'
'Usage:
'
'  sEBCDIC = Input(#1, 50)     ' input 50 characters
'  sASCII = Translate(sEBCDIC, ASCII_To_EBCDIC_Table())
'
'
' if calling the translation routine repeatedly,
'   you can cache the translation table in a String variable:
'
'  xlat = ASCII_To_EBCDIC_Table()
'  sASCII = Translate(sEBCDIC, xlat)
'
'***********************************************************'


Function Translate(ByVal InText As String, xlatTable As String) As String
'
' Uses a translation table to map InText from one character set to another.
'
Dim Temp As String
Dim I As Long
    Temp = Space$(Len(InText))
    For I = 1 To Len(InText)
        Mid$(Temp, I, 1) = Mid$(xlatTable, Asc(Mid$(InText, I, 1)) + 1, 1)
    Next I
    Translate = Temp
End Function

Function ASCII_To_EBCDIC_Table() As String
'
' Returns the following table as a string for use by the Translate
' function to translate an EBCDIC string to an ASCII-ISO/ANSI string.
'
' 00 01 02 03 37 2D 2E 2F 16 05 25 0B 0C 0D 0E 0F
' 10 11 12 13 3C 3D 32 26 18 19 3F 27 1C 1D 1E 1F
' 40 5A 7F 7B 5B 6C 50 7D 4D 5D 5C 4E 6B 60 4B 61
' F0 F1 F2 F3 F4 F5 F6 F7 F8 F9 7A 5E 4C 7E 6E 6F
' 7C C1 C2 C3 C4 C5 C6 C7 C8 C9 D1 D2 D3 D4 D5 D6
' D7 D8 D9 E2 E3 E4 E5 E6 E7 E8 E9 AD E0 BD 5F 6D
' 79 81 82 83 84 85 86 87 88 89 91 92 93 94 95 96
' 97 98 99 A2 A3 A4 A5 A6 A7 A8 A9 C0 4F D0 A1 07
' 20 21 22 23 24 15 06 17 28 29 2A 2B 2C 09 0A 1B
' 30 31 1A 33 34 35 36 08 38 39 3A 3B 04 14 3E E1
' 41 42 43 44 45 46 47 48 49 51 52 53 54 55 56 57
' 58 59 62 63 64 65 66 67 68 69 70 71 72 73 74 75
' 76 77 78 80 8A 8B 8C 8D 8E 8F 90 9A 9B 9C 9D 9E
' 9F A0 AA AB AC 4A AE AF B0 B1 B2 B3 B4 B5 B6 B7
' B8 B9 BA BB BC 6A BE BF CA CB CC CD CE CF DA dB
' DC DD DE DF EA EB EC ED EE EF FA FB FC FD FE FF
'
  ASCII_To_EBCDIC_Table = _
  HexToStr("00010203372D2E2F1605250B0C0D0E0F101112133C3D322618193F271C1D1E1F") & _
  HexToStr("405A7F7B5B6C507D4D5D5C4E6B604B61F0F1F2F3F4F5F6F7F8F97A5E4C7E6E6F") & _
  HexToStr("7CC1C2C3C4C5C6C7C8C9D1D2D3D4D5D6D7D8D9E2E3E4E5E6E7E8E9ADE0BD5F6D") & _
  HexToStr("79818283848586878889919293949596979899A2A3A4A5A6A7A8A9C04FD0A107") & _
  HexToStr("202122232415061728292A2B2C090A1B30311A333435360838393A3B04143EE1") & _
  HexToStr("4142434445464748495152535455565758596263646566676869707172737475") & _
  HexToStr("767778808A8B8C8D8E8F909A9B9C9D9E9FA0AAABAC4AAEAFB0B1B2B3B4B5B6B7") & _
  HexToStr("B8B9BABBBC6ABEBFCACBCCCDCECFDADBDCDDDEDFEAEBECEDEEEFFAFBFCFDFEFF")
End Function

Function EBCDIC_To_ASCII_Table() As String
'
' Returns the following table as a string for use by the Translate
' function to traslate an EBCDIC string to an ASCII-ISO/ANSI string.
'
' 00 01 02 03 9C 09 86 7F 97 8D 8E 0B 0C 0D 0E 0F    ....ú.Ü-çé.....
' 10 11 12 13 9D 85 08 87 18 19 92 8F 1C 1D 1E 1F    ....ù....á..'è....
' 80 81 82 83 84 0A 17 1B 88 89 8A 8B 8C 05 06 07    ÄÅÇÉ"...àâäãå...
' 90 91 16 93 94 95 96 04 98 99 9A 9B 14 15 9E 1A    ê'.""ï-.ò(tm)öõ..û.
' 20 A0 A1 A2 A3 A4 A5 A6 A7 A8 D5 2E 3C 28 2B 7C    . °¢£§•¶ß...<(+|
' 26 A9 AA AB AC AD AE AF B0 B1 21 24 2A 29 3B 5E    &(c)™"¨≠(r)Ø∞±!$*);^
' 2D 2F B2 B3 B4 B5 B6 B7 B8 B9 E5 2C 25 5F 3E 3F    -/≤≥¥µ∂∑∏π.,%_>?
' BA BB BC BD BE BF C0 C1 C2 60 3A 23 40 27 3D 22    ∫"1/41/23/4ø...`:#@'="
' C3 61 62 63 64 65 66 67 68 69 C4 C5 C6 C7 C8 C9    .abcdefghi......
' CA 6A 6B 6C 6D 6E 6F 70 71 72 CB CC CD CE CF D0    .jklmnopqr......
' D1 7E 73 74 75 76 77 78 79 7A D2 D3 D4 5B D6 D7    .~stuvwxyz...[..
' D8 D9 DA DB DC DD DE DF E0 E1 E2 E3 E4 5D E6 E7    .............]..
' 7B 41 42 43 44 45 46 47 48 49 E8 E9 EA EB EC ED    {ABCDEFGHI......
' 7D 4A 4B 4C 4D 4E 4F 50 51 52 EE EF F0 F1 F2 F3    }JKLMNOPQR......
' 5C 9F 53 54 55 56 57 58 59 5A F4 F5 F6 F7 F8 F9    \.STUVWXYZ......
' 30 31 32 33 34 35 36 37 38 39 FA FB FC FD FE FF    0123456789......
'
  EBCDIC_To_ASCII_Table = _
  HexToStr("000102039C09867F978D8E0B0C0D0E0F101112139D8508871819928F1C1D1E1F") & _
  HexToStr("80818283840A171B88898A8B8C050607909116939495960498999A9B14159E1A") & _
  HexToStr("20A0A1A2A3A4A5A6A7A8D52E3C282B7C26A9AAABACADAEAFB0B121242A293B5E") & _
  HexToStr("2D2FB2B3B4B5B6B7B8B9E52C255F3E3FBABBBCBDBEBFC0C1C2603A2340273D22") & _
  HexToStr("C3616263646566676869C4C5C6C7C8C9CA6A6B6C6D6E6F707172CBCCCDCECFD0") & _
  HexToStr("D17E737475767778797AD2D3D45BD6D7D8D9DADBDCDDDEDFE0E1E2E3E45DE6E7") & _
  HexToStr("7B414243444546474849E8E9EAEBECED7D4A4B4C4D4E4F505152EEEFF0F1F2F3") & _
  HexToStr("5C9F535455565758595AF4F5F6F7F8F930313233343536373839FAFBFCFDFEFF")
End Function

Function HexToStr(ByVal HexStr As String) As String
Dim Temp As String, I As Long
  Temp = Space$(Len(HexStr) \ 2)
  For I = 1 To Len(HexStr) \ 2
    Mid$(Temp, I, 1) = Chr$(Val("&H" & Mid$(HexStr, I * 2 - 1, 2)))
  Next I
  HexToStr = Temp
End Function
                



