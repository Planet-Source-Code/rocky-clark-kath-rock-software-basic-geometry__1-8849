Attribute VB_Name = "basCharacters"
Option Explicit

'Note:  This module is for arc demonstration purposes only.
'       You may modify and distribute this code in any way
'       you like, but the code is not commented at all and
'       there is very little error checking.

Public Type RectDbl
    Left    As Double
    Top     As Double
    Right   As Double
    Bottom  As Double
End Type

Public Type CharStruct
    rcBounds    As RectDbl
    ptCoords()  As PointDbl
    iPtCnt      As Integer
    iSetCnts()  As Integer
    iSetCnt     As Integer
End Type
Private Function CharB(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double

    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 16
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(0)
            .ptCoords(3).X = X + dWidth - (dHeight / 4.4) - (dHeight / 10)
            .ptCoords(3).Y = .ptCoords(2).Y
            .ptCoords(4).X = X
            .ptCoords(4).Y = Y + (dHeight / 2.2)
            .ptCoords(5).X = .ptCoords(3).X
            .ptCoords(5).Y = .ptCoords(4).Y
            .ptCoords(6) = .ptCoords(4)
            .ptCoords(7).X = X + dWidth - ((dHeight - (dHeight / 2.2)) / 2)
            .ptCoords(7).Y = .ptCoords(6).Y
            .ptCoords(8).X = X
            .ptCoords(8).Y = Y + dHeight
            .ptCoords(9).X = .ptCoords(7).X
            .ptCoords(9).Y = .ptCoords(8).Y
            .ptCoords(10) = .ptCoords(3)
            .ptCoords(12) = .ptCoords(5)
            .ptCoords(11).X = X + dWidth - (dHeight / 10)
            .ptCoords(11).Y = (.ptCoords(10).Y + .ptCoords(12).Y) / 2
            .ptCoords(13) = .ptCoords(7)
            .ptCoords(15) = .ptCoords(9)
            .ptCoords(14).X = X + dWidth
            .ptCoords(14).Y = (.ptCoords(13).Y + .ptCoords(15).Y) / 2
            
            .iSetCnt = 7
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
            .iSetCnts(3) = 2
            .iSetCnts(4) = 2
            .iSetCnts(5) = 3
            .iSetCnts(6) = 3
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 5
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight * 0.1)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + (dHeight / 2) + (dHeight * 0.03)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + (dHeight * 0.75)
            .ptCoords(4).X = X
            .ptCoords(4).Y = Y + dHeight - (dHeight * 0.03)
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 3
        End With
    
    End If
    
    CharB = Char
    Char = NoChar

End Function
Private Function CharC(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double

    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 3
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + dWidth
            .ptCoords(0).Y = Y + (dHeight * 0.075)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + (dHeight * 0.5)
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + dHeight - (dHeight * 0.075)
            
            .iSetCnt = 1
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 3
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + dWidth
            .ptCoords(0).Y = Y + (dHeight / 2) + (dHeight * 0.03)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + (dHeight * 0.75)
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + dHeight - (dHeight * 0.03)
            
            .iSetCnt = 1
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
        End With
    
    End If
    
    CharC = Char
    Char = NoChar

End Function

Private Function CharD(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double

    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 9
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(0)
            .ptCoords(3).X = X + dWidth - (dHeight / 2)
            .ptCoords(3).Y = .ptCoords(0).Y
            .ptCoords(4) = .ptCoords(1)
            .ptCoords(5).X = X + dWidth - (dHeight / 2)
            .ptCoords(5).Y = .ptCoords(1).Y
            .ptCoords(6) = .ptCoords(3)
            .ptCoords(7).X = X + dWidth
            .ptCoords(7).Y = Y + (dHeight * 0.5)
            .ptCoords(8) = .ptCoords(5)
            
            .iSetCnt = 4
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
            .iSetCnts(3) = 3
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 5
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + dWidth
            .ptCoords(0).Y = Y + (dHeight / 2) + (dHeight * 0.03)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + (dHeight * 0.75)
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + dHeight - (dHeight * 0.03)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + (dHeight * 0.1)
            .ptCoords(4).X = .ptCoords(3).X
            .ptCoords(4).Y = Y + dHeight
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 2
        End With
    
    End If
    
    CharD = Char
    Char = NoChar

End Function

Private Function CharE(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct
    
Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + (dWidth * 0.75)
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 8
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(0)
            .ptCoords(3).X = X + (dWidth * 0.75)
            .ptCoords(3).Y = .ptCoords(2).Y
            .ptCoords(4).X = X
            .ptCoords(4).Y = Y + (dHeight / 2.2)
            .ptCoords(5).X = X + (dWidth / 2)
            .ptCoords(5).Y = .ptCoords(4).Y
            .ptCoords(6) = .ptCoords(1)
            .ptCoords(7).X = .ptCoords(3).X
            .ptCoords(7).Y = .ptCoords(6).Y
            
            .iSetCnt = 4
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
            .iSetCnts(3) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 12
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight * 0.75)
            .ptCoords(1).X = X + dWidth
            .ptCoords(1).Y = .ptCoords(0).Y
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = .ptCoords(1).X
            .ptCoords(3).Y = Y + (dHeight / 2) + (dWidth / 2)
            .ptCoords(4) = .ptCoords(3)
            .ptCoords(5).X = X + (dWidth / 2)
            .ptCoords(5).Y = Y + (dHeight / 2)
            .ptCoords(6).X = X
            .ptCoords(6).Y = .ptCoords(4).Y
            .ptCoords(7) = .ptCoords(6)
            .ptCoords(8).X = .ptCoords(0).X
            .ptCoords(8).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(9) = .ptCoords(8)
            .ptCoords(10).X = X + (dWidth / 2)
            .ptCoords(10).Y = Y + dHeight
            .ptCoords(11).X = .ptCoords(1).X - (dWidth * 0.03)
            .ptCoords(11).Y = .ptCoords(9).Y + (dWidth * 0.2)
            
            .iSetCnt = 5
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 3
            .iSetCnts(3) = 2
            .iSetCnts(4) = 3
        End With
    
    End If
    
    CharE = Char
    Char = NoChar

End Function

Private Function CharF(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double

    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + (dWidth * 0.75)
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(0)
            .ptCoords(3).X = X + (dWidth * 0.75)
            .ptCoords(3).Y = .ptCoords(2).Y
            .ptCoords(4).X = X
            .ptCoords(4).Y = Y + (dHeight / 2.2)
            .ptCoords(5).X = X + (dWidth / 2)
            .ptCoords(5).Y = .ptCoords(4).Y
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5)
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 7
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + dWidth
            .ptCoords(0).Y = Y + (dHeight * 0.1) + ((dWidth * 0.75) / 2)
            .ptCoords(1).X = X + dWidth - ((dWidth * 0.75) / 2)
            .ptCoords(1).Y = Y + (dHeight * 0.1)
            .ptCoords(2).X = X + dWidth - (dWidth * 0.75)
            .ptCoords(2).Y = .ptCoords(0).Y
            .ptCoords(3) = .ptCoords(2)
            .ptCoords(4).X = .ptCoords(3).X
            .ptCoords(4).Y = Y + dHeight
            .ptCoords(5).X = X
            .ptCoords(5).Y = Y + (dHeight / 2)
            .ptCoords(6).X = .ptCoords(1).X
            .ptCoords(6).Y = .ptCoords(5).Y
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    End If
    
    CharF = Char
    Char = NoChar

End Function

Private Function CharG(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double

    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 7
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + dWidth
            .ptCoords(0).Y = Y + (dHeight * 0.075)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + (dHeight * 0.5)
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + dHeight - (dHeight * 0.075)
            .ptCoords(3) = .ptCoords(2)
            .ptCoords(4).X = .ptCoords(3).X
            .ptCoords(4).Y = Y + (dHeight / 2)
            .ptCoords(5) = .ptCoords(4)
            .ptCoords(6).X = X + (dWidth / 2)
            .ptCoords(6).Y = .ptCoords(5).Y
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 8
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + dWidth
            .ptCoords(0).Y = Y + (dHeight / 2) + (dHeight * 0.03)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + (dHeight * 0.75)
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + dHeight - (dHeight * 0.03)
            .ptCoords(3).X = .ptCoords(0).X
            .ptCoords(3).Y = Y + (dHeight / 2)
            .ptCoords(4).X = .ptCoords(0).X
            .ptCoords(4).Y = Y + (dHeight * (1 / 0.75)) - (dWidth / 2)
            .ptCoords(5) = .ptCoords(4)
            .ptCoords(6).X = X + (dWidth / 2)
            .ptCoords(6).Y = Y + (dHeight * (1 / 0.75))
            .ptCoords(7).X = .ptCoords(1).X
            .ptCoords(7).Y = .ptCoords(5).Y
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 2
            .iSetCnts(2) = 3
        End With
    
    End If
    
    CharG = Char
    Char = NoChar

End Function

Private Function CharH(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.5625
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y
            .ptCoords(3).X = .ptCoords(2).X
            .ptCoords(3).Y = Y + dHeight
            .ptCoords(4).X = .ptCoords(0).X
            .ptCoords(4).Y = Y + (dHeight / 2)
            .ptCoords(5).X = .ptCoords(2).X
            .ptCoords(5).Y = .ptCoords(4).Y
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 7
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight * 0.1)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + (dHeight / 2) + (dWidth / 2)
            .ptCoords(3).X = X + (dWidth / 2)
            .ptCoords(3).Y = Y + (dHeight / 2)
            .ptCoords(4).X = X + dWidth
            .ptCoords(4).Y = .ptCoords(2).Y
            .ptCoords(5) = .ptCoords(4)
            .ptCoords(6).X = .ptCoords(5).X
            .ptCoords(6).Y = .ptCoords(1).Y
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 3
            .iSetCnts(2) = 2
        End With
    
    End If

    CharH = Char
    Char = NoChar

End Function

Private Function CharI(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.5
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X + dWidth
            .ptCoords(1).Y = Y
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + dHeight
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + dHeight
            .ptCoords(4).X = X + (dWidth / 2)
            .ptCoords(4).Y = Y
            .ptCoords(5).X = X + (dWidth / 2)
            .ptCoords(5).Y = Y + dHeight
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.2)
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 8
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + (dWidth / 2) + (dWidth / 10)
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = .ptCoords(0).X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + dHeight
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + dHeight
            .ptCoords(4).X = X + (dWidth / 2)
            .ptCoords(4).Y = Y + (dHeight / 2) + (dWidth / 10)
            .ptCoords(5).X = .ptCoords(0).X
            .ptCoords(5).Y = .ptCoords(4).Y
            .ptCoords(6).X = .ptCoords(4).X
            .ptCoords(6).Y = Y + (dHeight / 2.5)
            .ptCoords(7).X = .ptCoords(0).X
            .ptCoords(7).Y = .ptCoords(6).Y
            
            .iSetCnt = 4
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
            .iSetCnts(3) = 2
        End With
    
    End If

    CharI = Char
    Char = NoChar

End Function

Private Function CharJ(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double

    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 7
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + (dWidth * 0.25)
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X + dWidth
            .ptCoords(1).Y = Y
            .ptCoords(2).X = (.ptCoords(0).X + .ptCoords(1).X) / 2
            .ptCoords(2).Y = Y
            .ptCoords(3).X = .ptCoords(2).X
            .ptCoords(3).Y = Y + dHeight - ((dWidth * 0.75) / 2)
            .ptCoords(4) = .ptCoords(3)
            .ptCoords(5).X = X + ((dWidth * 0.75) / 2)
            .ptCoords(5).Y = Y + dHeight
            .ptCoords(6).X = X
            .ptCoords(6).Y = .ptCoords(4).Y
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 3
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.25)
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 9
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + dWidth
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = .ptCoords(0).X
            .ptCoords(1).Y = Y + (dHeight * (1 / 0.75)) - (dWidth / 2)
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = X + (dWidth / 2)
            .ptCoords(3).Y = Y + (dHeight * (1 / 0.75))
            .ptCoords(4).X = X
            .ptCoords(4).Y = .ptCoords(2).Y
            .ptCoords(5).X = X + dWidth - (dWidth / 10)
            .ptCoords(5).Y = .ptCoords(0).Y + (dWidth / 10)
            .ptCoords(6).X = .ptCoords(0).X
            .ptCoords(6).Y = .ptCoords(5).Y
            .ptCoords(7).X = .ptCoords(5).X
            .ptCoords(7).Y = Y + (dHeight / 2.5)
            .ptCoords(8).X = .ptCoords(0).X
            .ptCoords(8).Y = .ptCoords(7).Y
            
            .iSetCnt = 4
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 3
            .iSetCnts(2) = 2
            .iSetCnts(3) = 2
        End With
    
    End If

    CharJ = Char
    Char = NoChar

End Function

Private Function CharK(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.5
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + (dHeight * 0.5)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y
            .ptCoords(4) = .ptCoords(2)
            .ptCoords(5).X = X + dWidth
            .ptCoords(5).Y = Y + dHeight
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.25)
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + (dHeight * 0.6)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + (dHeight * 0.5)
            .ptCoords(4).X = (.ptCoords(2).X + .ptCoords(3).X) / 2
            .ptCoords(4).Y = (.ptCoords(2).Y + .ptCoords(3).Y) / 2
            .ptCoords(5).X = X + dWidth
            .ptCoords(5).Y = Y + dHeight
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    End If

    CharK = Char
    Char = NoChar

End Function

Private Function CharL(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.5
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 4
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + dHeight
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.25)
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + (dWidth / 2) + (dWidth / 10)
            .ptCoords(0).Y = Y + (dHeight * 0.1)
            .ptCoords(1).X = .ptCoords(0).X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + dHeight
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = .ptCoords(2).Y
            .ptCoords(4).X = X + (dWidth / 2)
            .ptCoords(4).Y = .ptCoords(0).Y
            .ptCoords(5).X = .ptCoords(0).X
            .ptCoords(5).Y = .ptCoords(4).Y
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    End If

    CharL = Char
    Char = NoChar

End Function

Private Function CharM(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 8
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y
            .ptCoords(3).X = .ptCoords(2).X
            .ptCoords(3).Y = Y + dHeight
            .ptCoords(4) = .ptCoords(0)
            .ptCoords(5).X = X + (dWidth / 2)
            .ptCoords(5).Y = Y + (dHeight / 2)
            .ptCoords(6) = .ptCoords(5)
            .ptCoords(7) = .ptCoords(2)
            
            .iSetCnt = 4
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
            .iSetCnts(3) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5)
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 12
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + (dHeight / 2) + (dWidth / 4)
            .ptCoords(3).X = X + (dWidth / 4)
            .ptCoords(3).Y = Y + (dHeight / 2)
            .ptCoords(4).X = X + (dWidth / 2)
            .ptCoords(4).Y = Y + (dHeight / 2) + (dWidth / 4)
            .ptCoords(5) = .ptCoords(4)
            .ptCoords(6).X = .ptCoords(5).X
            .ptCoords(6).Y = Y + (dHeight * 0.75)
            .ptCoords(7) = .ptCoords(4)
            .ptCoords(8).X = X + (dWidth * 0.75)
            .ptCoords(8).Y = Y + (dHeight / 2)
            .ptCoords(9).X = X + dWidth
            .ptCoords(9).Y = Y + (dHeight / 2) + (dWidth / 4)
            .ptCoords(10) = .ptCoords(9)
            .ptCoords(11).X = .ptCoords(10).X
            .ptCoords(11).Y = Y + dHeight
            
            .iSetCnt = 5
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 3
            .iSetCnts(2) = 2
            .iSetCnts(3) = 3
            .iSetCnts(4) = 2
        End With
    
    End If

    CharM = Char
    Char = NoChar

End Function

Private Function CharN(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.5625
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y
            .ptCoords(3).X = .ptCoords(2).X
            .ptCoords(3).Y = Y + dHeight
            .ptCoords(4) = .ptCoords(0)
            .ptCoords(5) = .ptCoords(3)
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 7
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + (dHeight / 2) + (dWidth / 2)
            .ptCoords(3).X = X + (dWidth / 2)
            .ptCoords(3).Y = Y + (dHeight / 2)
            .ptCoords(4).X = X + dWidth
            .ptCoords(4).Y = .ptCoords(2).Y
            .ptCoords(5) = .ptCoords(4)
            .ptCoords(6).X = .ptCoords(5).X
            .ptCoords(6).Y = .ptCoords(1).Y
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 3
            .iSetCnts(2) = 2
        End With
    
    End If

    CharN = Char
    Char = NoChar

End Function

Private Function CharO(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 10
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dWidth / 2)
            .ptCoords(1).X = X + (dWidth / 2)
            .ptCoords(1).Y = Y
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + (dWidth / 2)
            .ptCoords(3).X = X
            .ptCoords(3).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(4).X = X + (dWidth / 2)
            .ptCoords(4).Y = Y + dHeight
            .ptCoords(5).X = X + dWidth
            .ptCoords(5).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(6) = .ptCoords(0)
            .ptCoords(7) = .ptCoords(3)
            .ptCoords(8) = .ptCoords(2)
            .ptCoords(9) = .ptCoords(5)
            
            .iSetCnt = 4
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 3
            .iSetCnts(2) = 2
            .iSetCnts(3) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 10
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2) + (dWidth / 2)
            .ptCoords(1).X = X + (dWidth / 2)
            .ptCoords(1).Y = Y + (dHeight / 2)
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + (dHeight / 2) + (dWidth / 2)
            .ptCoords(3).X = X
            .ptCoords(3).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(4).X = X + (dWidth / 2)
            .ptCoords(4).Y = Y + dHeight
            .ptCoords(5).X = X + dWidth
            .ptCoords(5).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(6) = .ptCoords(0)
            .ptCoords(7) = .ptCoords(3)
            .ptCoords(8) = .ptCoords(2)
            .ptCoords(9) = .ptCoords(5)
            
            .iSetCnt = 4
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 3
            .iSetCnts(2) = 2
            .iSetCnts(3) = 2
        End With
    
    End If

    CharO = Char
    Char = NoChar

End Function

Private Function CharP(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double

    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.6
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 9
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(0)
            .ptCoords(3).X = X + dWidth - (dHeight / 4)
            .ptCoords(3).Y = .ptCoords(2).Y
            .ptCoords(4) = .ptCoords(3)
            .ptCoords(5).X = X + dWidth
            .ptCoords(5).Y = Y + (dHeight / 4)
            .ptCoords(6).X = .ptCoords(4).X
            .ptCoords(6).Y = Y + (dHeight / 2)
            .ptCoords(7) = .ptCoords(6)
            .ptCoords(8).X = X
            .ptCoords(8).Y = .ptCoords(7).Y
            
            .iSetCnt = 4
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 3
            .iSetCnts(3) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 5
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + (dHeight * (1 / 0.75))
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + (dHeight / 2) + (dHeight * 0.03)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + (dHeight * 0.75)
            .ptCoords(4).X = X
            .ptCoords(4).Y = Y + dHeight - (dHeight * 0.03)
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 3
        End With
    
    End If

    CharP = Char
    Char = NoChar

End Function

Private Function CharQ(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 12
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dWidth / 2)
            .ptCoords(1).X = X + (dWidth / 2)
            .ptCoords(1).Y = Y
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + (dWidth / 2)
            .ptCoords(3).X = X
            .ptCoords(3).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(4).X = X + (dWidth / 2)
            .ptCoords(4).Y = Y + dHeight
            .ptCoords(5).X = X + dWidth
            .ptCoords(5).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(6) = .ptCoords(0)
            .ptCoords(7) = .ptCoords(3)
            .ptCoords(8) = .ptCoords(2)
            .ptCoords(9) = .ptCoords(5)
            .ptCoords(10).X = X + (dWidth * 0.75)
            .ptCoords(10).Y = Y + dHeight - (dWidth * 0.25)
            .ptCoords(11).X = X + dWidth
            .ptCoords(11).Y = Y + dHeight
            
            .iSetCnt = 5
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 3
            .iSetCnts(2) = 2
            .iSetCnts(3) = 2
            .iSetCnts(4) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 5
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + dWidth
            .ptCoords(0).Y = Y + (dHeight / 2) + (dHeight * 0.03)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + (dHeight * 0.75)
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + dHeight - (dHeight * 0.03)
            .ptCoords(3).X = .ptCoords(0).X
            .ptCoords(3).Y = Y + (dHeight / 2)
            .ptCoords(4).X = .ptCoords(0).X
            .ptCoords(4).Y = Y + (dHeight * (1 / 0.75))
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 2
        End With
    
    End If

    CharQ = Char
    Char = NoChar

End Function

Private Function CharR(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double

    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.6
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 11
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(0)
            .ptCoords(3).X = X + dWidth - (dHeight / 4)
            .ptCoords(3).Y = .ptCoords(2).Y
            .ptCoords(4) = .ptCoords(3)
            .ptCoords(5).X = X + dWidth
            .ptCoords(5).Y = Y + (dHeight / 4)
            .ptCoords(6).X = .ptCoords(4).X
            .ptCoords(6).Y = Y + (dHeight / 2)
            .ptCoords(7) = .ptCoords(6)
            .ptCoords(8).X = X
            .ptCoords(8).Y = .ptCoords(7).Y
            .ptCoords(9) = .ptCoords(7)
            .ptCoords(10).X = X + dWidth
            .ptCoords(10).Y = Y + dHeight
            
            .iSetCnt = 5
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 3
            .iSetCnts(3) = 2
            .iSetCnts(4) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 5
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + (dHeight / 2) + (dWidth / 2)
            .ptCoords(3).X = X + (dWidth / 2)
            .ptCoords(3).Y = Y + (dHeight / 2)
            .ptCoords(4).X = X + dWidth
            .ptCoords(4).Y = .ptCoords(2).Y
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 3
        End With
    
    End If

    CharR = Char
    Char = NoChar

End Function

Private Function CharS(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double

    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.6
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 13
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + (dHeight / 2)
            .ptCoords(0).Y = Y + (dHeight / 4)
            .ptCoords(1).X = X
            .ptCoords(1).Y = .ptCoords(0).Y
            .ptCoords(2).X = X + (dHeight / 4)
            .ptCoords(2).Y = Y + (dHeight / 2)
            .ptCoords(3) = .ptCoords(2)
            .ptCoords(4).X = X + dWidth - (dHeight / 4)
            .ptCoords(4).Y = Y + (dHeight / 2)
            .ptCoords(5) = .ptCoords(4)
            .ptCoords(6).X = X + dWidth
            .ptCoords(6).Y = Y + (dHeight * 0.75)
            .ptCoords(7).X = .ptCoords(4).X
            .ptCoords(7).Y = Y + dHeight
            .ptCoords(8) = .ptCoords(7)
            .ptCoords(9).X = .ptCoords(3).X
            .ptCoords(9).Y = .ptCoords(7).Y
            .ptCoords(10) = .ptCoords(9)
            .ptCoords(11).X = (X + (X + .ptCoords(10).X) / 2) / 2
            .ptCoords(11).Y = Y + dHeight - (dHeight / 16)
            .ptCoords(12).X = X
            .ptCoords(12).Y = Y + (dHeight * 0.75)
            
            .iSetCnt = 5
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 2
            .iSetCnts(2) = 3
            .iSetCnts(3) = 2
            .iSetCnts(4) = 3
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 13
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + (dHeight / 4)
            .ptCoords(0).Y = Y + (dHeight / 2) + (dHeight / 8)
            .ptCoords(1).X = X
            .ptCoords(1).Y = .ptCoords(0).Y
            .ptCoords(2).X = X + (dHeight / 8)
            .ptCoords(2).Y = Y + (dHeight * 0.75)
            .ptCoords(3) = .ptCoords(2)
            .ptCoords(4).X = X + dWidth - (dHeight / 8)
            .ptCoords(4).Y = Y + (dHeight * 0.75)
            .ptCoords(5) = .ptCoords(4)
            .ptCoords(6).X = X + dWidth
            .ptCoords(6).Y = Y + (dHeight * 0.875)
            .ptCoords(7).X = .ptCoords(4).X
            .ptCoords(7).Y = Y + dHeight
            .ptCoords(8) = .ptCoords(7)
            .ptCoords(9).X = .ptCoords(3).X
            .ptCoords(9).Y = .ptCoords(7).Y
            .ptCoords(10) = .ptCoords(9)
            .ptCoords(11).X = (X + (X + .ptCoords(10).X) / 2) / 2
            .ptCoords(11).Y = Y + dHeight - (dHeight / 32)
            .ptCoords(12).X = X
            .ptCoords(12).Y = Y + (dHeight * 0.875)
            
            .iSetCnt = 5
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 2
            .iSetCnts(2) = 3
            .iSetCnts(3) = 2
            .iSetCnts(4) = 3
        End With
    
    End If

    CharS = Char
    Char = NoChar

End Function

Private Function CharT(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.6
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 4
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X + dWidth
            .ptCoords(1).Y = Y
            .ptCoords(2).X = X + (dWidth / 2)
            .ptCoords(2).Y = Y
            .ptCoords(3).X = X + (dWidth / 2)
            .ptCoords(3).Y = Y + dHeight
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.4)
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 4
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + (dWidth / 2)
            .ptCoords(0).Y = Y + (dHeight * 0.1)
            .ptCoords(1).X = .ptCoords(0).X
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + (dHeight / 3)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + (dHeight / 3)
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
        End With
    
    End If

    CharT = Char
    Char = NoChar

End Function

Private Function CharU(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.6
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 7
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = X + (dWidth / 2)
            .ptCoords(3).Y = Y + dHeight
            .ptCoords(4).X = X + dWidth
            .ptCoords(4).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(5) = .ptCoords(4)
            .ptCoords(6).X = X + dWidth
            .ptCoords(6).Y = Y
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 3
            .iSetCnts(2) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 7
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = X + (dWidth / 2)
            .ptCoords(3).Y = Y + dHeight
            .ptCoords(4).X = X + dWidth
            .ptCoords(4).Y = Y + dHeight - (dWidth / 2)
            .ptCoords(5).X = X + dWidth
            .ptCoords(5).Y = Y + (dHeight / 2)
            .ptCoords(6).X = X + dWidth
            .ptCoords(6).Y = Y + dHeight
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 3
            .iSetCnts(2) = 2
        End With
    
    End If

    CharU = Char
    Char = NoChar

End Function

Private Function CharV(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.6
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 4
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X + (dWidth / 2)
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 4
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = X + (dWidth / 2)
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + (dHeight / 2)
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
        End With
    
    End If

    CharV = Char
    Char = NoChar

End Function

Private Function CharW(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 8
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X + (dWidth / 4)
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = X + (dWidth / 2)
            .ptCoords(3).Y = Y + (dHeight / 2)
            .ptCoords(4) = .ptCoords(3)
            .ptCoords(5).X = X + (dWidth * 0.75)
            .ptCoords(5).Y = Y + dHeight
            .ptCoords(6) = .ptCoords(5)
            .ptCoords(7).X = X + dWidth
            .ptCoords(7).Y = Y
            
            .iSetCnt = 4
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
            .iSetCnts(3) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5)
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 8
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = X + (dWidth / 4)
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = X + (dWidth / 2)
            .ptCoords(3).Y = Y + (dHeight / 2)
            .ptCoords(4) = .ptCoords(3)
            .ptCoords(5).X = X + (dWidth * 0.75)
            .ptCoords(5).Y = Y + dHeight
            .ptCoords(6) = .ptCoords(5)
            .ptCoords(7).X = X + dWidth
            .ptCoords(7).Y = Y + (dHeight / 2)
            
            .iSetCnt = 4
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
            .iSetCnts(3) = 2
        End With
    
    End If

    CharW = Char
    Char = NoChar

End Function

Private Function CharX(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.6
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 4
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X + dWidth
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + dHeight
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 4
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = X + dWidth
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + dHeight
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + (dHeight / 2)
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
        End With
    
    End If

    CharX = Char
    Char = NoChar

End Function

Private Function CharY(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.6
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X + (dWidth / 2)
            .ptCoords(1).Y = Y + (dHeight / 2)
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y
            .ptCoords(4) = .ptCoords(1)
            .ptCoords(5).X = X + (dWidth / 2)
            .ptCoords(5).Y = Y + dHeight
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 4
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = X + (dWidth / 2)
            .ptCoords(1).Y = Y + dHeight
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + (dHeight / 2)
            .ptCoords(3).X = .ptCoords(2).X
            .ptCoords(3).Y = Y + (dHeight * (1 / 0.75))
            .ptCoords(3) = PointOnLine(.ptCoords(2), .ptCoords(1), Distance(.ptCoords(2), .ptCoords(3)))
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
        End With
    
    End If

    CharY = Char
    Char = NoChar

End Function

Private Function CharZ(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double
    
    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.6
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y
            .ptCoords(1).X = X + dWidth
            .ptCoords(1).Y = Y
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + dHeight
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + dHeight
            .ptCoords(4) = .ptCoords(2)
            .ptCoords(5) = .ptCoords(1)
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + (dHeight / 2)
            .ptCoords(1).X = X + dWidth
            .ptCoords(1).Y = Y + (dHeight / 2)
            .ptCoords(2).X = X
            .ptCoords(2).Y = Y + dHeight
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + dHeight
            .ptCoords(4) = .ptCoords(2)
            .ptCoords(5) = .ptCoords(1)
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    End If

    CharZ = Char
    Char = NoChar

End Function

Private Function SpaceChar(ByVal dHeight As Double, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim dWidth  As Double
    
    With Char
        dWidth = dHeight * 0.28125
        With .rcBounds
            .Left = X
            .Top = Y
            .Right = X + dWidth
            .Bottom = Y + dHeight
        End With
        .iPtCnt = 0
    End With
    
    SpaceChar = Char
    
End Function

Private Function DotChar(ByVal dHeight As Double, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim dWidth  As Double
    
    With Char
        dWidth = dHeight * 0.03
        With .rcBounds
            .Left = X
            .Top = Y
            .Right = X + dWidth
            .Bottom = Y + dHeight
        End With
        .iPtCnt = 2
        ReDim .ptCoords(.iPtCnt - 1)
        .ptCoords(0).X = X
        .ptCoords(0).Y = Y + (dHeight * 0.75)
        .ptCoords(1).X = X + dWidth
        .ptCoords(1).Y = Y + (dHeight * 0.75)
        
        .iSetCnt = 1
        ReDim .iSetCnts(.iSetCnt - 1)
        .iSetCnts(0) = 2
    End With
    
    DotChar = Char
    
End Function

Public Sub DrawChars(picObj As Object, ByVal sText As String, ByVal iPixelHeight As Integer, Optional ByVal lColor As OLE_COLOR = vbButtonText, Optional ByVal iPixelLeft As Integer = 0, Optional ByVal iPixelTop As Integer = 0, Optional ByVal bAutoWrap As Boolean = False, Optional ByVal bShowPoints As Boolean = False, Optional ByVal lPtColor As OLE_COLOR = vbRed, Optional ByVal bShowRect As Boolean = False)

Dim iWidth  As Integer
Dim iSetIdx As Integer
Dim iPtIdx  As Integer
Dim iCnt    As Integer
Dim iCRCnt  As Integer
Dim iCRs()  As Integer
Dim lIdx    As Long
Dim lChrIdx As Long
Dim dHeight As Double
Dim dRadius As Double
Dim dOffset As Double
Dim dScale  As Double
Dim X       As Double
Dim Y       As Double
Dim sChar   As String
Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim Arc1    As ArcStruct

    If Len(sText) > 0 Then
        
        'Convert pixel coordinates to picObj.ScaleMode.
        dScale = picObj.ScaleX(1, vbPixels, picObj.ScaleMode)
        dRadius = ((picObj.DrawWidth + IIf(picObj.DrawWidth Mod 2 <> 0, 1, 0)) / 2) * dScale
        dHeight = iPixelHeight * dScale
        X = iPixelLeft * dScale
        Y = iPixelTop * dScale
        
        'Test for AutoWrap.
        If bAutoWrap Then
            
            For lIdx = 1 To Len(sText)
                'Test for CRLF
                While Mid$(sText, lIdx, 1) = Chr$(13)
                    If Len(sText) > lIdx Then
                        If Mid$(sText, lIdx + 1, 1) = Chr$(10) Then
                            X = iPixelLeft * dScale
                            Y = Y + dHeight * 1.1
                            lIdx = lIdx + 2
                        Else
                            X = iPixelLeft * dScale
                            lIdx = lIdx + 1
                        End If
                    Else
                        Exit For
                    End If
                Wend
                
                'Create the character's points and bounds.
NextChar:       Char = GetChar(Mid$(sText, lIdx, 1), dHeight, X, Y)
                
                'Offset coordinates by the DrawWidth to
                'keep the entire character within its bounds.
                If picObj.DrawWidth > 1 Then
                    dOffset = (picObj.DrawWidth / 2) * dScale
                    For iPtIdx = 0 To Char.iPtCnt - 1
                        Char.ptCoords(iPtIdx).X = Char.ptCoords(iPtIdx).X + dOffset
                        Char.ptCoords(iPtIdx).Y = Char.ptCoords(iPtIdx).Y + dOffset
                    Next
                    'Increase bounds By DrawWidth.
                    Char.rcBounds.Right = Char.rcBounds.Right + (dOffset * 2)
                    Char.rcBounds.Bottom = Char.rcBounds.Bottom + (dOffset * 2)
                End If
                
                'Test for edge of picObj.
                If Char.rcBounds.Right > picObj.ScaleWidth Then
                    For lChrIdx = lIdx To 1 Step -1
                        If iCRCnt > 0 Then
                            If lChrIdx = iCRs(iCRCnt - 1) Then
                                Exit For
                            End If
                        End If
                        sChar = Mid$(sText, lChrIdx, 1)
                        If sChar = Chr$(13) Then
                            Exit For
                        End If
                        If (sChar < "A" Or sChar > "Z") And (sChar < "a" Or sChar > "z") And sChar <> "." Then
                            lIdx = lChrIdx + 1
                            ReDim Preserve iCRs(iCRCnt)
                            iCRs(iCRCnt) = lChrIdx
                            iCRCnt = iCRCnt + 1
                            X = iPixelLeft * dScale
                            Y = Y + dHeight * 1.1
                            GoTo NextChar
                        End If
                    Next
                    ReDim Preserve iCRs(iCRCnt)
                    iCRs(iCRCnt) = lIdx - 1
                    iCRCnt = iCRCnt + 1
                    X = iPixelLeft * dScale
                    Y = Y + dHeight * 1.1
                    GoTo NextChar
                End If
                
                X = Char.rcBounds.Right + (dHeight / 10)
            
            Next
        End If
        
        For lIdx = 0 To iCRCnt - 1
            sText = Left$(sText, iCRs(lIdx) + (lIdx * 2)) & vbCrLf & Mid$(sText, iCRs(lIdx) + (lIdx * 2) + 1)
        Next
        
        X = iPixelLeft * dScale
        Y = iPixelTop * dScale
        
        'Step through each character in the string.
        For lIdx = 1 To Len(sText)
        
            'Test for CRLF
            While Mid$(sText, lIdx, 1) = Chr$(13)
                If Len(sText) > lIdx Then
                    If Mid$(sText, lIdx + 1, 1) = Chr$(10) Then
                        X = iPixelLeft * dScale
                        Y = Y + dHeight * 1.1
                        lIdx = lIdx + 2
                    Else
                        X = iPixelLeft * dScale
                        lIdx = lIdx + 1
                    End If
                Else
                    Exit For
                End If
            Wend
            
            'Create the character's points and bounds.
            Char = GetChar(Mid$(sText, lIdx, 1), dHeight, X, Y)
            
            'Offset coordinates by the DrawWidth to
            'keep the entire character within its bounds.
            If picObj.DrawWidth > 1 Then
                dOffset = (picObj.DrawWidth / 2) * dScale
                For iPtIdx = 0 To Char.iPtCnt - 1
                    Char.ptCoords(iPtIdx).X = Char.ptCoords(iPtIdx).X + dOffset
                    Char.ptCoords(iPtIdx).Y = Char.ptCoords(iPtIdx).Y + dOffset
                Next
                'Increase bounds By DrawWidth.
                Char.rcBounds.Right = Char.rcBounds.Right + (dOffset * 2)
                Char.rcBounds.Bottom = Char.rcBounds.Bottom + (dOffset * 2)
            End If
            
            'Draw the character.
            iPtIdx = 0
            For iSetIdx = 0 To Char.iSetCnt - 1
                'Get the point count for this arc or line
                iCnt = Char.iSetCnts(iSetIdx)
                
                If iCnt = 3 Then    '3 points for an arc.
                    'Calculate the arc.
                    Arc1 = CalcArc(Char.ptCoords(iPtIdx), Char.ptCoords(iPtIdx + 1), Char.ptCoords(iPtIdx + 2))
                    If Arc1.bValidArc Then  'Arc is valid
                        With Arc1
                            picObj.Circle (.ptCenter.X, .ptCenter.Y), .dRadius, lColor, .dRadsStart, .dRadsEnd
                        End With
                    Else    'Arc is too close to straight line, so use lines.
                        picObj.Line (Char.ptCoords(iPtIdx).X, Char.ptCoords(iPtIdx).Y)-(Char.ptCoords(iPtIdx + 1).X, Char.ptCoords(iPtIdx + 1).Y), lColor
                        picObj.Line (Char.ptCoords(iPtIdx + 1).X, Char.ptCoords(iPtIdx + 1).Y)-(Char.ptCoords(iPtIdx + 2).X, Char.ptCoords(iPtIdx + 2).Y), lColor
                    End If
                
                ElseIf iCnt = 2 Then    '2 points for a line.
                    picObj.Line (Char.ptCoords(iPtIdx).X, Char.ptCoords(iPtIdx).Y)-(Char.ptCoords(iPtIdx + 1).X, Char.ptCoords(iPtIdx + 1).Y), lColor
                End If
                
                iPtIdx = iPtIdx + iCnt
            
            Next
            
            If bShowPoints Then 'Show the character's points.
                iWidth = picObj.DrawWidth
                picObj.DrawWidth = 1
                For iPtIdx = 0 To Char.iPtCnt - 1
                    picObj.Circle (Char.ptCoords(iPtIdx).X, Char.ptCoords(iPtIdx).Y), dRadius, lPtColor
                Next
                picObj.DrawWidth = iWidth
            End If
            
            If bShowRect Then   'Show the character's bounding rect.
                iWidth = picObj.DrawWidth
                picObj.DrawWidth = 1
                picObj.Line (Char.rcBounds.Left, Char.rcBounds.Top)-(Char.rcBounds.Right, Char.rcBounds.Bottom), &H0&, B
                picObj.DrawWidth = iWidth
            End If
            
            'Move to end of character + 1/10th of the height for spacing.
            X = Char.rcBounds.Right + (dHeight / 10)
            
        Next
        
    End If
    
    Char = NoChar
    
End Sub

Private Function GetChar(ByVal sChar As String, ByVal dHeight As Double, Optional ByVal X As Double = 0, Optional ByVal Y As Double = 0) As CharStruct

Dim bUpper As Boolean

    bUpper = (UCase$(sChar) = sChar)
    Select Case UCase$(sChar)
        Case "A"
            GetChar = CharA(dHeight, bUpper, X, Y)
        Case "B"
            GetChar = CharB(dHeight, bUpper, X, Y)
        Case "C"
            GetChar = CharC(dHeight, bUpper, X, Y)
        Case "D"
            GetChar = CharD(dHeight, bUpper, X, Y)
        Case "E"
            GetChar = CharE(dHeight, bUpper, X, Y)
        Case "F"
            GetChar = CharF(dHeight, bUpper, X, Y)
        Case "G"
            GetChar = CharG(dHeight, bUpper, X, Y)
        Case "H"
            GetChar = CharH(dHeight, bUpper, X, Y)
        Case "I"
            GetChar = CharI(dHeight, bUpper, X, Y)
        Case "J"
            GetChar = CharJ(dHeight, bUpper, X, Y)
        Case "K"
            GetChar = CharK(dHeight, bUpper, X, Y)
        Case "L"
            GetChar = CharL(dHeight, bUpper, X, Y)
        Case "M"
            GetChar = CharM(dHeight, bUpper, X, Y)
        Case "N"
            GetChar = CharN(dHeight, bUpper, X, Y)
        Case "O"
            GetChar = CharO(dHeight, bUpper, X, Y)
        Case "P"
            GetChar = CharP(dHeight, bUpper, X, Y)
        Case "Q"
            GetChar = CharQ(dHeight, bUpper, X, Y)
        Case "R"
            GetChar = CharR(dHeight, bUpper, X, Y)
        Case "S"
            GetChar = CharS(dHeight, bUpper, X, Y)
        Case "T"
            GetChar = CharT(dHeight, bUpper, X, Y)
        Case "U"
            GetChar = CharU(dHeight, bUpper, X, Y)
        Case "V"
            GetChar = CharV(dHeight, bUpper, X, Y)
        Case "W"
            GetChar = CharW(dHeight, bUpper, X, Y)
        Case "X"
            GetChar = CharX(dHeight, bUpper, X, Y)
        Case "Y"
            GetChar = CharY(dHeight, bUpper, X, Y)
        Case "Z"
            GetChar = CharZ(dHeight, bUpper, X, Y)
        Case "."
            GetChar = DotChar(dHeight, X, Y)
        Case Else
            GetChar = SpaceChar(dHeight, X, Y)
    End Select
    
End Function

Private Function CharA(ByVal dHeight As Double, ByVal bUpper As Boolean, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As CharStruct

Dim Char    As CharStruct
Dim NoChar  As CharStruct
Dim dWidth  As Double

    dHeight = dHeight * 0.75
    If bUpper Then
        With Char
            dWidth = dHeight * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 6
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X
            .ptCoords(0).Y = Y + dHeight
            .ptCoords(1).X = X + (dWidth / 2)
            .ptCoords(1).Y = Y
            .ptCoords(2) = .ptCoords(1)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + dHeight
            .ptCoords(4).X = (X + .ptCoords(1).X) / 2
            .ptCoords(4).Y = (Y + .ptCoords(0).Y) / 2
            .ptCoords(5).X = (.ptCoords(1).X + .ptCoords(3).X) / 2
            .ptCoords(5).Y = (Y + .ptCoords(0).Y) / 2
            
            .iSetCnt = 3
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 2
            .iSetCnts(1) = 2
            .iSetCnts(2) = 2
        End With
    
    Else
        With Char
            dWidth = (dHeight * 0.5) * 0.75
            With .rcBounds
                .Left = X
                .Top = Y
                .Right = .Left + dWidth
                .Bottom = .Top + (dHeight * (1 / 0.75))
            End With
            
            .iPtCnt = 5
            ReDim .ptCoords(.iPtCnt - 1)
            .ptCoords(0).X = X + dWidth
            .ptCoords(0).Y = Y + (dHeight / 2) + (dHeight * 0.03)
            .ptCoords(1).X = X
            .ptCoords(1).Y = Y + (dHeight * 0.75)
            .ptCoords(2).X = X + dWidth
            .ptCoords(2).Y = Y + dHeight - (dHeight * 0.03)
            .ptCoords(3).X = X + dWidth
            .ptCoords(3).Y = Y + (dHeight / 2)
            .ptCoords(4).X = .ptCoords(3).X
            .ptCoords(4).Y = Y + dHeight
            
            .iSetCnt = 2
            ReDim .iSetCnts(.iSetCnt - 1)
            .iSetCnts(0) = 3
            .iSetCnts(1) = 2
        End With
    
    End If
    
    CharA = Char
    Char = NoChar
    
End Function

