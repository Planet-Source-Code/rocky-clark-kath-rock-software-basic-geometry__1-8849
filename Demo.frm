VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   Caption         =   "Drawing Arcs - (Right-click to Clear)"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbshowing   As Boolean
Private miPtCnt     As Integer
Private mPts(2)     As PointDbl
Private mptTemp     As PointDbl
Private mArc        As ArcStruct

Private Sub Form_Activate()

Dim sText As String

    Me.DrawWidth = 3
    
    sText = "This fully scalable text was created using nothing but lines " _
        & " and arcs. You only need two points for a line and three for " _
        & "an arc. The red dots show the locations of the points that " _
        & "created this text. Click three points on this form to see " _
        & "the arc routines in action."
    Call DrawChars(Me, sText, 50, vbBlue, 10, 10, True, True, vbRed, False)

    Me.DrawWidth = 1

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Note: Notice that I'm testing for a circumference of less than 25600
'       pixels (gdPi * (mArc.dRadius * 2) < 25600) before drawing the
'       arc. For some reason (unknown to me) VB has a problem drawing
'       an arc with a circumference greater than 25,600 pixels.

Dim fScale  As Single
Dim fCirc   As Single

    fScale = Me.ScaleX(1, Me.ScaleMode, vbPixels)
    
    If Button = vbLeftButton Then
    
        Select Case miPtCnt
            Case 0
                mbshowing = False
                
                'Draw a small red circle to show the point.
                Me.Circle (X, Y), Me.ScaleX(3, vbPixels, Me.ScaleMode), &HFF&
                
                'Set the new point.
                mPts(0).X = X
                mPts(0).Y = Y
                            
                miPtCnt = miPtCnt + 1
            
            Case 1
                'Erase the temp line.
                Me.DrawMode = vbInvert  '2 inverted draws = erase.
                Me.Line (mPts(0).X, mPts(0).Y)-(X, Y), &H0&
                Me.DrawMode = vbCopyPen
                
                'Draw a small red circle to show the point.
                Me.Circle (X, Y), Me.ScaleX(3, vbPixels, Me.ScaleMode), &HFF&
                
                'Set the new point.
                mPts(1).X = X
                mPts(1).Y = Y
                mptTemp = mPts(1)
                
                'Draw a line between 1st 2 points and mouse position.
                Me.DrawMode = vbInvert  'Use invert so line can be erased.
                Me.Line (mPts(0).X, mPts(0).Y)-(mptTemp.X, mptTemp.Y), &H0&
                Me.DrawMode = vbCopyPen
                
                miPtCnt = miPtCnt + 1
                
            Case 2
                'Erase the temp line.
                Me.DrawMode = vbInvert  '2 inverted draws = erase.
                
                'Calculate the arc and draw it.
                mArc = CalcArc(mPts(0), mPts(1), mptTemp)
                fCirc = (gdPi * (mArc.dRadius * 2)) * fScale 'Circumference in pixels.
                If mArc.bValidArc And fCirc < 25600 Then
                    With mArc
                        Me.Circle (.ptCenter.X, .ptCenter.Y), .dRadius, &H0&, .dRadsStart, .dRadsEnd
                    End With
                Else
                    Me.Line (mPts(0).X, mPts(0).Y)-(mPts(1).X, mPts(1).Y), &H0&
                    Me.Line (mPts(1).X, mPts(1).Y)-(mptTemp.X, mptTemp.Y), &H0&
                End If
                Me.DrawMode = vbCopyPen
                
                'Draw a small red circle to show the point.
                Me.Circle (X, Y), Me.ScaleX(3, vbPixels, Me.ScaleMode), &HFF&
                
                'Set the new point.
                mPts(2).X = X
                mPts(2).Y = Y
                
                'Calculate the arc and draw it in green.
                mArc = CalcArc(mPts(0), mPts(1), mPts(2))
                fCirc = (gdPi * (mArc.dRadius * 2)) * fScale 'Circumference in pixels.
                If mArc.bValidArc And fCirc < 25600 Then
                    With mArc
                        Me.Circle (.ptCenter.X, .ptCenter.Y), .dRadius, &H8000&, .dRadsStart, .dRadsEnd
                    End With
                Else
                    Me.Line (mPts(0).X, mPts(0).Y)-(mPts(1).X, mPts(1).Y), &H8000&
                    Me.Line (mPts(1).X, mPts(1).Y)-(mPts(2).X, mPts(2).Y), &H8000&
                End If
                
                miPtCnt = 0
                
        End Select
    
    ElseIf Button = vbRightButton Then
        Me.Cls
        miPtCnt = 0
        
    End If
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Note: Notice that I'm testing for a circumference of less than 25600
'       pixels (gdPi * (mArc.dRadius * 2) < 25600) before drawing the
'       arc. For some reason (unknown to me) VB has a problem drawing
'       an arc with a circumference greater than 25,600 pixels.

Dim fScale  As Single
Dim fCirc   As Single

    fScale = Me.ScaleX(1, Me.ScaleMode, vbPixels)
    
    Select Case miPtCnt
        Case 1
            Me.DrawMode = vbInvert  'Use invert so line can be erased.
            
            If mbshowing Then
                'Erase the temp line.
                Me.Line (mPts(0).X, mPts(0).Y)-(mptTemp.X, mptTemp.Y), &H0&
                mbshowing = False
            End If
            
            mptTemp.X = X
            mptTemp.Y = Y
            
            'Draw a line from point to mouse
            Me.Line (mPts(0).X, mPts(0).Y)-(mptTemp.X, mptTemp.Y), &H0&
            Me.DrawMode = vbCopyPen
            
            mbshowing = True
        Case 2
            'Erase the temp line.
            Me.DrawMode = vbInvert  '2 inverted draws = erase.
            
            If mbshowing Then
                'Calculate the arc and draw it.
                mArc = CalcArc(mPts(0), mPts(1), mptTemp)
                fCirc = (gdPi * (mArc.dRadius * 2)) * fScale 'Circumference in pixels.
                If mArc.bValidArc And fCirc < 25600 Then
                    With mArc
                        Me.Circle (.ptCenter.X, .ptCenter.Y), .dRadius, &H0&, .dRadsStart, .dRadsEnd
                    End With
                Else
                    Me.Line (mPts(0).X, mPts(0).Y)-(mPts(1).X, mPts(1).Y), &H0&
                    Me.Line (mPts(1).X, mPts(1).Y)-(mptTemp.X, mptTemp.Y), &H0&
                End If
                mbshowing = False
            End If
            
            'Set the new point.
            mptTemp.X = X
            mptTemp.Y = Y
            
            'Calculate the arc and draw it.
            mArc = CalcArc(mPts(0), mPts(1), mptTemp)
            fCirc = (gdPi * (mArc.dRadius * 2)) * fScale 'Circumference in pixels.
            If mArc.bValidArc And fCirc < 25600 Then
                With mArc
                    Me.Circle (.ptCenter.X, .ptCenter.Y), .dRadius, &H0&, .dRadsStart, .dRadsEnd
                End With
            Else
                Me.Line (mPts(0).X, mPts(0).Y)-(mPts(1).X, mPts(1).Y), &H0&
                Me.Line (mPts(1).X, mPts(1).Y)-(mptTemp.X, mptTemp.Y), &H0&
            End If
            
            Me.DrawMode = vbCopyPen
            mbshowing = True
            
    End Select

End Sub


