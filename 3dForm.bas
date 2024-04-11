Attribute VB_Name = "ThreeDForm"
Option Explicit ' Explicitly define variables

' This Sub is called by FormInner/Outer Bevel to draw the
' lines for FormInnerBevel and FormOuterBevel
Sub FormBevelLines(FormFrame As Form, side, wid, color)

    Dim X1, Y1, X2, Y2 As Integer
    Dim rightX, bottomY
    Dim dx1, dx2, dy1, dy2 As Integer
    Dim i
       
    rightX = FormFrame.ScaleWidth - 1
    bottomY = FormFrame.ScaleHeight - 1
       
    Select Case side
        Case 0 'Left side
            X1 = 0: dx1 = 1
            X2 = 0: dx2 = 1
            Y1 = 0: dy1 = 1
            Y2 = bottomY + 1: dy2 = -1
        Case 1 'Right side
            X1 = rightX: dx1 = -1
            X2 = X1: dx2 = dx1
            Y1 = 0: dy1 = 1
            Y2 = bottomY + 1: dy2 = -1
        Case 2 'Top side
            X1 = 0: dx1 = 1
            X2 = rightX: dx2 = -1
            Y1 = 0: dy1 = 1
            Y2 = 0: dy2 = 1
        Case 3 'Bottom side
            X1 = 1: dx1 = 1
            X2 = rightX + 1: dx2 = -1
            Y1 = bottomY: dy1 = -1
            Y2 = Y1: dy2 = dy1
    End Select
    
    For i = 1 To wid

        FormFrame.Line (X1, Y1)-(X2, Y2), color
        X1 = X1 + dx1
        X2 = X2 + dx2
        Y1 = Y1 + dy1
        Y2 = Y2 + dy2
                    
    Next i

End Sub

' Here are the 2 main routines:

' This sub draws raised bevels on a Form
' Parameters TypeComments
' FormFrameFormthe Form to bevel
' BevelWidthintegerwidth of bevel in pixels
Sub FormOuterBevel(FormFrame As Form, BevelWidth As Integer)

    FormFrame.ScaleMode = 3 ' Pixels

    FormBevelLines FormFrame, 0, BevelWidth, QBColor(15) ' White

    FormBevelLines FormFrame, 1, BevelWidth, QBColor(8) ' D.Gray

    FormBevelLines FormFrame, 2, BevelWidth, QBColor(15) ' White

    FormBevelLines FormFrame, 3, BevelWidth, QBColor(8) ' D.Gray
    
End Sub


' This sub draws recessed bevels on a Form
' Parameters TypeComments
' FormFrameFormthe Form to bevel
' BevelWidthintegerwidth of bevel in pixels
Sub FormInnerBevel(FormFrame As Form, BevelWidth As Integer)
    
    FormFrame.ScaleMode = 3 ' Pixels

    FormBevelLines FormFrame, 0, BevelWidth, QBColor(8) ' D.Gray
    
    FormBevelLines FormFrame, 1, BevelWidth, QBColor(15) ' White
    
    FormBevelLines FormFrame, 2, BevelWidth, QBColor(8) ' D.Gray
    
    FormBevelLines FormFrame, 3, BevelWidth, QBColor(15) ' White
                                                                      
End Sub
