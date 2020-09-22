VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Tint Shop"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   526
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   676
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   6495
      Left            =   3960
      ScaleHeight     =   6435
      ScaleWidth      =   3675
      TabIndex        =   15
      Top             =   240
      Width           =   3735
      Begin VB.PictureBox picOut 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   5055
         Left            =   0
         ScaleHeight     =   333
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   16
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   7080
   End
   Begin VB.PictureBox Picture1 
      Height          =   6495
      Left            =   120
      MouseIcon       =   "frmMain.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   6435
      ScaleWidth      =   3675
      TabIndex        =   13
      Top             =   240
      Width           =   3735
      Begin VB.PictureBox picIn 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   5055
         Left            =   0
         ScaleHeight     =   333
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   14
         Top             =   0
         Width           =   3615
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help ?"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   6120
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Picture"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   5640
         Width           =   1575
      End
      Begin VB.OptionButton optAffect 
         Caption         =   "Affect Shadow"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   3600
         Width           =   1815
      End
      Begin VB.OptionButton optAffect 
         Caption         =   "Affect Midtones"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   3960
         Width           =   1815
      End
      Begin VB.OptionButton optAffect 
         Caption         =   "Affect Highlights"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   4320
         Width           =   1815
      End
      Begin VB.OptionButton optAffect 
         Caption         =   "Affect All"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   4680
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Picture"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.HScrollBar hsGroup 
         Height          =   255
         Left            =   120
         Max             =   25
         Min             =   1
         TabIndex        =   2
         Top             =   1320
         Value           =   10
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Select Groups To Add (ctrl-click to multiselect)"
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Affect:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Number of groups (1-25)"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This program was written after I saw a program Tint, I didn't look at the code
' but wanted to see if I could copy the effect. Not really optimized
' so make sure you compile if you use a large number of groups.

Private Type RGBType
    r As Integer
    g As Integer
    b As Integer
End Type
    
Private Type HSLType
    H As Double
    s As Double
    l As Double
End Type
    
Private Type pix
    hsl As HSLType
    rgb As RGBType
    greyRGB As RGBType
    group As Integer
    clr As Long
End Type

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long

Dim pixData() As pix

Private Sub cmdGo_Click()
    Me.MousePointer = vbHourglass
    Tint
    cmdSave.Enabled = True
    Me.MousePointer = vbDefault
End Sub

Private Function Tint()
    Dim numGroups As Integer
    Dim step As Single
    Dim X As Integer, Y As Integer, i As Integer
    Dim minStep As Single, maxStep As Single
    Dim isInGroup As Boolean, isInTone As Boolean
    
    numGroups = hsGroup.Value
    step = 1 / numGroups ' How large is each piece, hsl h is 0-1
    
    For Y = 0 To picIn.ScaleHeight
        For X = 0 To picIn.ScaleWidth
            For i = 0 To numGroups + 1
                minStep = i * step
                maxStep = (i + 1) * step
                If pixData(X, Y).hsl.H >= minStep Then
                    If pixData(X, Y).hsl.H <= maxStep Then
                        If List1.Selected(i) Then
                            ' Its in the group so paint it colored
                            isInGroup = True
                            If optAffect(3).Value Then
                                ' Affect All
                                isInTone = True
                                Exit For
                            ElseIf optAffect(0).Value Then
                                ' Shadow
                                If pixData(X, Y).hsl.l <= 0.3 Then
                                    isInTone = True
                                    Exit For
                                End If
                            ElseIf optAffect(1).Value Then
                                ' MidTone
                                If pixData(X, Y).hsl.l <= 0.6 Then
                                    If pixData(X, Y).hsl.l >= 0.3 Then
                                        isInTone = True
                                        Exit For
                                    End If
                                End If
                            ElseIf optAffect(2).Value Then
                                ' Highlight
                                If pixData(X, Y).hsl.l <= 1 Then
                                    If pixData(X, Y).hsl.l >= 0.6 Then
                                        isInTone = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
            
            If isInGroup Then
                If isInTone Then
                    ' Its in the group so paint it colored
                    SetPixel picOut.hDC, X, Y, pixData(X, Y).clr
                Else
                    ' Paint it gray scale
                    SetPixel picOut.hDC, X, Y, rgb(pixData(X, Y).greyRGB.r, pixData(X, Y).greyRGB.g, pixData(X, Y).greyRGB.b)
                End If
            Else
                ' Paint it gray scale
                SetPixel picOut.hDC, X, Y, rgb(pixData(X, Y).greyRGB.r, pixData(X, Y).greyRGB.g, pixData(X, Y).greyRGB.b)
            End If
            isInGroup = False
            isInTone = False
        Next X
    Next Y
    
    picOut.Refresh
End Function

Private Sub cmdHelp_Click()
    frmHelp.Show
End Sub

Private Sub cmdLoad_Click()
    Me.MousePointer = vbHourglass
    cmdGo.Enabled = True
    LoadPic
    Me.MousePointer = vbDefault
End Sub

Private Sub LoadPic()
    Dim X As Integer, Y As Integer, i As Integer, j As Integer
    Dim W As Integer, H As Integer
    Dim clr As Long, rgb1 As RGBType
    CD1.ShowOpen
    picIn.Picture = LoadPicture(CD1.FileName)
    W = picIn.ScaleWidth
    H = picIn.ScaleHeight
    
    picOut.Width = picIn.Width
    picOut.Height = picIn.Height
    ReDim pixData(W, H)
    
    For Y = 0 To H
        For X = 0 To W
            clr = GetPixel(picIn.hDC, X, Y)
            rgb1 = LongtoRGB(clr)
            pixData(X, Y).rgb = rgb1
            pixData(X, Y).greyRGB = GrayScale(pixData(X, Y).rgb)
            pixData(X, Y).hsl = RGBToHSL(pixData(X, Y).rgb.g, pixData(X, Y).rgb.r, pixData(X, Y).rgb.b)
            pixData(X, Y).clr = clr
        Next X
    Next Y
End Sub

Private Function GrayScale(rgb2 As RGBType) As RGBType
    Dim GrayValue As Double
    Dim retVal As RGBType
    Dim r As Double, g As Double, b As Double
    
    r = rgb2.r
    g = rgb2.g
    b = rgb2.b
    GrayValue = ((222 * r) + (707 * g) + (71 * b)) / 1000
    
    retVal.r = GrayValue
    retVal.g = GrayValue
    retVal.b = GrayValue
    
    GrayScale = retVal
End Function

Private Function HSLtoRGB(ByVal H As Double, ByVal s As Double, ByVal l As Double) As RGBType
        Dim r As Integer, g As Integer, b As Integer
        Dim var_1 As Double
        Dim var_2 As Double
        Dim retVal As RGBType

        If (s = 0) Then '                       //HSL values = 0 ÷ 1
            r = l * 255 '                      //RGB results = 0 ÷ 255
            g = l * 255
            b = l * 255
        Else
            If (l < 0.5) Then
                var_2 = l * (1 + s)
            Else
                var_2 = (l + s) - (s * l)
            End If
            var_1 = 2 * l - var_2

            r = 255 * Hue_2_RGB(var_1, var_2, H + (1 / 3))
            g = 255 * Hue_2_RGB(var_1, var_2, H)
            b = 255 * Hue_2_RGB(var_1, var_2, H - (1 / 3))
        End If

        retVal.r = r
        retVal.g = g
        retVal.b = b

        HSLtoRGB = retVal
    End Function
    
        Private Function Hue_2_RGB(ByVal v1, ByVal v2, ByVal vH) As Double
        ' Used with HSL conversion
        If (vH < 0) Then vH = vH + 1
        If (vH > 1) Then vH = vH - 1
        If ((6 * vH) < 1) Then
            Hue_2_RGB = (v1 + (v2 - v1) * 6 * vH)
            Exit Function
        End If
        If ((2 * vH) < 1) Then
            Hue_2_RGB = (v2)
            Exit Function
        End If
        
        If ((3 * vH) < 2) Then
            Hue_2_RGB = (v1 + (v2 - v1) * ((2 / 3) - vH) * 6)
            Exit Function
        End If
        Hue_2_RGB = (v1)
    End Function
    
Private Function RGBToHSL(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As HSLType
        Dim var_R As Double, var_G As Double, var_B As Double
        Dim retVal As HSLType
        Dim var_X As Double, var_Y As Double, var_Z As Double
        Dim var_Min As Double, var_Max As Double, del_Max As Double
        Dim H As Double, s As Double, l As Double
        Dim del_R As Double, del_G As Double, del_B As Double

'        If (R = 0 And G = 0 And B = 0) Or (R = 255 And G = 255 And B = 255) Then ' Undefined division below so bail now
'            retVal.H = H
'            retVal.S = S
'            retVal.L = L
'
'            RGBToHSL = retVal
'            Exit Function
'        End If
        
        var_R = (r / 255) '                     //Where RGB values = 0 ÷ 255
        var_G = (g / 255)
        var_B = (b / 255)

        var_Min = Min(Min(var_R, var_G), var_B)  '  //Min. value of RGB
        var_Max = Max(Max(var_R, var_G), var_B) '    //Max. value of RGB
        del_Max = var_Max - var_Min             '//Delta RGB value

        l = (var_Max + var_Min) / 2

        If (del_Max = 0) Then                     '//This is a gray, no chroma...
            H = 0 '                                //HSL results = 0 ÷ 1
            s = 0
        Else '                                    //Chromatic data...
            If (l < 0.5) Then
                s = del_Max / (var_Max + var_Min)
            Else
                s = del_Max / (2 - var_Max - var_Min)
            End If
        End If
        
        If del_Max = 0 Then del_Max = 0.00001 ' Undefined, just put in a tiny amount
        
        del_R = (((var_Max - var_R) / 6) + (del_Max / 2)) / del_Max
        del_G = (((var_Max - var_G) / 6) + (del_Max / 2)) / del_Max
        del_B = (((var_Max - var_B) / 6) + (del_Max / 2)) / del_Max

        If (var_R = var_Max) Then
            H = del_B - del_G
        ElseIf (var_G = var_Max) Then
            H = (1 / 3) + del_R - del_B
        ElseIf (var_B = var_Max) Then
            H = (2 / 3) + del_G - del_R
        End If

        If (H < 0) Then H = H + 1
        If (H > 1) Then H = H - 1

        retVal.H = H
        retVal.s = s
        retVal.l = l

        RGBToHSL = retVal

    End Function
    
    Private Function Min(A As Double, b As Double) As Double
        If A <= b Then
            Min = A
        Else
            Min = b
        End If
    End Function
    
    Private Function Max(A As Double, b As Double) As Double
        If A <= b Then
             Max = b
        Else
            Max = A
        End If
    End Function
Private Function LongtoRGB(clr As Long) As RGBType
        Dim rgb1 As RGBType
        
        rgb1.r = clr And 255
        rgb1.g = (clr And 65280) \ 256&
        rgb1.b = (clr And 16711680) \ 65535
        
        LongtoRGB = rgb1
    End Function

Private Sub cmdSave_Click()
    CD1.ShowSave
    SavePicture picOut.Image, CD1.FileName
End Sub

Private Sub Form_Load()
    hsGroup_Change
End Sub

Private Sub Form_Resize()
    Frame1.Left = Me.ScaleWidth - Frame1.Width
    Picture1.Height = Me.ScaleHeight - 20
    Picture2.Height = Me.ScaleHeight - 20
    Picture1.Width = (Frame1.Left - Picture1.Left - 20) \ 2
    Picture2.Left = Picture1.Left + Picture1.Width + 10
    Picture2.Width = Picture1.Width
End Sub

Private Sub hsGroup_Change()
    Dim i As Integer
    Dim numGroups As Integer
    numGroups = hsGroup.Value
    
    List1.Clear
    
    For i = 1 To numGroups
        List1.AddItem i
    Next i
End Sub

Private Sub picIn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrMove.Enabled = True
    ReleaseCapture
    SendMessage picIn.hWnd, &HA1, 2, ByVal 0&
End Sub

Private Sub picIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrMove.Enabled = False
End Sub

Private Sub tmrMove_Timer()
    picOut.Left = picIn.Left
    picOut.Top = picIn.Top
End Sub
