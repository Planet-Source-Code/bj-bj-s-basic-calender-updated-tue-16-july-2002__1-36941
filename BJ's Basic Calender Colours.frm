VERSION 5.00
Begin VB.Form frmColours 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BJ's Basic Calender - Colours."
   ClientHeight    =   1740
   ClientLeft      =   3780
   ClientTop       =   6450
   ClientWidth     =   4980
   Icon            =   "BJ's Basic Calender Colours.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1740
   ScaleWidth      =   4980
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   40
      TabIndex        =   5
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton optBackColor 
         Caption         =   "Back Color"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Click to change Back Colour"
         Top             =   1080
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optTextColor 
         Caption         =   "Text Color"
         Height          =   195
         Left            =   2040
         TabIndex        =   23
         ToolTipText     =   "Click to change Text Colour"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         TabIndex        =   6
         Top             =   160
         Width           =   3015
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   420
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   22
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   40
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   21
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   15
            Left            =   2600
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   20
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   14
            Left            =   2240
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   19
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   13
            Left            =   1880
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   18
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   12
            Left            =   1520
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   17
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   11
            Left            =   1160
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   16
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   10
            Left            =   800
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   15
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   9
            Left            =   420
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   14
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   8
            Left            =   40
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   13
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   2600
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   12
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   2240
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   11
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   1880
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   10
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   1520
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   9
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   1160
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   8
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox picColorArr 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   800
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   7
            ToolTipText     =   "Click Back or Text Colour then click here"
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.Label Default 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Default Back Color"
         Height          =   255
         Left            =   840
         TabIndex        =   25
         ToolTipText     =   "Change Colour back to Default Colours"
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5040
      Top             =   600
   End
   Begin VB.CommandButton cmdCustomColor 
      Caption         =   "Custom Colors"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      ToolTipText     =   "Click for customs colours"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Click to close colours"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label CurrentBackColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Back Colour"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Change colours back to current colours"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label CurrentTextColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Text Colour"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Change colours back to current colours"
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmColours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private ReadyToClose As Boolean
Dim colors

'Description: Calls the "Choose Color Dialog" without need for an OCX

Private Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long

Private Sub RemoveMenus(frm As Form, remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(hwnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_close Then DeleteMenu hMenu, 5, MF_BYPOSITION ' Removes Seperator between Maximize and Close
    End Sub

Private Sub cmdCustomColor_Click()

    Dim cc As ChooseColor
        Dim CustColor(16) As Long
        cc.lStructSize = Len(cc)
        cc.hwndOwner = frmColours.hwnd
        cc.hInstance = App.hInstance
        cc.flags = 0
        cc.lpCustColors = String$(16 * 4, 0)
        Dim a
        Dim X
        Dim c1
        Dim c2
        Dim c3
        Dim c4
        a = ChooseColor(cc)
        Cls
        If (a) Then
 '           MsgBox "Color chosen:" & Str$(cc.rgbResult)

                For X = 1 To Len(cc.lpCustColors) Step 4
                        c1 = Asc(Mid$(cc.lpCustColors, X, 1))
                        c2 = Asc(Mid$(cc.lpCustColors, X + 1, 1))
                        c3 = Asc(Mid$(cc.lpCustColors, X + 2, 1))
                        c4 = Asc(Mid$(cc.lpCustColors, X + 3, 1))
                        CustColor(X / 4) = (c1) + (c2 * 256) + (c3 * 65536) + (c4 * 16777216)
'                        MsgBox "Custom Color " & Int(x / 4) & " = " & CustColor(x / 4)
                Next X
        Else
        
        End If
'*********************************




    If optBackColor.Value = True Then
'CommonDialog1.ShowColor
    frmBasicCalender.BackColor = Str$(cc.rgbResult)
    frmBasicCalender.Frame1.BackColor = Str$(cc.rgbResult)
    frmBasicCalender.Frame2.BackColor = Str$(cc.rgbResult)
    frmBasicCalender.Frame3.BackColor = Str$(cc.rgbResult)
    frmBasicCalender.Frame4.BackColor = Str$(cc.rgbResult)
    Else
    If optTextColor.Value = True Then
'CommonDialog1.ShowColor
    frmBasicCalender.lblMonth.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.lblYear.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.lblTime.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.lblDayofYear.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.lblWeekofYear.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.Label1.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.Label2.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.Label3.ForeColor = &H80FF&
    frmBasicCalender.Label4.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.Label6.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.Label7.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.Label8.ForeColor = Str$(cc.rgbResult)
    frmBasicCalender.lblHrs.ForeColor = &HFFFF&
    frmBasicCalender.lblMins.ForeColor = &HFFFF&
    frmBasicCalender.lblDays.ForeColor = &HFFFF&
    frmBasicCalender.lblSecs.ForeColor = &HFFFF&
    End If
    End If
    
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Basic Calender", "Calender Text Colour", Str$(cc.rgbResult)
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Basic Calender", "Calender Back Colour", Str$(cc.rgbResult)
    End Sub

Private Sub cmdExit_Click()
Unload Me
frmBasicCalender.Show

End Sub

Private Sub CurrentBackColor_Click()
    frmBasicCalender.BackColor = CurrentBackColor.BackColor
    frmBasicCalender.Frame1.BackColor = CurrentBackColor.BackColor
    frmBasicCalender.Frame2.BackColor = CurrentBackColor.BackColor
    frmBasicCalender.Frame3.BackColor = CurrentBackColor.BackColor
    frmBasicCalender.Frame4.BackColor = CurrentBackColor.BackColor
    
End Sub

Private Sub CurrentTextColor_Click()
    frmBasicCalender.lblMonth.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.lblYear.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.lblTime.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.lblDayofYear.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.lblWeekofYear.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.Label1.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.Label2.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.Label3.ForeColor = &H80FF&
    frmBasicCalender.Label4.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.Label6.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.Label7.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.Label8.ForeColor = CurrentTextColor.ForeColor
    frmBasicCalender.lblHrs.ForeColor = &HFFFF&
    frmBasicCalender.lblMins.ForeColor = &HFFFF&
    frmBasicCalender.lblDays.ForeColor = &HFFFF&
    frmBasicCalender.lblSecs.ForeColor = &HFFFF&

End Sub

Private Sub Default_Click()
    If optBackColor.Value = True Then
    frmBasicCalender.BackColor = Default.BackColor
    frmBasicCalender.Frame1.BackColor = Default.BackColor
    frmBasicCalender.Frame2.BackColor = Default.BackColor
    frmBasicCalender.Frame3.BackColor = Default.BackColor
    frmBasicCalender.Frame4.BackColor = Default.BackColor
    Else
    If optTextColor.Value = True Then
    frmBasicCalender.lblMonth.ForeColor = vbBlack
    frmBasicCalender.lblYear.ForeColor = vbBlack
    frmBasicCalender.lblTime.ForeColor = vbBlack
    frmBasicCalender.lblDayofYear.ForeColor = vbBlack
    frmBasicCalender.lblWeekofYear.ForeColor = vbBlack
    frmBasicCalender.Label1.ForeColor = vbBlack
    frmBasicCalender.Label2.ForeColor = vbBlack
    frmBasicCalender.Label3.ForeColor = &H80FF&
    frmBasicCalender.Label4.ForeColor = vbBlack
    frmBasicCalender.Label6.ForeColor = vbBlack
    frmBasicCalender.Label7.ForeColor = vbBlack
    frmBasicCalender.Label8.ForeColor = vbBlack
    frmBasicCalender.lblHrs.ForeColor = &HFFFF&
    frmBasicCalender.lblMins.ForeColor = &HFFFF&
    frmBasicCalender.lblDays.ForeColor = &HFFFF&
    frmBasicCalender.lblSecs.ForeColor = &HFFFF&
    End If
    End If
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Basic Calender", "Calender Text Colour", frmBasicCalender.lblMonth.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Basic Calender", "Calender Back Colour", frmBasicCalender.BackColor
    
End Sub

'*************************************************
' Purpose:  Unload the form.
'*************************************************
'*************************************************
' Purpose:  Initialize the form by setting the
'           colors of the picture boxes.
'*************************************************
Private Sub Form_Load()

frmColours.Left = frmBasicCalender.Left + 1200



'If App.PrevInstance = True Then
'End
'End If
    Dim intI As Integer ' counter
    RemoveMenus Me, True
    For intI = 0 To 15 '16 colors
  '  For intI = 0 To 15 '16 colors
        ' set color
        picColorArr(intI).BackColor = QBColor(intI)
    Next intI
        Default.Caption = "Default Back Colour"
        On Error Resume Next
    CurrentTextColor.ForeColor = frmBasicCalender.Label1.ForeColor
    CurrentBackColor.BackColor = frmBasicCalender.Frame2.BackColor
    CurrentTextColor.BackColor = frmBasicCalender.Frame2.BackColor
    CurrentBackColor.ForeColor = frmBasicCalender.Label1.ForeColor
End Sub


Private Sub optBackColor_Click()
    Default.Caption = "Default Back Colour"

End Sub

Private Sub optTextColor_Click()
    Default.Caption = "Default Text Colour"

End Sub

'*************************************************
' Purpose:  Sets the text color of the selection
'           on the calling form.
' Inputs:   intIndex: The index of the clicked pict.
'*************************************************
Private Sub picColorArr_Click(intIndex As Integer)
    If optBackColor.Value = True Then
    frmBasicCalender.BackColor = QBColor(intIndex)
    frmBasicCalender.Frame1.BackColor = QBColor(intIndex)
    frmBasicCalender.Frame2.BackColor = QBColor(intIndex)
    frmBasicCalender.Frame3.BackColor = QBColor(intIndex)
    frmBasicCalender.Frame4.BackColor = QBColor(intIndex)
    Else
    If optTextColor.Value = True Then
    frmBasicCalender.lblMonth.ForeColor = QBColor(intIndex)
    frmBasicCalender.lblYear.ForeColor = QBColor(intIndex)
    frmBasicCalender.lblTime.ForeColor = QBColor(intIndex)
    frmBasicCalender.lblDayofYear.ForeColor = QBColor(intIndex)
    frmBasicCalender.lblWeekofYear.ForeColor = QBColor(intIndex)
    frmBasicCalender.Label1.ForeColor = QBColor(intIndex)
    frmBasicCalender.Label2.ForeColor = QBColor(intIndex)
    frmBasicCalender.Label3.ForeColor = &H80FF&
    frmBasicCalender.Label4.ForeColor = QBColor(intIndex)
    frmBasicCalender.Label6.ForeColor = QBColor(intIndex)
    frmBasicCalender.Label7.ForeColor = QBColor(intIndex)
    frmBasicCalender.Label8.ForeColor = QBColor(intIndex)
    frmBasicCalender.lblHrs.ForeColor = &HFFFF&
    frmBasicCalender.lblMins.ForeColor = &HFFFF&
    frmBasicCalender.lblDays.ForeColor = &HFFFF&
    frmBasicCalender.lblSecs.ForeColor = &HFFFF&
    End If
    End If
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Basic Calender", "Calender Text Colour", frmBasicCalender.lblMonth.ForeColor
SetStringValue "HKEY_CURRENT_USER\Software\BJ\BJ's How to Get...\Basic Calender", "Calender Back Colour", frmBasicCalender.BackColor
End Sub

Private Sub Colours()
'Place the following code in under a command button or in a menu, etc...

    Dim cc As ChooseColor
        Dim CustColor(16) As Long
        cc.lStructSize = Len(cc)
        cc.hwndOwner = frmColours.hwnd
        cc.hInstance = App.hInstance
        cc.flags = 0
        cc.lpCustColors = String$(16 * 4, 0)
        Dim a
        Dim X
        Dim c1
        Dim c2
        Dim c3
        Dim c4
        a = ChooseColor(cc)
        Cls
        If (a) Then
            MsgBox "Color chosen:" & Str$(cc.rgbResult)

                For X = 1 To Len(cc.lpCustColors) Step 4
                        c1 = Asc(Mid$(cc.lpCustColors, X, 1))
                        c2 = Asc(Mid$(cc.lpCustColors, X + 1, 1))
                        c3 = Asc(Mid$(cc.lpCustColors, X + 2, 1))
                        c4 = Asc(Mid$(cc.lpCustColors, X + 3, 1))
                        CustColor(X / 4) = (c1) + (c2 * 256) + (c3 * 65536) + (c4 * 16777216)
                        MsgBox "Custom Color " & Int(X / 4) & " = " & CustColor(X / 4)
                Next X
        Else
                MsgBox "Cancel was pressed"
        End If

End Sub

Private Sub Timer1_Timer()
Frame1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)

End Sub
