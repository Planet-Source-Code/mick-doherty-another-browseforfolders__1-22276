VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BrowseForFolder Demo"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDisplay 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox txtReturn 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   2895
   End
   Begin VB.ComboBox cmbSpecial 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0049
      TabIndex        =   7
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "New Style"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
      Begin VB.OptionButton optBrowse 
         Caption         =   "Browse For Printer"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   2655
      End
      Begin VB.OptionButton optBrowse 
         Caption         =   "Browse For Folder"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.CheckBox chkOK 
         Caption         =   "Disable OK Button"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox chkFlags 
         Caption         =   "Include Files"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox chkFlags 
         Caption         =   "Current Selection Label"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chkFlags 
         Caption         =   "Edit Box"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   120
         X2              =   2760
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   2760
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   120
         X2              =   2760
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   2760
         Y1              =   2280
         Y2              =   2280
      End
   End
   Begin VB.TextBox txtInstruction 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Type Your Own Instruction Here."
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Display Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Returned Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Open At:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Instruction Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOK_Click()
    OKEnable = Not OKEnable
End Sub

Private Sub cmbSpecial_Change()
    cmbSpecial_Click
End Sub

Private Sub cmbSpecial_Click()
    Dim i As Integer

    If cmbSpecial.ListIndex > 0 Then
        chkFlags(0).Value = 0
        chkFlags(0).Enabled = False
    Else
        chkFlags(0).Enabled = True
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdBrowse_Click()
    Dim Flags As Long
    Dim i As Integer

    If optBrowse(0) Then
        Select Case cmbSpecial.ListIndex
            Case 0 To 11: SpecialFolder = cmbSpecial.ListIndex
            Case 12 To 17: SpecialFolder = cmbSpecial.ListIndex + 4
            Case 18 To 19: SpecialFolder = cmbSpecial.ListIndex + 8
            Case 20 To 22: SpecialFolder = cmbSpecial.ListIndex + 12
            Case Else: StartFolder = cmbSpecial.Text
        End Select

        If chkFlags(0).Value = 1 Then Flags = Flags + BIF_USENEWUI
        If chkFlags(1).Value = 1 Then Flags = Flags + BIF_EDITBOX
        If chkFlags(2).Value = 1 Then Flags = Flags + BIF_STATUSTEXT
        If chkFlags(3).Value = 1 Then Flags = Flags + BIF_BROWSEINCLUDEFILES
    Else
        Flags = BIF_BROWSEFORPRINTER
        SpecialFolder = 4
    End If

    txtReturn = FolderBrowse(Me.hwnd, txtInstruction, Flags)
    
    If szDisplay <> "Printers" And szDisplay <> "Add Printer" Then
        txtDisplay = szDisplay
    Else
        txtDisplay = ""
    End If

    SpecialFolder = 0
    StartFolder = ""
    szDisplay = ""
    
End Sub

Private Sub Form_Load()
    cmbSpecial.ListIndex = 0
    cmbSpecial = CurDir
    OKEnable = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
