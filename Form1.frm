VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Save Settings To A TextFile"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   2153
      Width           =   4815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3353
      TabIndex        =   13
      Top             =   3473
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save Settings To File"
      Height          =   615
      Left            =   1913
      TabIndex        =   12
      Top             =   3473
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Settings From File"
      Height          =   615
      Left            =   473
      TabIndex        =   11
      Top             =   3473
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog c1 
      Left            =   1320
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Br&owse"
      Height          =   285
      Left            =   3840
      TabIndex        =   10
      Top             =   2993
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   2993
      Width           =   3615
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1313
      Width           =   2175
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   953
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   593
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   233
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select an option"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   113
      Width           =   2295
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Any Text : "
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1793
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Save to file : "
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2633
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fso As New FileSystemObject
Private strm As TextStream
Private strName As String
Private Sub Command1_Click()
    With c1
        .Filter = "All Files|*.*"
        .InitDir = App.Path
        .ShowSave
        Text1.Text = .FileName
    End With
End Sub

Private Sub Command2_Click()
    With c1
        .Filter = "All Files|*.*"
        .InitDir = App.Path
        .ShowOpen
        GetSettings .FileName
    End With
End Sub

Private Sub GetSettings(ByVal FileName As String)
    Set strm = fso.OpenTextFile(FileName, ForReading)
    With strm
        Option1.Value = .ReadLine
        Option2.Value = .ReadLine
        Option3.Value = .ReadLine
        
        Check1.Value = .ReadLine
        Check2.Value = .ReadLine
        Check3.Value = .ReadLine
        Check4.Value = .ReadLine
        
        Text2.Text = .ReadLine
        
        Text1.Text = FileName
        
        .Close
    End With
End Sub

Private Sub SaveSettings(ByVal FileName As String)
    Set strm = fso.CreateTextFile(FileName, True)
    With strm
        .WriteLine Option1.Value
        .WriteLine Option2.Value
        .WriteLine Option3.Value
        
        .WriteLine Check1.Value
        .WriteLine Check2.Value
        .WriteLine Check3.Value
        .WriteLine Check4.Value
        
        .WriteLine Text2.Text
        
        .Close
    End With
    MsgBox "The following settings were saved to " & FileName & " : " & vbCrLf & vbCrLf & _
        "Option1.Value = " & Option1.Value & vbCrLf & _
        "Option2.Value = " & Option2.Value & vbCrLf & _
        "Option3.Value = " & Option3.Value & vbCrLf & _
        "Check1.Value = " & Check1.Value & vbCrLf & _
        "Check2.Value = " & Check2.Value & vbCrLf & _
        "Check3.Value = " & Check3.Value & vbCrLf & _
        "Check4.Value = " & Check4.Value & vbCrLf & _
        "Text = " & Text2.Text, vbOKOnly + vbInformation, "Results"
End Sub

Private Sub Command3_Click()
    SaveSettings Text1.Text
End Sub

Private Sub Command4_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Text1.Text = App.Path & "\Settings.BartNet"
End Sub
