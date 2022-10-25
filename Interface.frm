VERSION 5.00
Begin VB.Form InterfaceWindow 
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6180
   Icon            =   "Interface.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   21.25
   ScaleMode       =   4  'Character
   ScaleWidth      =   51.5
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ResultsBox 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1932
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      ToolTipText     =   "Displays information regarding the specified GUIDs after a search."
      Top             =   3000
      Width           =   5895
   End
   Begin VB.CommandButton SearchButton 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5040
      TabIndex        =   2
      Top             =   2520
      Width           =   972
   End
   Begin VB.TextBox GUIDListBox 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1932
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      ToolTipText     =   "Enter the GUID's to search for here. Each GUID should be on its own line."
      Top             =   480
      Width           =   5895
   End
   Begin VB.Label ResultsLabel 
      Caption         =   "Results:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   732
   End
   Begin VB.Label GUIDsLabel 
      Caption         =   "GUIDs:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   612
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface window.
Option Explicit

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Caption = App.Title & " v" & CStr(App.Major) & "." & CStr(App.Minor) & CStr(App.Revision) & " - by: " & App.CompanyName
   Me.Width = Screen.Width / 1.5
   Me.Height = Screen.Height / 1.5
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure adjusts this window to its new size.
Private Sub Form_Resize()
On Error Resume Next
   GUIDListBox.Width = Me.ScaleWidth - 2
   GUIDListBox.Height = (Me.ScaleHeight / 2) - 2.5
   
   ResultsBox.Width = Me.ScaleWidth - 2
   ResultsBox.Height = (Me.ScaleHeight / 2) - 2.5
   ResultsBox.Top = GUIDListBox.Top + GUIDListBox.Height + 2.5
   
   ResultsLabel.Top = GUIDListBox.Top + GUIDListBox.Height + 1
   
   SearchButton.Left = (Me.ScaleWidth - 2) - SearchButton.Width
   SearchButton.Top = GUIDListBox.Top + GUIDListBox.Height + 0.5
End Sub

'This procedure gives the command to start searching for the specified GUIDs.
Private Sub SearchButton_Click()
On Error GoTo ErrorTrap
Dim GUID As Variant

   With ResultsBox
      .Text = vbNullString
      For Each GUID In Split(GUIDListBox.Text, vbCrLf)
         .Text = .Text & FindGUID(CStr(GUID))
         DoEvents
      Next GUID
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


