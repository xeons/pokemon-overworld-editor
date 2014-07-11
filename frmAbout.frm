VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   3960
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblAbout 
      Caption         =   $"frmAbout.frx":151A
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "Coded by Xeon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblApplicationTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pokémon Overworld Sprite Editor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image imgProgramIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":162C
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape shpHeaderBackground 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & " by Xeon"
End Sub
