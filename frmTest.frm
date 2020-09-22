VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "clsOnTop Test Program"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmSubs 
      Caption         =   "Subroutines"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5655
      Begin VB.CommandButton cmdNormal 
         Caption         =   "Make"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdMakeTopMost 
         Caption         =   "Make"
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Make this form topmost."
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblLabels 
         Caption         =   "Make this form normal again."
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblLabels 
         Caption         =   "MakeNormal:"
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
         Index           =   8
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "MakeTopMost:"
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
         Index           =   9
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label lblInfo 
      Caption         =   "Welcome to the clsOnTop Testing Program by Aerodynamica Software (http://aerodynamica.port5.com)."
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declarations (needed)
Private ontop As New clsOnTop

'form load procedures
Private Sub Form_Load()
    Set ontop = New clsOnTop
    
    'make normal, just in case
    ontop.MakeNormal hWnd
End Sub

'make top most button
Private Sub cmdMakeTopMost_Click()
    ontop.MakeTopMost hWnd
End Sub

'make normal button
Private Sub cmdNormal_Click()
    ontop.MakeNormal hWnd
End Sub
