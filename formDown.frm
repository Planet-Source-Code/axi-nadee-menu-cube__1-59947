VERSION 5.00
Begin VB.Form formDown 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   Picture         =   "formDown.frx":0000
   ScaleHeight     =   4615.211
   ScaleMode       =   0  'User
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblUp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lower Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "formDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
