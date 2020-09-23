VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Gradient Form"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin GradientProject.Duncan_GradientBackground Duncan_GradientBackground2 
      Left            =   3720
      Top             =   2160
      _extentx        =   1323
      _extenty        =   1323
      colourtop       =   16710908
      colourbottom    =   16572875
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "See notes within control for more details."
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Resize this form now to see how little lag is generated"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Demonstration of Background Gradient UserControl"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This has got to be the QUICKEST and EASIEST way to give your forms that professional look."
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

