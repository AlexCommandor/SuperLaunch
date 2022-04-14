VERSION 5.00
Object = "{043CDC0A-B54E-4AD9-A637-BBD69EE04568}#31.0#0"; "INIPRO.OCX"
Begin VB.Form SuperLaunch 
   Caption         =   "SuperLauncher"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   Icon            =   "SuperLaunch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin IniPro.Ini Ini1 
      Left            =   0
      Top             =   0
      _ExtentX        =   3387
      _ExtentY        =   635
   End
End
Attribute VB_Name = "SuperLaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bIniError As Boolean

