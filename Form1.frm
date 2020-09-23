VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tmx Directory Tool Beta 0.1a"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ilist 
      Left            =   8280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   53
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":39C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":42A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":545A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6612
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":77CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":80A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8982
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":925E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A416
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ACF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B5CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BEAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C786
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D5DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DEB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E792
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FF26
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":125BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":154E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":163BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1729A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AA0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B6E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C3C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D09E
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DD7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EA56
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F732
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2040E
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":210EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":21DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":22AA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2377E
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2445A
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25136
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":266EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":278A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":28782
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tool 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "iconList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "desk"
            Object.ToolTipText     =   "Go to desktop"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "drv"
            Object.ToolTipText     =   "Select a new path"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "del"
            Object.ToolTipText     =   "Put file in recycle bin"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cpu"
            Object.ToolTipText     =   "Cpu Monitor"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "prof"
            Object.ToolTipText     =   "Options : Set your Preferences"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ph"
            Style           =   3
            Object.Width           =   3500
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "gen"
            Object.ToolTipText     =   "Automaticly genearte description"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "clr"
            Object.ToolTipText     =   "Clear Descrpition"
            ImageIndex      =   29
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save Desciption"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print description"
            ImageIndex      =   32
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList iconList 
         Left            =   6360
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   39
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2905E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2993A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2A216
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2AAF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2B3CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2BCAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2C586
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2CE62
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2D73E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2E01A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2EE6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2F74A
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":30026
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":30E7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":31756
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":32032
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":32D0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":335EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":342C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":34BA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3547E
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3615A
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":36A36
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":37312
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":37BEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":388CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":395A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":39E82
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3A75E
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3B43A
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3BD16
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3C5F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3CECE
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3D7AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3E086
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3E962
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3F23E
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3FB1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":403F6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "File Info "
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Width           =   11655
      Begin VB.CommandButton SetBtn 
         Caption         =   "Set Attibutes"
         Height          =   195
         Left            =   9120
         TabIndex        =   16
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Normal"
         Height          =   255
         Left            =   10200
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Read-Only"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   10200
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox Check6 
         Caption         =   "hidden"
         Height          =   255
         Left            =   10200
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox Check5 
         Caption         =   "system"
         Height          =   255
         Left            =   10200
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Volume"
         Height          =   255
         Left            =   8760
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Directory"
         Height          =   255
         Left            =   8760
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Archive"
         Height          =   255
         Left            =   8760
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Alias"
         Height          =   255
         Left            =   8760
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.PictureBox desic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   735
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2055
      End
      Begin VB.Line Line3 
         X1              =   1200
         X2              =   4920
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   600
         Width           =   3135
      End
      Begin VB.Line Line2 
         X1              =   8640
         X2              =   8640
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   960
         X2              =   960
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Label1"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Fdes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File Description"
      Height          =   5655
      Left            =   5520
      TabIndex        =   1
      Top             =   840
      Width           =   6375
      Begin RichTextLib.RichTextBox Des 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   9340
         _Version        =   393217
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form1.frx":40CD2
      End
   End
   Begin VB.FileListBox File1 
      Height          =   5550
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   720
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Exte As String
Dim FFname As String
Dim Fpath As String

Private Sub File1_Click()
Exte = GetFileExt(File1.filename)
 If Right(File1.Path, 1) <> "\" Then FFname = File1.Path & "\" & File1.filename Else FFname = File1.Path & File1.filename
LoadIcon (Exte)
Label2.Caption = UCase(Exte)
Label1.Caption = GetFnameWITHOUText(File1.filename)
GetAttributes FFname
Check1.Value = at.Alias
Check2.Value = at.Archive
Check3.Value = at.Directory
Check4.Value = at.Volume
Check5.Value = at.System
Check6.Value = at.Hidden
Check7.Value = at.ReadOnly
Check8.Value = at.Normal
Label4.Caption = GetFileSize(FFname)
Label5.Caption = GetFileDate(FFname)
GetDescription FFname
If LCase(Exte) = "jpg" Or LCase(Exte) = "bmp" Or LCase(Exte) = "gif" Then Image1.Picture = LoadPicture(FFname) Else Image1.Picture = Nothing

End Sub

Private Sub File1_DblClick()
LaunchWithDefaultApp FFname
End Sub

Private Sub Form_Load()
If Not FileExists(App.Path & "\descr") Then CreateDir (App.Path & "\descr")


End Sub

Private Sub SetBtn_Click()
Dim temp As VbFileAttribute
temp = 0
If Check1.Value = 1 Then temp = temp + vbAlias
If Check2.Value = 1 Then temp = temp + vbArchive
If Check3.Value = 1 Then temp = temp + vbDirectory
If Check4.Value = 1 Then temp = temp + vbVolume
If Check5.Value = 1 Then temp = temp + vbSystem
If Check6.Value = 1 Then temp = temp + vbHidden
If Check7.Value = 1 Then temp = temp + vbReadOnly
If Check8.Value = 1 Then temp = temp + vbNormal
SetAttributes FFname, temp

End Sub

Private Sub Tool_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
    Case "home"
        MsgBox "Will open the defaul WebBroser to your favorite web site"
    Case "del"
        PutInRecycle FFname
        File1.Refresh
    Case "desk"
        File1.Path = GetSpecialfolder(CSIDL_DESKTOP) & "\"
    Case "drv"
        File1.Path = BrowseFF
        Label6.Caption = File1.Path & " : " & File1.ListCount & " files in directory"
    Case "cpu"
        If frmMonitor.Visible = True Then
            Unload frmMonitor
            Button.MixedState = False
            Button.ToolTipText = " load the cpu monitor"
        Else
            frmMonitor.Visible = True
            Button.MixedState = True
            Button.ToolTipText = " Unload the cpu monitor"
        End If
    Case "print"
        ppview.Show vbModal, Me
    Case "clr"
        Des.Text = ""
        Des.Refresh
    Case "save"
        If MsgBox("Do you want to save this file description ?", _
        vbQuestion + vbYesNo) = vbYes Then _
        Des.SaveFile App.Path & "\descr\" & _
        GetFnameWITHOUText(GetFileName(FFname)) & ".des"
     Case "gen"
     Beep
     Dialog.Show vbModal, Me
        
        
End Select

End Sub
Public Function LoadIcon(ext As String)
Tool.Buttons.Item(8).Enabled = True
Select Case LCase(ext)
Case "rar"
    desic.Picture = ilist.ListImages(23).Picture
    Label3.Caption = "RAR file: Compresion protocol"
Case "zip"
    desic.Picture = ilist.ListImages(49).Picture
    Label3.Caption = "ZIP file: Compresion protocol"
Case "ace"
    desic.Picture = ilist.ListImages(53).Picture
    Label3.Caption = "ACE file: Compresion protocol"
Case "txt", "rtf"
    desic.Picture = ilist.ListImages(47).Picture
    Label3.Caption = UCase(Exte) & " file:  Text file"
Case "sys", "bak", "bat"
    desic.Picture = ilist.ListImages(1).Picture
    Label3.Caption = UCase(Exte) & " file:  System file"
Case "exe"
    desic.Picture = ilist.ListImages(52).Picture
    Label3.Caption = "EXE file: Program, install or sef extracting file"
Case "des"
    desic.Picture = ilist.ListImages(30).Picture
    Label3.Caption = "DES file: Tmx description File"
Case "jpg"
    desic.Picture = ilist.ListImages(17).Picture
    Label3.Caption = "JPEG file: Image file"
Case "bmp"
    desic.Picture = ilist.ListImages(16).Picture
    Label3.Caption = "BMP file: Image file (windows)"
Case "gif"
    desic.Picture = ilist.ListImages(18).Picture
    Label3.Caption = "GIF file: Image file (compuserve)"
Case "mp3"
    desic.Picture = ilist.ListImages(15).Picture
    Label3.Caption = "MP3 file: Compressed music file"
Case "avi", "mpeg", "ram"
    desic.Picture = ilist.ListImages(14).Picture
    Label3.Caption = UCase(Exte) & " file: Video file"
Case "qtm"
    desic.Picture = ilist.ListImages(44).Picture
    Label3.Caption = "QTM file: QuickTime Movie file"
Case "lnk"
    desic.Picture = ilist.ListImages(6).Picture
    Label3.Caption = "LNK file: Shortcut File"
Case "log"
    desic.Picture = ilist.ListImages(7).Picture
    Label3.Caption = "LOG file: Some program's debug file"
'Case "vbp", "res", "frm", "ocx", "bas", "cls", "pag"

'Case "swf"
     
Case Else
    Dim tico As Long
    desic.Picture = Nothing
    desic.Cls
    tico = GetIconFromFile(FFname)
    DrawIconEx desic.hdc, 0, 0, tico, 0, 0, 0, 0, DI_NORMAL  'ilist.ListImages(36).Picture
    Label3.Caption = "Not Suported"
    DestroyIcon tico
End Select
End Function
Private Function GetDescription(filename As String)
Dim temp As String
Dim Tex As String
Tool.Buttons.Item(8).Enabled = True
Tex = GetFileExt(filename)
temp = GetFileName(filename)
temp = GetFnameWITHOUText(temp)
temp = App.Path & "\descr\" & temp & ".des"
Select Case LCase(Tex)
Case "des", "txt", "nfo", "htm"
    Des.LoadFile filename
Case "bat", "bak", "sys", "log"
    Des.LoadFile (FFname)
    Tool.Buttons.Item(8).Enabled = False

Case Else
    If FileExists(temp) Then Des.LoadFile (temp) Else Des.Text = "No Description"
End Select



End Function
