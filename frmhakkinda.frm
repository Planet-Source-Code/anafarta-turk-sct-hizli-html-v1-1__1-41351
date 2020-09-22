VERSION 5.00
Begin VB.Form frmhakkinda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hýzlý HTML Hakkýnda"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin hizlihtml.Command Command2 
      Height          =   1215
      Left            =   4440
      TabIndex        =   11
      Top             =   960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   2143
      Caption         =   "Tamam"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Önemli :"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   5055
      Begin VB.Label Label6 
         Caption         =   $"frmhakkinda.frx":0000
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmhakkinda.frx":00BB
      ScaleHeight     =   855
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label9 
      Caption         =   "www.sct.tr.cx"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "blau_devil@hotmail.com"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Blue Devil"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Web"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "E-Posta"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Coder:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright © 1999-2003"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "[SCT] Hýzlý HTML"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmhakkinda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' °                                                     '
'                ÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛ   '
'          ÛÛ ÚÚÚÚÚÚÚ ÚÚÚÚÚÚÚÚÚÚÚÚÚÚÚÚÚÚÚtÚÛÛÛÛÛÛÛ      '
'            ÛÛ     ÛssssÛ   ÙÙcccÙÙÙÙ     ÚtÛ      ÛÛ  '
'           ÛÛ Ú   ÛsÛÛÛÛ ° ÙcÙÙÙÙÙccÙ  °  ÚtÚ       ÛÛ '
'          Û     ÛÛÛÛ      ÙÙÙÙ    ÙcÙ     ÚtÚ          '
'            º  ÛÛsÛ  °    ÙcÙ      ÙÙ  °  ÚtÚ      º   '
'   ¹           ÛsÛs      ÙcÙÙ             ÚtÚ          '
'                ÛÛssÛs   ÙccÙ             ÚtÚ          '
'                  ÛÛss   ÙÙcÙ         ¹   ÚtÚ          '
'       º    º     ÛsÛÛ    ÙÙcÙ    ÙÙÙ     ÚtÚ          '
'                  ÛssÛ     ÙcÙ   ÙÙÙÙ  °  ÚtÚ   °      '
'                ÛÛsÛÛ   °   ÙcÙÙÙÙcÙ      ÚtÚ      º   '
'    °          ÛssÛ °       Ùccc cÙÙ      ÚtÚ    ¹     '
'               sÛÛ           ÙÙÙÙÙÙ       ÚtÚ          '
'              ÛÛ   SOLDiER CRACKERS TEAM  ÚÚÚ          '
'                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*******************************************************'
'*  project: [SCT] Hýzlý HTML Editörü                  *'
'*   author: Anafarta Türk                             *'
'*   e-mail: blau_devil@hotmail.com                    *'
'*      web: http://www.sct.tr.cx/                     *'
'*     date: 30.11.2002                                *'
'*******************************************************'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
