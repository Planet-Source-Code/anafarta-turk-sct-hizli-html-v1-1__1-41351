VERSION 5.00
Begin VB.Form frmyazi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Yazý Düzenleyicisi"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsymbol 
      Caption         =   "Özel Karakter Ekleyin"
      Height          =   360
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmdhr 
      Caption         =   "Yatay Çizgi"
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdlist 
      Caption         =   "Liste Yapar"
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdMarquee 
      Caption         =   "Marquee"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox chkTItle 
      Caption         =   "Baþlýk"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox chkSup 
      Caption         =   "Super Script"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Eðik(italik) Font"
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Baþlýk Büyüklükleri"
      Height          =   1215
      Left            =   720
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
      Begin VB.CheckBox chkh6 
         Caption         =   "6"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkh5 
         Caption         =   "5"
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox chkh4 
         Caption         =   "4"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Chkh3 
         Caption         =   "3"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkh2 
         Caption         =   "2"
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdBreak 
      Caption         =   "Satýr Baþý(Break)"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdp 
      Caption         =   "Yeni Paragraf"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sub Script"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox chkUnder 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Alt Çizgi Fontu"
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox chkStrike 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ortasý Çizik"
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Kalýn Font"
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmyazi"
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
Private Sub Check1_Click()
frmAna.textHTML.SelRTF = "<sub></sub>"
'selrtf text olarak deðil seçileni onun formatýnda alýr
Unload Me
End Sub

Private Sub chkBold_Click()
frmAna.textHTML.SelRTF = "<b></b>"
Unload Me
End Sub

Private Sub chkh2_Click()
frmAna.textHTML.SelRTF = "<h2></h2>"
Unload Me
End Sub

Private Sub Chkh3_Click()
frmAna.textHTML.SelRTF = "<h3></h3>"
Unload Me
End Sub

Private Sub chkh4_Click()
frmAna.textHTML.SelRTF = "<h4></h4>"
Unload Me
End Sub

Private Sub chkh5_Click()
frmAna.textHTML.SelRTF = "<h5></h5>"
Unload Me
End Sub

Private Sub chkh6_Click()
frmAna.textHTML.SelRTF = "<h6></h6>"
Unload Me
End Sub

Private Sub chkItalic_Click()
frmAna.textHTML.SelRTF = "<i></i>"
Unload Me
End Sub

Private Sub chkStrike_Click()
frmAna.textHTML.SelRTF = "<s></s>"
Unload Me
End Sub

Private Sub chkSup_Click()
frmAna.textHTML.SelRTF = "<sup></sup>"
Unload Me
End Sub

Private Sub chkTItle_Click()
frmAna.textHTML.SelRTF = "<title></title>"
Unload Me
End Sub

Private Sub chkUnder_Click()
frmAna.textHTML.SelRTF = "<u></u>"
Unload Me
End Sub

Private Sub cmdBreak_Click()
frmAna.textHTML.SelRTF = "<br>"
Unload Me
End Sub

Private Sub cmdhr_Click()
Dim inputtext
inputtext = InputBox("Yatay çigi için istediðniz bir rengi giriniz:", "Yatay Çizgi Rengi")
frmAna.textHTML.SelRTF = "<hr color=""" + inputtext + """>"
Unload Me
End Sub

Private Sub cmdlist_Click()
frmliste.Show vbModal
End Sub

Private Sub cmdMarquee_Click()
Dim inputtext
inputtext = InputBox("Kayan yazý için istediðiniz yazýyý yazýnýz:", "Kayan Yazýsý")
frmAna.textHTML.SelRTF = "<marquee>" + inputtext + "</marquee>"
Unload Me
End Sub

Private Sub cmdp_Click()
frmAna.textHTML.SelRTF = "<p>"
Unload Me
End Sub

Private Sub cmdsymbol_Click()
frmsembol.Show vbModal
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
