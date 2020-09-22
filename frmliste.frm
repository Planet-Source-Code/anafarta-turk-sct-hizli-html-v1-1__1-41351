VERSION 5.00
Begin VB.Form frmliste 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liste Düzenleyicisi"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin hizlihtml.Command cmdinsert 
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Caption         =   "Tamam"
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   50
      Min             =   1
      TabIndex        =   6
      Top             =   2040
      Value           =   1
      Width           =   4455
   End
   Begin VB.TextBox txtlist 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.OptionButton optbullet 
      Caption         =   "Noktalý Liste(Bullet)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
   Begin VB.OptionButton optsquare 
      Caption         =   "Kare Noktalý Liste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.OptionButton optnumber 
      Caption         =   "Numaralý Liste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.OptionButton optroman 
      Caption         =   "Roman Rakamlý Liste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.OptionButton optletter 
      Caption         =   "Abeceli Liste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Liste türü seç"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Liste öðelerinin sayýsýný seçiniz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   1320
      Width           =   3015
   End
End
Attribute VB_Name = "frmliste"
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
Private Sub cmdinsert_Click()
Dim X As Long, Y
    ' NOKTA (BULLET) LÝSTESÝ
    If optbullet.Value = True Then
    frmAna.textHTML.SelRTF = "<UL>" + vbCrLf
        Do
            X = txtlist.Text
            frmAna.textHTML.SelRTF = "<LI> Öðelerinizi burada listeleyin </LI>"
            Y = Y + 1
        Loop While Y < X
    frmAna.textHTML.SelRTF = "</UL>" + vbCrLf
    End If
    ' ALFABETÝK SIRAYLA HARF LÝSTESÝ
    If optletter.Value = True Then
    frmAna.textHTML.SelRTF = "<OL TYPE=" + Chr(34) + "A" + Chr(34) + ">" + vbCrLf
        Do
            X = txtlist.Text
            frmAna.textHTML.SelRTF = "<LI> Öðelerinizi burada listeleyin </LI>"
            Y = Y + 1
        Loop While Y < X
    frmAna.textHTML.SelRTF = "</OL>" + vbCrLf
    End If
' SAYI LÝSTESÝ
If optnumber.Value = True Then
frmAna.textHTML.SelRTF = "<OL>" + vbCrLf
Do
X = txtlist.Text
frmAna.textHTML.SelRTF = "<LI> Öðelerinizi burada listeleyin </LI>"
Y = Y + 1
Loop While Y < X
frmAna.textHTML.SelRTF = "</OL>" + vbCrLf
End If
' KARELÝ LÝSTE
If optsquare.Value = True Then
frmAna.textHTML.SelRTF = "<UL TYPE=" + Chr(34) + "square" + Chr(34) + ">" + vbCrLf
Do
X = txtlist.Text
frmAna.textHTML.SelRTF = "<LI> Öðelerinizi burada listeleyin </LI>"
Y = Y + 1
Loop While Y < X
frmAna.textHTML.SelRTF = "</UL>" + vbCrLf
End If
' ROMAN RAKAMLI LÝSTE
If optroman.Value = True Then
frmAna.textHTML.SelRTF = "<OL TYPE=" + Chr(34) + "I" + Chr(34) + ">" + vbCrLf
Do
X = txtlist.Text
frmAna.textHTML.SelRTF = "<LI> Öðelerinizi burada listeleyin </LI>"
Y = Y + 1
Loop While Y < X
frmAna.textHTML.SelRTF = "</OL>" + vbCrLf
End If
Unload Me
End Sub

Private Sub Form_Load()
    txtlist.Text = HScroll1.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub HScroll1_Change()
    txtlist.Text = HScroll1.Value
End Sub


