VERSION 5.00
Begin VB.Form frmsembol 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Özel Karakterler"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin hizlihtml.Command cmdclose 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Kapat"
   End
   Begin hizlihtml.Command cmdinsert 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Tamam"
   End
   Begin VB.ListBox List1 
      Columns         =   15
      Height          =   3960
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "§ýmbõl|Ér"
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bir tane özel karakter seçin ve tamama týklayýn."
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   4320
         Width           =   3375
      End
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      BorderWidth     =   5
      Height          =   5505
      Left            =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmsembol"
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
Private Sub Command1_Click()
Clipboard.SetText (List1)
txtWord.Text = txtWord.Text + List1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
MsgBox "Never Ever Simpler! Just Select a symbol from list hit copy and it will automatically goto the textbox also to your clipboard so have fun!"
End Sub

Private Sub Command4_Click()
MsgBox "Sup this is a program to create cool symbols for Stacraft/Edit, made by SkaterRob from SCMaps.com a great website! This Program is Copyright ©2001 SCMaps.com"
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdinsert_Click()
frmAna.textHTML.SelRTF = List1
End Sub

Private Sub Form_Load()
List1.AddItem "À"
List1.AddItem "à"
List1.AddItem "Á"
List1.AddItem "á"
List1.AddItem "Â"
List1.AddItem "â"
List1.AddItem "Ã"
List1.AddItem "ã"
List1.AddItem "Ä"
List1.AddItem "ä"
List1.AddItem "Å"
List1.AddItem "å"
List1.AddItem "Æ"
List1.AddItem "æ"
List1.AddItem "Ç"
List1.AddItem "ç"
List1.AddItem "Œ"
List1.AddItem "œ"
List1.AddItem "È"
List1.AddItem "è"
List1.AddItem "É"
List1.AddItem "é"
List1.AddItem "Ê"
List1.AddItem "ê"
List1.AddItem "Ë"
List1.AddItem "ë"
List1.AddItem "Ì"
List1.AddItem "ì"
List1.AddItem "Í"
List1.AddItem "í"
List1.AddItem "Î"
List1.AddItem "Î"
List1.AddItem "Ï"
List1.AddItem "ï"
List1.AddItem "Ð"
List1.AddItem "ð"
List1.AddItem "Ñ"
List1.AddItem "ñ"
List1.AddItem "Ò"
List1.AddItem "ò"
List1.AddItem "Ó"
List1.AddItem "ó"
List1.AddItem "Ô"
List1.AddItem "ô"
List1.AddItem "Õ"
List1.AddItem "õ"
List1.AddItem "Ö"
List1.AddItem "ö"
List1.AddItem "Ø"
List1.AddItem "ø"
List1.AddItem "Ù"
List1.AddItem "ù"
List1.AddItem "Ú"
List1.AddItem "ú"
List1.AddItem "Û"
List1.AddItem "û"
List1.AddItem "Ü"
List1.AddItem "ü"
List1.AddItem "Ý"
List1.AddItem "ý"
List1.AddItem "Þ"
List1.AddItem "þ"
List1.AddItem "ß"
List1.AddItem "ÿ"
List1.AddItem "±"
List1.AddItem "¡"
List1.AddItem "²"
List1.AddItem "¢"
List1.AddItem "³"
List1.AddItem "£"
List1.AddItem "´"
List1.AddItem "¤"
List1.AddItem "µ"
List1.AddItem "¥"
List1.AddItem "¶"
List1.AddItem "¦"
List1.AddItem "·"
List1.AddItem "§"
List1.AddItem "¸"
List1.AddItem "¨"
List1.AddItem "¹"
List1.AddItem "©"
List1.AddItem "º"
List1.AddItem "ª"
List1.AddItem "»"
List1.AddItem "«"
List1.AddItem "¼"
List1.AddItem "¬"
List1.AddItem "½"
List1.AddItem "­"
List1.AddItem "¾"
List1.AddItem "®"
List1.AddItem "¿"
List1.AddItem "¯"
List1.AddItem "×"
List1.AddItem "™"
List1.AddItem "¢"
List1.AddItem "÷"
List1.AddItem "ƒ"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
