VERSION 5.00
Begin VB.Form frmlinkler 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Link ve Resim Linki D�zenleyicisi"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin hizlihtml.Command cmdIMGinsert 
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Tamam"
   End
   Begin hizlihtml.Command cmdclear 
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Temizle"
   End
   Begin hizlihtml.Command cmdimgopen 
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   "..."
   End
   Begin hizlihtml.Command cmdLink 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Tamam"
   End
   Begin VB.Frame Frame5 
      Caption         =   "Resim Linki Se�enekleri"
      Height          =   2775
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Width           =   2295
      Begin VB.TextBox txtborder 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Text            =   "Border B�y�kl���"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtlinkalt 
         Height          =   405
         Left            =   120
         TabIndex        =   7
         Text            =   "ALT Yaz�s�"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtimagelink 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Text            =   "Resim Linki"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtImage 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Text            =   "Resim Yolu"
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Link Se�enekleri"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.CheckBox chkNou 
         Caption         =   "Alt �izgi Yok"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtLink 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "Link Yaz�s�"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtahref 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Link URL"
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmlinkler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' �                                                     '
'                ������������������������������������   '
'          �� ������� �������������������t��������      '
'            ��     �ssss�   ��ccc����     �t�      ��  '
'           �� �   �s���� � �c�����cc�  �  �t�       �� '
'          �     ����      ����    �c�     �t�          '
'            �  ��s�  �    �c�      ��  �  �t�      �   '
'   �           �s�s      �c��             �t�          '
'                ��ss�s   �cc�             �t�          '
'                  ��ss   ��c�         �   �t�          '
'       �    �     �s��    ��c�    ���     �t�          '
'                  �ss�     �c�   ����  �  �t�   �      '
'                ��s��   �   �c����c�      �t�      �   '
'    �          �ss� �       �ccc c��      �t�    �     '
'               s��           ������       �t�          '
'              ��   SOLDiER CRACKERS TEAM  ���          '
'                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*******************************************************'
'*  project: [SCT] H�zl� HTML Edit�r�                  *'
'*   author: Anafarta T�rk                             *'
'*   e-mail: blau_devil@hotmail.com                    *'
'*      web: http://www.sct.tr.cx/                     *'
'*     date: 30.11.2002                                *'
'*******************************************************'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdclear_Click()
txtImage.Text = ""
txtimagelink.Text = ""
txtlinkalt.Text = ""
txtborder.Text = ""
End Sub

Private Sub cmdIMGinsert_Click()
frmAna.textHTML.SelRTF = "<a href=""" + txtimagelink.Text + """>" + "<img src=""" + txtImage.Text + """ border=""" + txtborder.Text + """ alt=""" + txtlinkalt.Text + """>" + "</a>"
'frmana.textHTML. diye ba�lamam�z�n sebebi textHTMLnin frmlinklerde de�il
'frmAna �zerinde olmas�.e�er koymazsan�z hata verir
Unload Me 'resim linkini koyduktan sonra formu kapat�r
End Sub

Private Sub cmdimgopen_Click()
frmAna.cd1.Filter = "JPG Files(*.jpg)|*.jpg|All files(*.*)|*.*"
frmAna.cd1.ShowOpen
'frmana.cd1. olmas�n� sebebi commondiyalog kutusunun
'frmAna �zerinde olmas�d�r
On Error Resume Next
txtImage.Text = "file://" + frmAna.cd1.FileName

End Sub

Private Sub cmdlink_Click()
If chkNou.Value = 0 Then
frmAna.textHTML.SelRTF = "<a href=""" + txtahref.Text + """>" + txtLink.Text + "</a>"
Else
frmAna.textHTML.SelRTF = "<a href=""" + txtahref.Text + """ style=text-decoration:none>" + txtLink.Text + "</a>"
End If
Unload Me 'resim linkini koyduktan sonra formu kapat�r
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
