VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAna 
   Caption         =   "[SCT] Hýzlý HTML"
   ClientHeight    =   6930
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9525
   Icon            =   "frmAna.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin hizlihtml.Command cmdyazi 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Yazý Düzenle"
   End
   Begin hizlihtml.Command cmdlink 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Link Ekle"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   -480
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox textHTML 
      Height          =   3000
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5292
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmAna.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Menu mndosya 
      Caption         =   "&Dosya"
      Begin VB.Menu mnyeni 
         Caption         =   "Yeni Sayfa"
      End
      Begin VB.Menu mnac 
         Caption         =   "Sayfa Aç"
      End
      Begin VB.Menu mnkaydet 
         Caption         =   "Kaydet"
      End
      Begin VB.Menu mnkesme1 
         Caption         =   "-"
      End
      Begin VB.Menu mncikis 
         Caption         =   "Çýkýþ"
      End
   End
   Begin VB.Menu mnduzen 
      Caption         =   "Dü&zen"
      Begin VB.Menu mnkes 
         Caption         =   "Kes"
      End
      Begin VB.Menu mnkopyala 
         Caption         =   "Kopyala"
      End
      Begin VB.Menu mnyapistir 
         Caption         =   "Yapýþtýr"
      End
      Begin VB.Menu mntumusec 
         Caption         =   "Tümünü Seç"
      End
   End
   Begin VB.Menu mnhakkinda 
      Caption         =   "Hakkýnda"
      Begin VB.Menu mnhizlihtml 
         Caption         =   "Hýzlý HTML Hakkýnda"
      End
      Begin VB.Menu mnsctweb 
         Caption         =   "Sct Web Sitesi"
      End
   End
End
Attribute VB_Name = "frmAna"
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
Option Explicit
'Bu deðiþkenler Undo ve Redo için (geri al veya tersi için)
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(100000) As String

Private Sub cmdlink_Click()
frmlinkler.Show vbModal
End Sub

Private Sub cmdyazi_Click()
frmyazi.Show vbModal
End Sub

Private Sub Form_Load()
File1.FileName = "*.htm*" 'fileda sadece htm ve html uzantýlý dosyalarý gösterir
Command1.Visible = False 'bunun amacý sadece rtbyi formun büyüklüðüne ayar yapmak için kullanancaz
End Sub

Private Sub Drive1_Change()
On Local Error GoTo hata 'hata olursa "hata" yazan yere git
Dir1.Path = Drive1.Drive 'sürücü deðiþimi
'dizin yolu sürücüye eþitlendi
ChDrive Drive1.Drive 'aktif sürücüyü deðiþtirir
Exit Sub 'hata olmazsa aþaðýya devam etme
hata: 'hata isimli yer
MsgBox ("Hata") 'hata mesajý
Exit Sub 'alt programdan çýk
End Sub

Private Sub Dir1_Change()
On Local Error GoTo hata 'hata olursa "hata" yazan yere git
File1.Path = Dir1.Path 'dizin seçildiðinde dosyalarý göster
ChDir Dir1.Path 'Aktif dizini deðiþtirir
Exit Sub 'hata olmazsa aþaðýya devam etme
hata: 'hata isimli yer
MsgBox ("Dizin Bulunamadý")
Exit Sub 'alt programdan çýk
End Sub

Private Sub File1_Click()
Dim filenumber
    On Error Resume Next
    filenumber = FreeFile
    
    Open File1.FileName For Input As #filenumber 'dosya1deki dosyayý açar filename i kullandýðýmýza dikkat edelim
    textHTML.Text = Input(LOF(filenumber), #filenumber)
    Close
End Sub

Private Sub Form_Resize() 'formun büyüyüp küçülmesine göre
'texthtml ninde boyutlanmasýný saðladýk
On Local Error Resume Next
textHTML.Move 2000, 100, ScaleWidth - 2000, ScaleHeight - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnac_Click()
Dim filenumber
On Error Resume Next
filenumber = FreeFile 'iþte bununla kodlarý görebiliyoruz:)
cd1.DialogTitle = "Aç" 'dialogun baþlýðý
cd1.Filter = "*.htm|*.htm|*.html|*.html" 'açýlacak dosya uzantýlarý
cd1.ShowOpen 'common dialodun aç formu
Open cd1.FileName For Input As #filenumber
textHTML.Text = Input(LOF(filenumber), #filenumber)
Close
End Sub

Private Sub mngerial_Click()

'This says that if the Index is = to 0,
'then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
   textHTML.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub mncikis_Click()
End
End Sub

Private Sub mnhizlihtml_Click()
frmhakkinda.Show vbModal
End Sub

Private Sub mnkaydet_Click()
Dim filenumber As Integer
On Local Error Resume Next
filenumber = FreeFile
cd1.DialogTitle = "Kaydet"
cd1.Filter = "*.htm|*.htm|*.html|*.html"
cd1.ShowSave
Open cd1.FileName For Output As #filenumber
Print #filenumber, frmAna.textHTML.Text
Close #filenumber
End Sub

Private Sub mnkes_Click()
If textHTML.SelText = "" Then
        Exit Sub
    Else
        Clipboard.Clear
        Clipboard.SetText textHTML.SelText
        textHTML.SelText = ""
    End If
End Sub

Private Sub mnkopyala_Click()
If textHTML.SelText = "" Then
        Exit Sub
    Else
        Clipboard.Clear
        Clipboard.SetText textHTML.SelText
    End If
End Sub

Private Sub mnsctweb_Click()
OpenIt Me, "http://www.sct.tr.cx"
End Sub

Private Sub mnyapistir_Click()
 textHTML.SelText = Clipboard.GetText
End Sub

Private Sub mntumusec_Click()
    'Sets the cursors position to zero
    textHTML.SelStart = 0
    'Selects the full length of rtfText
    textHTML.SelLength = Len(textHTML.Text)
    'Sets the Focus to rtfText
   textHTML.SetFocus

End Sub

Private Sub mnyeni_Click()
textHTML.Text = ""
End Sub
