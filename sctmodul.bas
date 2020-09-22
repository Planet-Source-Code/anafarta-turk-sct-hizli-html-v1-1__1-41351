Attribute VB_Name = "sctmodul"
' °
'                ÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛÛ
'          ÛÛ ÚÚÚÚÚÚÚ ÚÚÚÚÚÚÚÚÚÚÚÚÚÚÚÚÚÚÚtÚÛÛÛÛÛÛÛ
'            ÛÛ     ÛssssÛ   ÙÙcccÙÙÙÙ     ÚtÛ      ÛÛ
'           ÛÛ Ú   ÛsÛÛÛÛ ° ÙcÙÙÙÙÙccÙ  °  ÚtÚ       ÛÛ
'          Û     ÛÛÛÛ      ÙÙÙÙ    ÙcÙ     ÚtÚ
'            º  ÛÛsÛ  °    ÙcÙ      ÙÙ  °  ÚtÚ      º
'   ¹           ÛsÛs      ÙcÙÙ             ÚtÚ
'                ÛÛssÛs   ÙccÙ             ÚtÚ
'                  ÛÛss   ÙÙcÙ         ¹   ÚtÚ
'       º    º     ÛsÛÛ    ÙÙcÙ    ÙÙÙ     ÚtÚ
'                  ÛssÛ     ÙcÙ   ÙÙÙÙ  °  ÚtÚ   °
'                ÛÛsÛÛ   °   ÙcÙÙÙÙcÙ      ÚtÚ      º
'    °          ÛssÛ °       Ùccc cÙÙ      ÚtÚ    ¹
'               sÛÛ           ÙÙÙÙÙÙ       ÚtÚ
'              ÛÛ   SOLDiER CRACKERS TEAM  ÚÚÚ
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*******************************************************'
'*    proje: [SCT] Hýzlý HTML Editörü                  *'
'*    yazar: Anafarta Türk                             *'
'*  e-posta: blau_devil@hotmail.com                    *'
'*      web: http://www.sct.tr.cx/                     *'
'*    tarih: 30.11.2002                                *'
'*******************************************************'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'shellexecute windowsta shell32.dll dosyasýný->
'kullanarak bize formumuzda bir internet adresini->
'mail adresini veya bir yolu explorerda açmamýzý saðlýyor.

Public Function OpenIt(frm As Form, ToOpen As String)
'KULLANIM: OPENIT "c:\windows\notepad.exe"
'KULLANIM: OPENIT "http://www.sct.tr.cx"
'KULLANIM: OPENIT "http://www.turkey.com"
'KULLANIM: OPENIT "mailto: blau_devil@hotmail.com"
ShellExecute frm.hwnd, "Open", ToOpen, &O0, &O0, SW_NORMAL
End Function

