Attribute VB_Name = "sctmodul"
' �
'                ������������������������������������
'          �� ������� �������������������t��������
'            ��     �ssss�   ��ccc����     �t�      ��
'           �� �   �s���� � �c�����cc�  �  �t�       ��
'          �     ����      ����    �c�     �t�
'            �  ��s�  �    �c�      ��  �  �t�      �
'   �           �s�s      �c��             �t�
'                ��ss�s   �cc�             �t�
'                  ��ss   ��c�         �   �t�
'       �    �     �s��    ��c�    ���     �t�
'                  �ss�     �c�   ����  �  �t�   �
'                ��s��   �   �c����c�      �t�      �
'    �          �ss� �       �ccc c��      �t�    �
'               s��           ������       �t�
'              ��   SOLDiER CRACKERS TEAM  ���
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*******************************************************'
'*    proje: [SCT] H�zl� HTML Edit�r�                  *'
'*    yazar: Anafarta T�rk                             *'
'*  e-posta: blau_devil@hotmail.com                    *'
'*      web: http://www.sct.tr.cx/                     *'
'*    tarih: 30.11.2002                                *'
'*******************************************************'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'shellexecute windowsta shell32.dll dosyas�n�->
'kullanarak bize formumuzda bir internet adresini->
'mail adresini veya bir yolu explorerda a�mam�z� sa�l�yor.

Public Function OpenIt(frm As Form, ToOpen As String)
'KULLANIM: OPENIT "c:\windows\notepad.exe"
'KULLANIM: OPENIT "http://www.sct.tr.cx"
'KULLANIM: OPENIT "http://www.turkey.com"
'KULLANIM: OPENIT "mailto: blau_devil@hotmail.com"
ShellExecute frm.hwnd, "Open", ToOpen, &O0, &O0, SW_NORMAL
End Function

