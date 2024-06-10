VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Public Enum LongPtr
        [_]
    End Enum

    Private Declare Sub ZipInit Lib "ZipArchive.dll" ()
    Private Declare Function ZipDirectory Lib "ZipArchive.dll" (ByVal ziparchive As LongPtr, ByVal directory As LongPtr) As Boolean
    Private Declare Sub OpenZipFile Lib "ZipArchive.dll" (ByVal ziparchive As LongPtr)
    Private Declare Function ZipFileCount Lib "ZipArchive.dll" () As Long
    Private Declare Function IsValidZip Lib "ZipArchive.dll" (ByVal ziparchive As LongPtr) As Boolean
    Private Declare Function UnCompressZipFile Lib "ZipArchive.dll" (ByVal desdirectory As LongPtr) As Boolean
    Private Declare Sub GetFileName Lib "ZipArchive.dll" (ByVal index As Long, ByRef filename As LongPtr)
    Private Declare Sub ReadTextFile Lib "ZipArchive.dll" (ByVal Filenameinzip As LongPtr, ByRef textResult As LongPtr, ByRef Length As Long)
    Private Declare Sub ExtractFile Lib "ZipArchive.dll" (ByVal Filenameinzip As LongPtr, ByVal Despath As LongPtr)
    Private Declare Function GetEntryIndex Lib "ZipArchive.dll" (ByVal Filenameinzip As LongPtr) As Long
    Private Declare Sub CloseZipFile Lib "ZipArchive.dll" ()
    Private Declare Sub ZipFree Lib "ZipArchive.dll" ()
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
Private Function PtrToStr(ByVal ptr As LongPtr) As Byte() '�ַ�������ָ�붼�������������
    Dim buffer() As Byte
    Dim n As Long
    n = lstrlenW(ptr) * 2
    ReDim buffer(0 To n - 1)
    ' �����ڴ浽��ȫ���ֽ�����
    CopyMemory buffer(0), ByVal ptr, n
    ' ���ֽ�����ת��Ϊ�ַ���
    PtrToStr = buffer()
End Function
Private Function bytePtrToStr(ByVal ptr As LongPtr, ByVal bytelen As Long) As Byte() '���������ReadTextFile��ȡѹ�������ݵ��ֽ����顣��Ϊֱ���ַ������漰���ļ��������⡣���԰ѱ��������vba������
    Dim buffer() As Byte
    Dim n As Long
    n = bytelen
    ReDim buffer(0 To n - 1)
    ' �����ڴ浽��ȫ���ֽ�����
    CopyMemory buffer(0), ByVal ptr, n
    ' ָ������ת��Ϊ�ֽ�����
   bytePtrToStr = buffer()
End Function

Private Sub Form_Load()
Dim cc As Long, PicCount&, i&, fn As LongPtr, ZipFilePath$, Filenameinzip As LongPtr, unzipFilepath$, Drawingxml$, Fileindex As Long, szText As LongPtr, Textbuffer() As Byte, Fso As Object, fz As String, Textlen&
Dim xmlDom As Object, nodes As Object, pos(), DrawingRelationship, Filenameinzip1 As LongPtr, Fileindex1 As Long, szText1 As LongPtr, Textbuffer1() As Byte, dic As Object, h&, ID$, target$, Textlen1&
Dim xmlLoaded As Boolean
ChDrive App.Path
#If Win64 Then
    ChDir App.Path & "\win64"
#Else
    ChDir App.Path & "\win32"
#End If
Set Fso = CreateObject("Scripting.FileSystemObject")
ZipFilePath = App.Path & "\ͼƬԴ.xlsx"
If Len(Dir(ZipFilePath)) = 0 Then MsgBox "xlsx�ļ�������": Exit Sub
ZipInit
If IsValidZip(StrPtr(ZipFilePath)) Then Debug.Print "����һ����Ч��zipѹ���ĵ�" Else GoTo label1
unzipFilepath = App.Path & "\ͼƬ" '���ͼƬ��λ��
If Len(Dir(unzipFilepath, vbDirectory)) Then Fso.deletefolder unzipFilepath
OpenZipFile StrPtr(ZipFilePath)
MkDir unzipFilepath
cc = ZipFileCount
For i = 0 To cc - 1
    GetFileName i, fn '��ȡѹ��Ŀ¼�ļ���ָ��
    fz = PtrToStr(fn) 'ָ��ת��Ϊ�ַ���
    If InStr(fz, "xl/media/image") Then PicCount = PicCount + 1: ExtractFile fn, StrPtr(unzipFilepath)
Next i
    ReDim pos(1 To PicCount, 1 To 3)
    Filenameinzip = StrPtr("xl/drawings/drawing1.xml")  '����ļ��ﺬ��ͼƬ��λ�ú�id(r:embed����ֵ)��Ϣ
    Fileindex = GetEntryIndex(Filenameinzip)
    If Fileindex = -1 Then MsgBox "��Ҫ���ʵ�ѹ��������ļ�������": GoTo label2
    ReadTextFile Filenameinzip, szText, Textlen '��ȡxml�ļ�����ָ��
     '����ָ��ת��Ϊ�ַ�����xml�ı���Ϊutf-8
    Textbuffer() = bytePtrToStr(szText, Textlen)
    Drawingxml = zm(Textbuffer(), "UTF-8")
label2:
    CloseZipFile
label1:
    ZipFree
    MsgBox Drawingxml
End Sub

Function zm(ByRef arr() As Byte, ByVal encoding As String) As String
Dim Stream As Object
Set Stream = CreateObject("ADODB.Stream")
With Stream
    .Type = 1
    .Mode = 3
    .Open
    .Write arr()
    .Position = 0
    .Type = 2
    .Charset = encoding
    zm = .ReadText
    .Close
End With
 Set Stream = Nothing
End Function

