Attribute VB_Name = "函数-ftp及UTF"
Option Compare Database
Option Explicit

Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const scuseragent = "vb wininet"
Private Const INTERNET_FLAG_PASSIVE = &H8000000
'调用设置环境
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'连接服务器
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'上传涵数
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Dim hOpen As Long
Dim hConnection As Long

'上传文件模块
Public Function UpLoadFile(IP As String, filename As String, userName As String, PASSWORD As String) As Boolean
    Dim ShortName As String
    Dim ret As Boolean
    ShortName = GetShortName(filename)
    hOpen = TestServer
    If hOpen <> 0 Then
        hConnection = InterConnection(IP, userName, PASSWORD)
        If hConnection <> 0 Then
            ret = FtpPutFile(hConnection, filename, ShortName, 2, 0)
            UpLoadFile = ret
        Else
            UpLoadFile = False
        End If
    Else
        UpLoadFile = False
    End If
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
End Function
'FTP下载文件
Public Function DownLoadFile(IP As String, filename As String, LocalFileName As String, userName As String, PASSWORD As String)
    Dim ret As Boolean
    hOpen = TestServer
    If hOpen <> 0 Then
        hConnection = InterConnection(IP, userName, PASSWORD)
        If hConnection <> 0 Then
            ret = FtpGetFile(hConnection, filename, LocalFileName, 0, 0, 1, 0)
            DownLoadFile = ret
        Else
            DownLoadFile = False
        End If
    Else
        DownLoadFile = False
    End If
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
End Function
'调用设置环境
Private Function TestServer() As Long
    Dim i As Long
'    i = InternetOpen(scuseragent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    i = InternetOpen(vbNullString, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    TestServer = i
End Function
'连接服务器
Private Function InterConnection(IP As String, userName As String, PASSWORD As String) As Long
    Dim i As Long
    i = InternetConnect(hOpen, IP, 0, userName, PASSWORD, 1, INTERNET_FLAG_PASSIVE, 0)
    InterConnection = i
End Function
'得到文件的短文件名
Private Function GetShortName(filename As String) As String
    Dim stemp() As String
    stemp = Split(filename, "\")
    If UBound(stemp) > 0 Then
        GetShortName = stemp(UBound(stemp))
    Else
        GetShortName = ""
    End If
End Function
'输出UTF-8编码的文件
Function SaveTextAsUTF8(filepath, Text)
        Const adTypeText = 2
        Const adSaveCreateOverWrite = 2

        'Create Stream object
        Dim TextStream
        Set TextStream = CreateObject("ADODB.Stream")
        With TextStream
                .Open
                .Charset = "UTF-8"
                .position = TextStream.Size
                .WriteText Text
                .SaveToFile filepath, adSaveCreateOverWrite
                .Close
        End With
        Set TextStream = Nothing
        
End Function
