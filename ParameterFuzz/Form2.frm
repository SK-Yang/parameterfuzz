VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4260
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8724
   LinkTopic       =   "Form2"
   ScaleHeight     =   4260
   ScaleWidth      =   8724
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Del/Render"
      Height          =   192
      Left            =   7560
      TabIndex        =   5
      Top             =   3960
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   6240
      TabIndex        =   4
      Text            =   "www.dmoz.org"
      Top             =   240
      Width           =   2292
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Render"
      Height          =   192
      Left            =   6600
      TabIndex        =   3
      Top             =   3960
      Width           =   852
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3132
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   8412
      ExtentX         =   14838
      ExtentY         =   5524
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   192
      Left            =   5520
      TabIndex        =   1
      Top             =   3960
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   3132
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   8412
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents sending As CSocketMaster
Attribute sending.VB_VarHelpID = -1
Private Sub Command1_Click()
sending.CloseSck
sending.Connect Text2.Text, 80 'Conectamos al host
End Sub

Private Sub Command2_Click()
WebBrowser1.Navigate App.Path & "\Respuesta.html"

If WebBrowser1.Visible = True Then
    WebBrowser1.Visible = False
Else
WebBrowser1.Visible = True
End If

End Sub

Private Sub Command3_Click()
If Not Dir(App.Path & "/Respuesta.html") = "" Then
    Kill App.Path & "/Respuesta.html"
    MsgBox "Eliminado!"
Else

End If
End Sub

Private Sub sending_DataArrival(ByVal bytesTotal As Long)
'Dim data As String
'On Error Resume Next
'sending.GetData data
'Text1.Text = data & vbCrLf
'WebBrowser1.Navigate Text1.Text



Dim Datos$, Cabecera, Html$, Canal% 'Declaramos variables
On Error Resume Next
sending.GetData Datos 'Obtenemos datos
Text1.Text = Datos 'Ponemos contenido en la caja de texto
If InStr(1, Datos, vbCrLf & vbCrLf, vbTextCompare) <> 0 And InStr(1, Datos, "HTTP/1.1 200 OK", vbTextCompare) Then 'Si es cabecera de servidor entonces..
    Cabecera = Split(Datos, vbCrLf & vbCrLf, 2) 'Dividimos la respuesta en dos partes, la cabecera y el html
    Html = Cabecera(1) 'El html es la segunda parte, la que es luego de los dos vbCrLf
Else
    Html = Datos 'Sino es cabecera entonces el html son los datos recibidos
End If
Canal = FreeFile
Open App.Path & "\Respuesta.html" For Binary Access Write As Canal 'Abrimos el archivo de respuesta.html en modo binario
Put #Canal, LOF(Canal) + 1, Html 'Escribimos al final del archivo
Close #Canal 'Cerramos canal

End Sub


Private Sub sending_Connect()

sending.SendData Text1.Text
End Sub


Private Sub Form_Load()
'declaramos un nuevo sending
Set sending = New CSocketMaster
'Text1.Text = Form1.Text3.Text + " " + Form1.Text4.Text & Form1.Text6.Text & Form1.Text5.Text & " " & " HTTP/1.1" & vbCrLf & "Host: " & Form1.Text2.Text & vbCrLf & "Connection: keep-alive" & vbCrLf & "User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11" & vbCrLf & "Cookie: " & Form1.Text12.Text & vbCrLf & vbCrLf 'Armamos la peticion segun la estructura que te di
Text1.Text = "GET /search?q=hola HTTP/1.1" & vbCrLf & "Host: " & Text2.Text & vbCrLf & "User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.4 (KHTML, like Gecko) Chrome/22.0.1229.94 Safari/537.4" & vbCrLf & vbCrLf


End Sub
