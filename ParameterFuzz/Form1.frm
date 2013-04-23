VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Socket.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ParameterFuzz -- //EnElPC.com\\"
   ClientHeight    =   10164
   ClientLeft      =   2676
   ClientTop       =   768
   ClientWidth     =   14412
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10164
   ScaleWidth      =   14412
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command33 
      Caption         =   "Command4"
      Height          =   612
      Left            =   1680
      TabIndex        =   37
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cookie"
      Height          =   732
      Left            =   3360
      TabIndex        =   35
      Top             =   840
      Width           =   4092
      Begin VB.TextBox Text12 
         Height          =   288
         Left            =   120
         TabIndex        =   36
         Text            =   "PHPSESSID=TipicalSpanish"
         Top             =   240
         Width           =   3732
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2280
      Top             =   8640
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   960
      TabIndex        =   33
      Text            =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      Caption         =   "-"
      Height          =   255
      Left            =   8760
      TabIndex        =   30
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   600
      TabIndex        =   29
      Text            =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Desconectar"
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   9720
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Enviar Petición"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Borrar y Desconectar"
      Height          =   255
      Left            =   11880
      TabIndex        =   24
      Top             =   9480
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   8400
      TabIndex        =   19
      Text            =   "80"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Text            =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   12360
      TabIndex        =   17
      Text            =   "10"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "-"
      Height          =   255
      Left            =   12000
      TabIndex        =   16
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "+"
      Height          =   255
      Left            =   12000
      TabIndex        =   15
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Automático"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9120
      TabIndex        =   12
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ver HTML Local"
      Height          =   255
      Left            =   9600
      TabIndex        =   10
      Top             =   9480
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   7680
      TabIndex        =   9
      Text            =   "?Variable="
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   10800
      TabIndex        =   8
      Text            =   "abcdefghij"
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5400
      TabIndex        =   5
      Text            =   "dmoz.org"
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   10080
      TabIndex        =   4
      Text            =   "/search"
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Text            =   "GET"
      Top             =   360
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Left            =   2760
      Top             =   8640
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3240
      Top             =   8640
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Descargar HTML Local"
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   9480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   8040
      Width           =   13695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   9360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volcado HTML Local"
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   7800
      Width           =   14175
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2280
         TabIndex        =   32
         Top             =   1680
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5415
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   14175
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   5052
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   13932
         ExtentX         =   24574
         ExtentY         =   8911
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Valor"
      Height          =   735
      Left            =   10680
      TabIndex        =   22
      Top             =   840
      Width           =   3255
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   1200
         X2              =   1200
         Y1              =   240
         Y2              =   600
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Numérico Incremental"
         Height          =   375
         Left            =   2280
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Parametro"
      Height          =   735
      Left            =   7560
      TabIndex        =   26
      Top             =   840
      Width           =   3015
      Begin VB.CommandButton Command10 
         Caption         =   "+"
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   1800
      Width           =   11175
   End
   Begin VB.Label Label4 
      Caption         =   "ParameterFuzz"
      BeginProperty Font 
         Name            =   "Realvirtue"
         Size            =   28.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   2892
   End
   Begin VB.Label Label3 
      Caption         =   "Puerto:"
      Height          =   255
      Left            =   7800
      TabIndex        =   20
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo:"
      Height          =   375
      Index           =   1
      Left            =   9360
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Petición:"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Host:"
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim hostcaja As String


Private Sub Command1_Click()
hostcaja = Text2.Text 'defino el host a mandar la petición
On Error Resume Next
Winsock1.Connect hostcaja, Text9.Text 'Conectamos al host
Label6 = "Conectado!"
Label6.ForeColor = vbGreen
End Sub
 
Private Sub Command10_Click()
Text10.Text = Text10.Text + 1
Dim lectura As String
Open App.Path & "\Parametros.txt" For Input As #1
For cuenta = 1 To Text10.Text
Line Input #1, lectura
On Error Resume Next
Next cuenta
Close #1
Text6.Text = lectura
End Sub

Private Sub Command11_Click()
Text10.Text = Text10.Text - 1
Dim lectura As String
Open App.Path & "\Parametros.txt" For Input As #1
For cuenta = 1 To Text10.Text
Line Input #1, lectura
On Error Resume Next
Next cuenta
Close #1
Text6.Text = lectura
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim Peticion$

If Check1.Value = 0 Then

hostcaja = Text2.Text 'defino el host a mandar la petición
Peticion = Text3.Text + " " + Text4.Text & Text6.Text & Text5.Text & " " & " HTTP/1.1" & vbCrLf & "Host: " & hostcaja & vbCrLf & "Connection: close" & vbCrLf & "User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11" & vbCrLf & "Cookie: " & Text12.Text & vbCrLf & vbCrLf 'Armamos la peticion segun la estructura que te di
Winsock1.SendData Peticion 'Mandamos la peticion
    Else
    
Text8.Text = Text8.Text + 1
    Dim lectura As String
Open App.Path & "\Parametros.txt" For Input As #1
For cuenta = 1 To Text8.Text
Line Input #1, lectura
Next cuenta
'MsgBox lectura
Close #1
    
    Text6.Text = lectura
    
hostcaja = Text2.Text 'defino el host a mandar la petición
Peticion = Text3.Text + " " + Text4.Text & Text6.Text & Text5.Text & " " & " HTTP/1.1" & vbCrLf & "Host: " & hostcaja & vbCrLf & "Connection: close" & vbCrLf & "User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11" & vbCrLf & "Cookie: " & Text12.Text & vbCrLf & vbCrLf 'Armamos la peticion segun la estructura que te di
Winsock1.SendData Peticion 'Mandamos la peticion
End If

'Dim Peticion$
'hostcaja = Text2.Text 'defino el host a mandar la petición

'Peticion = Text3.Text + " " + Text4.Text & Text6.Text & Text5.Text & " " & " HTTP/1.1" & vbCrLf & "Host: " & hostcaja & vbCrLf & "Connection: close" & vbCrLf & "User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11" & vbCrLf & vbCrLf 'Armamos la peticion segun la estructura que te di

'Winsock1.SendData Peticion 'Mandamos la peticion


End Sub

Private Sub Command3_Click()
WebBrowser1.Navigate App.Path & "\Respuesta.html"
End Sub

Private Sub Command4_Click()

If Not Dir(App.Path & "/Respuesta.html") = "" Then
    Kill App.Path & "/Respuesta.html"
Else

End If

End Sub

Private Sub Command33_Click()
Form2.Show
End Sub

Private Sub Command5_Click()
Text5.Text = Val(Text5.Text) + Val(Text7.Text)
End Sub

Private Sub Command6_Click()
Text5.Text = Text5.Text - Text7.Text
End Sub

Private Sub Command7_Click()
Winsock1.Close
Label6 = "Desconectado!"
Label6.ForeColor = vbRed
End Sub

Private Sub Command8_Click()
Text1.Text = ""
WebBrowser1.Navigate ""
If Not Dir(App.Path & "/Respuesta.html") = "" Then
    Kill App.Path & "/Respuesta.html"
    MsgBox "Eliminado!"
Else

End If

Winsock1.Close
Label6 = "Desconectado!"
Label6.ForeColor = vbRed

End Sub

Private Sub Command9_Click()
If Check1.Value = 0 Then

WebBrowser1.Navigate "http://" & Text2.Text & ":" & Text9.Text & Text4.Text & Text6.Text & Text5.Text
Label7.Caption = "http://" & Text2.Text & ":" & Text9.Text & Text4.Text & Text6.Text & Text5.Text

    Else
    
    
Text11.Text = Text11.Text + 1
Dim lectura As String
Open App.Path & "\Parametros.txt" For Input As #1
For cuenta = 1 To Text11.Text
On Error Resume Next
Line Input #1, lectura
Next cuenta
Close #1
    
    Text6.Text = lectura
    
       
WebBrowser1.Navigate "http://" & Text2.Text & ":" & Text9.Text & Text4.Text & Text6.Text & Text5.Text
Label7.Caption = "http://" & Text2.Text & ":" & Text9.Text & Text4.Text & Text6.Text & Text5.Text

    End If
End Sub



Private Sub Form_Load()
Label6 = "Desconectado!"
Label6.ForeColor = vbRed
End Sub

Private Sub Timer1_Timer()
If Winsock1.State = sckConnected Then 'Checamos si la conexion es activa o no
    Command1.Enabled = False
    Command2.Enabled = True
Else
    Command1.Enabled = True
    Command2.Enabled = False
End If
End Sub
Private Sub Timer2_Timer()
If Text6.Text = "" Then
Text6.Text = "?Variable="
Else
End If
End Sub

'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'Dim Datos$
'Winsock1.GetData Datos, vbString, bytesTotal 'Obtenemos los datos
'Text1.Text = Datos
'End Sub
 
'Si queres grabar la respuesta en un html podes reemplazar lo anterior a
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Datos$, Cabecera, Html$, Canal% 'Declaramos variables
On Error Resume Next
Winsock1.GetData Datos 'Obtenemos datos
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

