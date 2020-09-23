VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mike Idea & Santiago Coder Compresion"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   1710
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   603
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   886
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Cargar"
      Height          =   2415
      Left            =   6840
      TabIndex        =   11
      Top             =   6600
      Width           =   1815
      Begin VB.CheckBox chkFreePixels 
         Caption         =   "Free Pixels"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkClones 
         Caption         =   "Clones"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkMasters 
         Caption         =   "Masters"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Cargar MSC"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Load a msc file"
         Top             =   1845
         Width           =   1335
      End
   End
   Begin VB.Frame framSave 
      Caption         =   "Grabar"
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   6600
      Width           =   5175
      Begin VB.CheckBox chkProgress 
         Caption         =   "Mostrar progreso (Mas lento)"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         ToolTipText     =   "Show progress. It will be slower"
         Top             =   1920
         Width           =   2535
      End
      Begin MSComctlLib.Slider sldIgualdad 
         Height          =   495
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Percentaje of similar blocks"
         Top             =   1200
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   95
         TickFrequency   =   10
         Value           =   95
         TextPosition    =   1
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Grabar como MSC"
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         ToolTipText     =   "Save the file"
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Similitud de Bloque (Block Similar percentage)):"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label lblDest 
         Caption         =   "C:\A.MSC"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         ToolTipText     =   "The destination file path. Click to select where you want save"
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label lblLabel5 
         Caption         =   "Destino:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "The destination file path. Click to select where you want save"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblImg 
         Caption         =   "C:\"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   840
         TabIndex        =   5
         ToolTipText     =   "Original image path. Click to open"
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Origen:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Original image path. Click to open"
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   5760
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMSC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2835
      Left            =   6840
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox picOriginal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2835
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.Shape shpCloneLast 
         BorderColor     =   &H0000C000&
         DrawMode        =   10  'Mask Pen
         Height          =   120
         Left            =   240
         Top             =   0
         Width           =   120
      End
      Begin VB.Shape shpClone 
         BorderColor     =   &H00FF0000&
         DrawMode        =   10  'Mask Pen
         Height          =   120
         Left            =   120
         Top             =   0
         Width           =   120
      End
      Begin VB.Shape shpMaster 
         BorderColor     =   &H000000FF&
         DrawMode        =   7  'Invert
         Height          =   120
         Left            =   0
         Top             =   0
         Width           =   120
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
On Error GoTo Solucion
    ' Muestro el cuadro de diálogo de origen
    ' Show the open dialog
    With Cd1
        .CancelError = True
        .Filter = "Mike & Santiago Compresion |*.msc"
        .ShowOpen
    End With
    
    ' Load the file
    LoadMSC Cd1.FileName, picMSC, CBool(chkMasters), CBool(chkClones), CBool(chkFreePixels)
   
    Exit Sub
Solucion:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox "Error " & Err.Number & " lblimg_click"
    End If
End Sub

Private Sub cmdSave_Click()
    If MsgBox("This is a very slow algorithm. Do you wanna continue?" & vbCrLf & " I took me around 6 minutes on a 300 x 350 image on a Celeron 400 Mhz", vbInformation + vbYesNo) = vbYes Then
        ' Save the file
        SaveAsMSC picOriginal, lblDest, sldIgualdad.Value, CBool(chkProgress)
    End If
End Sub

Private Sub Form_Load()
    MsgBox "It is a representation and test of what Mike said" & vbCrLf & _
           "It's a very slow algorythm because is only for test that he said" & vbCrLf & _
           "It will find the pieces of image that looks similar and put them" & vbCrLf & _
           "into a master block in the file compresed to make it less size" & vbCrLf & _
           vbCrLf & _
           "The result will be comparated to a BMP image without compression" & vbCrLf & _
            vbCrLf & _
           "Compile if you wanna do it 3 or 4 times more faster" & vbCrLf & _
           "Enjoy and try to make it better" & vbCrLf & _
           "Thanks, Santiago F.", vbInformation
End Sub

Private Sub lblDest_Click()
On Error GoTo Solucion
    ' Muestro el cuadro de diálogo de origen
    ' Show the save dialog
    With Cd1
        .CancelError = True
        .Filter = "Mike & Santiago Compresion |*.msc"
        .ShowSave
    End With
    
    lblDest = Cd1.FileName
    
    Exit Sub
Solucion:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox "Error " & Err.Number & " lblimg_click"
    End If
End Sub

Private Sub lblImg_Click()
On Error GoTo Solucion
    ' Muestro el cuadro de diálogo de origen
    ' Show the open dialog
    With Cd1
        .CancelError = True
        .Filter = "Archivos de imágen|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.dib;*.wmf;*.emf"
        .ShowOpen
    End With
    
    lblImg = Cd1.FileName
    picOriginal.Picture = LoadPicture(Cd1.FileName)
    
    Exit Sub
Solucion:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox "Error " & Err.Number & " lblimg_click"
    End If
End Sub
