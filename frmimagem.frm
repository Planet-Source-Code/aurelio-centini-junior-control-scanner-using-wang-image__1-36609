VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{009541A3-3B81-101C-92F3-040224009C02}#1.0#0"; "IMGADMIN.OCX"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#1.0#0"; "IMGEDIT.OCX"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "imgscan.ocx"
Object = "{E1A6B8A3-3603-101C-AC6E-040224009C02}#1.0#0"; "imgthumb.ocx"
Begin VB.Form frmimagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wang image control example Vote me if you like"
   ClientHeight    =   6555
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   11505
   Icon            =   "frmimagem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   11505
   Begin ScanLibCtl.ImgScan Scan 
      Left            =   9930
      Top             =   1140
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   820
      _StockProps     =   0
      DestImageControl=   "ImgEdit1"
      Image           =   "c:\sistemas\imagem\img."
      PageOption      =   3
      CompressionType =   2
      CompressionInfo =   29
      MultiPage       =   -1  'True
      ScanTo          =   1
      ShowSetupBeforeScan=   0   'False
   End
   Begin VB.CommandButton btn_func 
      Caption         =   "Exclude"
      Height          =   795
      Index           =   7
      Left            =   3270
      Picture         =   "frmimagem.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Don´t Work yet"
      Top             =   60
      Visible         =   0   'False
      Width           =   705
   End
   Begin ThumbnailLibCtl.ImgThumbnail ImgThumbnail1 
      Height          =   5565
      Left            =   30
      TabIndex        =   17
      Top             =   900
      Width           =   1995
      _Version        =   65536
      _ExtentX        =   3519
      _ExtentY        =   9816
      _StockProps     =   97
      BackColor       =   12632256
      BorderStyle     =   1
      ThumbCaptionStyle=   3
   End
   Begin VB.CommandButton btn_func 
      Caption         =   "Novo"
      Height          =   795
      Index           =   6
      Left            =   3240
      Picture         =   "frmimagem.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   60
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btn_func 
      Caption         =   "Scan"
      Height          =   795
      Index           =   4
      Left            =   120
      Picture         =   "frmimagem.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Start scan"
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stamp"
      Height          =   795
      Left            =   7350
      Picture         =   "frmimagem.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Stamp warnings"
      Top             =   60
      Width           =   675
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Tools"
      Height          =   795
      Left            =   8040
      Picture         =   "frmimagem.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Tool´s Box"
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdcores 
      Caption         =   "Color"
      Height          =   795
      Left            =   6780
      Picture         =   "frmimagem.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Change the color of image"
      Top             =   60
      Width           =   555
   End
   Begin VB.CommandButton btn_func 
      Caption         =   "Zoom"
      Height          =   795
      Index           =   3
      Left            =   6255
      Picture         =   "frmimagem.frx":167E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Zoom "
      Top             =   60
      Width           =   525
   End
   Begin VB.CommandButton btn_func 
      Caption         =   "Invert"
      Height          =   795
      Index           =   2
      Left            =   5610
      Picture         =   "frmimagem.frx":1988
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Invert the image"
      Top             =   60
      Width           =   645
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   795
      Left            =   10710
      Picture         =   "frmimagem.frx":21BA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close the aplication"
      Top             =   60
      Width           =   705
   End
   Begin VB.CommandButton btn_func 
      Caption         =   "Left"
      Height          =   795
      Index           =   1
      Left            =   4815
      Picture         =   "frmimagem.frx":24C4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Rotate left"
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton btn_func 
      Caption         =   "Right"
      Height          =   795
      Index           =   0
      Left            =   4230
      Picture         =   "frmimagem.frx":27CE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Rotate right"
      Top             =   60
      Width           =   585
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Text            =   "50"
      ToolTipText     =   "Altere o ZOOM para mais ou menos (valores de 10 ate 6000)"
      Top             =   120
      Width           =   375
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   10215
      TabIndex        =   2
      ToolTipText     =   "Change the ZOOM factor  (from  10 to 6000)"
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   100
      BuddyControl    =   "Text1"
      BuddyDispid     =   196614
      OrigLeft        =   2940
      OrigTop         =   780
      OrigRight       =   3180
      OrigBottom      =   1155
      Max             =   200
      Min             =   2
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   9420
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   795
      Left            =   2640
      Picture         =   "frmimagem.frx":2AD8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Clique neste botao para imprimir a imagem na impressora padrao"
      Top             =   60
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   795
      Left            =   720
      Picture         =   "frmimagem.frx":2DE2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Load an image from Disk"
      Top             =   60
      Width           =   705
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   5565
      Left            =   2070
      TabIndex        =   7
      Top             =   900
      Width           =   9405
      _Version        =   65536
      _ExtentX        =   16589
      _ExtentY        =   9816
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      Zoom            =   50
      AnnotationLineStyle=   1
      AutoRefresh     =   -1  'True
      Begin VB.Frame fra_impressao 
         Caption         =   "Select images to print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   3975
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   6975
         Begin VB.CommandButton Command7 
            Height          =   795
            Left            =   6120
            Picture         =   "frmimagem.frx":30EC
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Voltar para a tela principal"
            Top             =   3120
            Width           =   705
         End
         Begin VB.CommandButton cmd_print 
            Height          =   795
            Left            =   5400
            Picture         =   "frmimagem.frx":33F6
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Print"
            Top             =   3120
            Width           =   645
         End
         Begin ThumbnailLibCtl.ImgThumbnail Thumb2 
            Height          =   2655
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   6615
            _Version        =   65536
            _ExtentX        =   11668
            _ExtentY        =   4683
            _StockProps     =   97
            BackColor       =   12632256
            BorderStyle     =   1
            ScrollDirection =   0
         End
      End
   End
   Begin VB.CommandButton btn_func 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   795
      Index           =   5
      Left            =   2040
      Picture         =   "frmimagem.frx":3700
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Save the image"
      Top             =   60
      Width           =   585
   End
   Begin AdminLibCtl.ImgAdmin admin1 
      Left            =   9330
      Top             =   420
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   767
      _StockProps     =   0
      Filter          =   "*.TIF|TIFF|*.BMP|Bitmap|"
      DialogTitle     =   "Controle de Imagens"
      DefaultExt      =   "TIF"
      InitDir         =   "c:\"
      PrintStartPage  =   0
      PrintEndPage    =   0
   End
   Begin VB.Label Label1 
      Caption         =   "ZOOM"
      Height          =   255
      Left            =   9300
      TabIndex        =   6
      Top             =   150
      Width           =   555
   End
   Begin VB.Menu mnuopt 
      Caption         =   "Opções"
      Visible         =   0   'False
      Begin VB.Menu mnuapaga 
         Caption         =   "&Apagar"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnucopiar 
         Caption         =   "C&opiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuRecortar 
         Caption         =   "&Recortar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnucolar 
         Caption         =   "&Colar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuzomm 
         Caption         =   "&Zomm"
      End
      Begin VB.Menu mnuconvert 
         Caption         =   "Converter Imagem"
         Begin VB.Menu mnuConvPB 
            Caption         =   "Preto e Branco"
         End
         Begin VB.Menu mnuConvCinza4 
            Caption         =   "Cinza (4 Escalas)"
         End
         Begin VB.Menu mnuConvCinza8 
            Caption         =   "Cinza ( 8 Escalas )"
         End
         Begin VB.Menu mnuConvCorSimples 
            Caption         =   "Colorido Simples"
         End
         Begin VB.Menu mnuconvCorMedio 
            Caption         =   "Colorido Médio"
         End
         Begin VB.Menu mnuconvCorComp 
            Caption         =   "Colorido Complexo"
         End
      End
   End
End
Attribute VB_Name = "frmimagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I hope it is usefull to you
'This program uses Wang image control to manipulate the image and Scanner
'if you have some dificulty email me centini@hpg.com.br
'Please vote if you like


Private Sub btn_func_Click(Index As Integer)
On Error GoTo Erro_Scaner
 Dim X As Integer
 Select Case Index
     Case 0
       ImgEdit1.RotateRight
     Case 1
       ImgEdit1.RotateLeft
     Case 2
       ImgEdit1.Flip
     Case 3
       ImgEdit1.ZoomToSelection
       
     Case 4
        If Dir(ImgEdit1.Image) <> "" Then
           'clear the image
           ImgEdit1.DisplayBlankImage 100, 100, 100, 100, BlackAndWhite
           Scan.ShowSetupBeforeScan = False
           Scan.OpenScanner
           Scan.ResetScanner
           'Set the image resolution
           ImgEdit1.ImageResolutionX = 150
           ImgEdit1.ImageResolutionY = 150
           ImgEdit1.AutoRefresh = True
           
           Scan.Image = ImgEdit1.Image
           
           'set the type of image
           Scan.PageType = BlackAndWhite
           'set the compression type
           Scan.CompressionType = CCITTGroup4_2d_Fax '  74 k per image
           Scan.CompressionInfo = 2 ' 73 k
           'Type of inclusion
           Scan.PageOption = AppendPages
     
           'Scan the image
           Scan.StartScan
           ' free the scanner
           Scan.ResetScanner
           Scan.CloseScanner
           'Dysplay image
           ImgEdit1.Display
           ImgEdit1.Save
           ImgEdit1.Refresh
           'update the thumbnail
           ImgThumbnail1.Image = ImgEdit1.Image
           ImgEdit1.Page = ImgEdit1.PageCount
           ImgEdit1.Display
           
           'Convert page type to lower resolution
           ' You can modify this to higher resolution
           ImgEdit1.ConvertPageType 1, True
           ImgEdit1.ImageResolutionX = 150
           ImgEdit1.ImageResolutionY = 150
           ImgEdit1.AutoRefresh = True
           ImgThumbnail1.Image = ImgEdit1.Image
           ImgEdit1.Save
        Else
           'Case is the first page to be scanned
           Scan.PageOption = CreateNewFile
           Scan.ShowScanNew False
           Me.ZOrder 10
        End If
        ImgEdit1.AutoRefresh = True
     Case 5
        ' save the image
        admin1.ShowFileDialog SaveDlg
        ' Clear the thumbnails
        ImgThumbnail1.ClearThumbs
        
        ImgThumbnail1.Refresh
        
        ImgEdit1.Display
        ImgEdit1.Save
        ImgThumbnail1.Refresh
        
     Case 6
        '
        ImgEdit1.DisplayBlankImage 100, 100, 200, 200, BlackAndWhite
        Scan.ShowSetupBeforeScan = True
        Scan.ShowScanNew
        Scan.ShowSetupBeforeScan = False
        'transfer the image to Imgedit
        ImgEdit1.Image = Scan.Image
        'update o thumbnail
        ImgThumbnail1.Image = ImgEdit1.Image
        ImgThumbnail1.Refresh
        ImgThumbnail1.DisplayThumbs
        ' close the scanner
        Scan.ResetScanner
        Scan.CloseScanner
        
     Case 7
        'ImgEdit1.DeleteImageData
 End Select
 Exit Sub

Erro_Scaner:
    MsgBox "Erro " & Err.Number & Chr(10) & Err.Description
    Resume Next
End Sub


Private Sub cmd_print_Click()
    'to print selected image
    Dim icnt As Integer
    Dim rx As Long
    Dim ry As Long
    For icnt = 1 To Thumb2.ThumbCount
         If Thumb2.ThumbSelected(icnt) = True Then
            rx = ImgEdit1.ImageResolutionX
            ry = ImgEdit1.ImageResolutionY
            ImgEdit1.ImageResolutionX = 400
            ImgEdit1.ImageResolutionY = 400
            '
            ImgEdit1.PrintImage icnt
            ImgEdit1.ImageResolutionX = rx
            ImgEdit1.ImageResolutionY = ry
          
         End If
    Next icnt

End Sub

Private Sub cmdcores_Click()
   ' define the colors of annotation tool
   dlg1.ShowColor
   ImgEdit1.AnnotationLineColor = dlg1.Color
   ImgEdit1.AnnotationFontColor = dlg1.Color
   ImgEdit1.Refresh
End Sub

Private Sub Command1_Click()
  'load a pre scanned image
  On Error GoTo erro_carga
  dlg1.DialogTitle = "Select the Image file"
  dlg1.FileName = ""
  dlg1.Filter = "*.tif; *.bmp"
  dlg1.Action = 1
  If dlg1.FileName = "" Then Exit Sub
  ImgEdit1.Image = dlg1.FileName
  ImgEdit1.Display
  ImgThumbnail1.Image = dlg1.FileName
  ImgThumbnail1.DisplayThumbs
  Scan.Image = ImgThumbnail1.Image
  Exit Sub
erro_carga:
  MsgBox Error$, vbCritical, "Image Control"

End Sub

Private Sub Command2_Click()
   'call the print form
    Thumb2.Image = ImgEdit1.Image
    fra_impressao.Visible = True
   
End Sub

Private Sub Command3_Click()
   'show the annotation tools pallete
   ImgEdit1.ShowAnnotationToolPalette
   
End Sub

Private Sub Command4_Click()
   'Show the stamp dialog
   ImgEdit1.ShowRubberStampDialog
End Sub


Private Sub Command5_Click()
End
End Sub


Private Sub Command7_Click()
 fra_impressao.Visible = False
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

End Sub

Private Sub ImgEdit1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = 2 Then
      PopupMenu mnuopt
   End If
End Sub

Private Sub ImgThumbnail1_Click(ByVal ThumbNumber As Long)
   'Display a selected image
   If ThumbNumber > 0 Then
      ImgEdit1.Image = ImgThumbnail1.Image
      ImgEdit1.Page = ThumbNumber
      ImgEdit1.Display
   End If
End Sub

Private Sub mnuapaga_Click()
' Delete selected area
  ImgEdit1.DeleteSelectedAnnotations
End Sub

Private Sub mnucolar_Click()
   'Paste the image
   ImgEdit1.ClipboardPaste
End Sub

'***************************************************************
'  Convertion Functions
'***************************************************************
'PageType Settings

'The following list shows the valid PageType settings:

'Setting Description
'1   Black-and-white
'2   Gray4
'3   Gray8
'4   Palettized4 (not available as a parameter for the ConvertPageType method unless the image presently has a 4-bit palette)
'5   Palettized8
'6   RGB24
'7   BGR24

'Copyright Wang Laboratories, Inc. 1995-1996


Private Sub mnuConvCinza4_Click()
  ' convert image to Gray Scale 4 tons
  ImgEdit1.ConvertPageType 2
End Sub

Private Sub mnuConvCinza8_Click()
  
  ' convert image to Gray Scale 8 tons
  ImgEdit1.ConvertPageType 3

End Sub

Private Sub mnuconvCorComp_Click()

  ' convert image to Gray Scale 8 tons
   ImgEdit1.ConvertPageType 7
End Sub

Private Sub mnuconvCorMedio_Click()
ImgEdit1.ConvertPageType 6
End Sub

Private Sub mnuConvCorSimples_Click()
ImgEdit1.ConvertPageType 5
End Sub

Private Sub mnuConvPB_Click()
'convert image
 ImgEdit1.ConvertPageType 1
 ImgEdit1.Save
End Sub

Private Sub mnucopiar_Click()

'object.ClipboardCopy [Left,Top,Width,Height]

'Arguments

'Parameter   Data Type   Setting
'Left    Long    The upper-left corner of the selection rectangle in pixel coordinates of the displayed image
'Top Long    The top of the selection rectangle in pixel coordinates of the displayed image
'Width   Long    The width of the selection rectangle in pixels
'Height  Long    The height of the selection rectangle in pixels


'Copyright Wang Laboratories, Inc. 1995-1996
   
   
   On Error Resume Next
   ImgEdit1.ClipboardCopy
End Sub

Private Sub mnuRecortar_Click()
   On Error Resume Next
   ImgEdit1.ClipboardCut
End Sub

Private Sub mnuzomm_Click()
   On Error Resume Next
   'Applies Zoom to selected Area
   ImgEdit1.ZoomToSelection
End Sub

Private Sub Text1_Change()
  'apply zoom factor to the image
   On Error GoTo teste_erro
   ImgEdit1.Zoom = Text1.Text
   ImgEdit1.Refresh
   Exit Sub
teste_erro:
   MsgBox Error$, vbCritical, Caption
End Sub

Private Sub Thumb2_Click(ByVal ThumbNumber As Long)
 'load a selected image
    If Thumb2.ThumbSelected(ThumbNumber) = True Then
       Thumb2.ThumbSelected(ThumbNumber) = False
    Else
       Thumb2.ThumbSelected(ThumbNumber) = True
    End If
End Sub
