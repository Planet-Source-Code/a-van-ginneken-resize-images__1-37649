VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmResize 
   Caption         =   "Resizen"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form2"
   ScaleHeight     =   5415
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileSize 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog codi 
      Left            =   4080
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "Open file"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdProces 
      Caption         =   "Resize"
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   720
      ScaleHeight     =   1395
      ScaleWidth      =   3075
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CheckBox chkStretch 
      Caption         =   "Orginal"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtImageHeight 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      MaxLength       =   4
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtImageWidth 
      Height          =   285
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblFileSize 
      Caption         =   "File size"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblHeight 
      Caption         =   "Height"
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblWidth 
      Caption         =   "Width"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   720
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2895
   End
End
Attribute VB_Name = "frmResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public orgHeight As Long
Public orgWidth As Long
Public sFileName As String
Public sSecFileName As String
Private Declare Function DIWriteJpg Lib "DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long

Private Sub chkStretch_Click()

    If chkStretch.Value Then
        txtImageWidth.Text = orgWidth
        txtImageHeight.Text = orgHeight
        Image1.Stretch = False
    Else
        Image1.Width = txtImageWidth.Text * 15
        Image1.Height = txtImageHeight.Text * 15
        Image1.Stretch = True
    End If

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdOpenFile_Click()
    
    codi.DialogTitle = "Select a photo or a file"
    codi.InitDir = "C:\"
    codi.ShowOpen
  
    sSecFileName = codi.FileName
    sFileName = codi.FileTitle
    
    LoadImage
    
End Sub

Private Sub cmdProces_Click()

    Picture1.Picture = CaptureClient(frmResize, Image1.Left / 15, Image1.Top / 15, Image1.Width / 15, Image1.Height / 15)
    Picture1.AutoSize = True
    
    loadstr = "c:\resized" & sFileName
    
    'Required by DIjpg.dll
    SavePicture Picture1.Image, "C:\tmp.bmp"
    'Save to JPEG
    retval = DIWriteJpg(loadstr, 100, 0)
    
    If retval = 1 Then  'Success
        frmView.Image1.Picture = LoadPicture(loadstr)
        frmView.lblFileName = "File: " & loadstr
       
        
        frmView.txtFileSize.Text = Format(FileLen(loadstr) / 1024, "##.##") & " Kb"
        frmView.Show vbModal, Me
        'frmjpg.SetFileName (loadStr)
    End If
    Kill "c:\tmp.bmp"
    
End Sub

Private Function LoadImage()
Dim x As Single
Dim y As Single
Dim ii As New CImageInfo
Dim msg As String
Dim sFileName As String

    Image1.Stretch = False
    frmResize.WindowState = vbMaximized
 
    Set frmResize.Image1.Picture = LoadPicture(sSecFileName)
    
    If sSecFileName <> "" Then
        
        Picture1.Picture = LoadPicture(sSecFileName)
        
        ii.ReadImageInfo (sSecFileName)
        msg = msg & "FileName: " & sFileName & vbCrLf
        If ii.ImageType Then
            txtImageWidth.Text = ii.Width
            txtImageHeight.Text = ii.Height
            orgWidth = ii.Width
            orgHeight = ii.Height
    
        Else
        'Image not reconized
            txtImageWidth.Text = 100
            txtImageHeight.Text = 100
            orgWidth = 100
            orgHeight = 100
        End If
       
        Image1.Width = txtImageWidth * 15
        Image1.Height = txtImageHeight * 15
        
        Image1.Stretch = True
        Picture1.Top = Image1.Top + Image1.Height + 3000
        txtFileSize.Text = Format(FileLen(sSecFileName) / 1024, "##.##") & " Kb"
    End If

End Function

Private Sub Form_Load()
   Me.WindowState = vbMaximized
End Sub

Private Sub Form_Unload(Cancel As Integer)
    orgHeight = 0
    orgWidth = 0
    Unload Me
End Sub

Private Sub slVer_Click()
  Image1.Height = ((orgHeight / slVer.Value * 100))
  txtImageHeight.Text = Image1.Height / 15
  Image1.Stretch = True
End Sub

Private Sub slSize_Change()
  Image1.Height = orgHeight / (slSize + 1) * 100
  Image1.Width = orgWidth / (slSize + 1) * 100
  txtImageHeight.Text = Image1.Height / 15
  txtImageWidth.Text = Image1.Width / 15
End Sub

Private Sub txtImageHeight_Change()
    If txtImageHeight.Text <> "" Then
      Image1.Height = txtImageHeight.Text * 15
    End If
End Sub

Private Sub txtImageHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    If txtImageHeight.Text <> "" Then
      Image1.Height = txtImageHeight.Text * 15
    End If
End Sub

Private Sub txtImageWidth_KeyUp(KeyCode As Integer, Shift As Integer)
Dim a As Single

    If txtImageWidth.Text <> "" Then
        chkStretch.Value = 0
        a = txtImageWidth.Text / orgHeight
        a = orgWidth / txtImageWidth.Text
        txtImageHeight.Text = Format(orgHeight / a, "##.##")
        If Right(txtImageHeight.Text, 1) = "," Then
          txtImageHeight.Text = Mid(txtImageHeight.Text, 1, Len(txtImageHeight.Text) - 1)
        End If
        Image1.Width = txtImageWidth.Text * 15
        Image1.Height = txtImageHeight.Text * 15
    End If
    
End Sub
