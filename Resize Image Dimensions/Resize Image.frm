VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Resize Image Dimensions"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   14235
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   465
      Left            =   6825
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   7
      Top             =   5400
      Width           =   465
   End
   Begin VB.VScrollBar vsbHeight 
      Height          =   5265
      LargeChange     =   150
      Left            =   6825
      Max             =   5000
      Min             =   500
      SmallChange     =   10
      TabIndex        =   4
      Top             =   150
      Value           =   500
      Width           =   465
   End
   Begin VB.HScrollBar hsbWidth 
      Height          =   465
      LargeChange     =   150
      Left            =   75
      Max             =   6600
      Min             =   500
      SmallChange     =   10
      TabIndex        =   3
      Top             =   5400
      Value           =   500
      Width           =   6765
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   7425
      Picture         =   "Resize Image.frx":0000
      ScaleHeight     =   7200
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   150
      Width           =   6675
   End
   Begin VB.CommandButton cmdResizeImage 
      Caption         =   "Resize Image and paint it"
      Default         =   -1  'True
      Height          =   615
      Left            =   75
      TabIndex        =   1
      Top             =   7950
      Width           =   7215
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5190
      Left            =   75
      ScaleHeight     =   5190
      ScaleWidth      =   6690
      TabIndex        =   0
      Top             =   150
      Width           =   6690
   End
   Begin VB.Label lblDestHeight 
      Caption         =   "Destination Height :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   9
      Top             =   6450
      Width           =   7215
   End
   Begin VB.Label lblDestWidth 
      Caption         =   "Destination Width :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   8
      Top             =   6000
      Width           =   7215
   End
   Begin VB.Label lblNewAspectRatio 
      Caption         =   "New Aspect Ratio :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   6
      Top             =   7350
      Width           =   7215
   End
   Begin VB.Label lblOldAspectRatio 
      Caption         =   "Original Aspect Ratio :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   6900
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// This program resizes the image's width and height based on the
'/// destination's width and height while maintaining aspect ratio.
'/// The image's aspect ratio is defined as : Aspect Ratio = Image's Height / Image's Width
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// NOTE : No error-handling included
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdResizeImage_Click()
      Dim ImageWidth As Single      'Original image's width
      Dim ImageHeight As Single     'Original image's height
      Dim ResizedWidth As Single     'Resized image's width
      Dim ResizedHeight As Single    'Resized image's height
      Dim DestWidth As Single         'Destination picturebox's width
      Dim DestHeight As Single         'Destination picturebox's height
      Dim AspectRatio As Single        'Image's aspect ratio ( NOTE : aspect ratio = height / width )
      
      'Destination picturebox's dimensions
      DestWidth = picBox.Width
      DestHeight = picBox.Height
      
      'Stores the image's original dimensions
      ImageWidth = picImage.Width
      ImageHeight = picImage.Height
      
      'Initializes the resized dimensions
      ResizedWidth = ImageWidth
      ResizedHeight = ImageHeight
                  
      'Calculate image's original aspect ratio and display it in lblOldAspectRatio
      AspectRatio = (ImageHeight / ImageWidth)
      lblOldAspectRatio = "Original Aspect Ratio : " & AspectRatio
      
      'Now resize the dimensions...
      Call AdjustImageDimensions(ResizedWidth, ResizedHeight, DestWidth, DestHeight)
      
      'Calculate image's new aspect ratio and display it in lblNewAspectRatio
      AspectRatio = (ResizedHeight / ResizedWidth)
      lblNewAspectRatio = "New Aspect Ratio : " & AspectRatio
      
      'Paint the image onto picBox
      picBox.Cls
      picBox.PaintPicture picImage.Picture, 0, 0, ResizedWidth, ResizedHeight
                                
End Sub

Private Sub Form_Load()
      hsbWidth.Value = hsbWidth.Max
      vsbHeight.Value = vsbHeight.Max
End Sub

Private Sub vsbHeight_Change()
      picBox.Height = vsbHeight.Value
      lblDestHeight = "Destination Height : " & picBox.Height
End Sub

Private Sub hsbWidth_Change()
      picBox.Width = hsbWidth.Value
      lblDestWidth = "Destination Width : " & picBox.Width
End Sub
