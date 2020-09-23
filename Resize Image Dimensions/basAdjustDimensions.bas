Attribute VB_Name = "basAdjustDimensions"
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// This sub resizes the image's width and height based on the
'/// destination's width and height while maintaining aspect ratio.
'/// The image's aspect ratio is defined as : Aspect Ratio = Image's Height / Image's Width
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// NOTE : No error-handling included
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'The variable names explain themselves...
Public Sub AdjustImageDimensions( _
                ByRef ImageWidth As Single, ByRef ImageHeight As Single, _
                ByVal DestWidth As Single, ByVal DestHeight As Single)

      Dim WidthRatio As Single, HeightRatio As Single
      
      If ImageWidth > DestWidth Then
            WidthRatio = (DestWidth / ImageWidth)
            
            ImageWidth = (ImageWidth * WidthRatio)
            ImageHeight = (ImageHeight * WidthRatio)
      End If
      
      If ImageHeight > DestHeight Then
            HeightRatio = (DestHeight / ImageHeight)
            
            ImageWidth = (ImageWidth * HeightRatio)
            ImageHeight = (ImageHeight * HeightRatio)
      End If
End Sub
