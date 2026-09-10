VERSION 5.00
Begin VB.UserControl AxConGrid 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   PropertyPages   =   "AxConGrid.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "AxConGrid.ctx":0011
   Begin AxConectionGrid.ucScrollbar ucScrollV 
      Height          =   3375
      Left            =   4455
      Top             =   135
      Width           =   165
      _extentx        =   291
      _extenty        =   5953
      style           =   4
      showbuttons     =   0
      thumbtooltipfont=   "AxConGrid.ctx":0323
   End
   Begin AxConectionGrid.ucScrollbar ucScrollH 
      Height          =   165
      Left            =   75
      Top             =   3225
      Width           =   4215
      _extentx        =   7435
      _extenty        =   291
      orientation     =   1
      style           =   4
      showbuttons     =   0
      thumbtooltipfont=   "AxConGrid.ctx":034B
   End
End
Attribute VB_Name = "AxConGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxRelationsGrid
'Version  : 0.03.01a
'Editor   : David Rojas [AxioUK]
'Date     : 2026-09-10
'------------------------------------
Option Explicit

Const sVersion = "0.03.01a"

Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function SetRECL Lib "user32" Alias "SetRect" (lpRect As RECTL, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

' GDI BackBuffer APIs
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' GDI+ APIs
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipGraphicsClear Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mColor As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal Brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal RGBA As Long, ByRef Brush As Long) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTS, ByVal mFormat As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As eStringTrimming) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As eStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As eStringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipCreateLineBrush Lib "gdiplus" (point1 As POINTS, point2 As POINTS, ByVal color1 As Long, ByVal Color2 As Long, ByVal WrapMd As Long, lineGradient As Long) As Long
Private Declare Function GdipSetPenBrushFill Lib "gdiplus" (ByVal pen As Long, ByVal Brush As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As DashStyle) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
Private Declare Function GdipDrawBezier Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal X3 As Single, ByVal Y3 As Single, ByVal X4 As Single, ByVal Y4 As Single) As Long
Private Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal pen As Long, ByVal StartCap As LineCap) As Long
Private Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal pen As Long, ByVal EndCap As LineCap) As Long

Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
'---

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type RECTS
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type POINTS
   X As Single
   Y As Single
End Type

Private Type POINTL
    X As Long
    Y As Long
End Type

Private Enum CallOutPosition
  coLeft
  coTop
  coRight
  coBottom
End Enum

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PicBmp
  Size As Long
  type As Long
  hBmp As Long
  hpal As Long
  Reserved As Long
End Type

Public Enum eStringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum
  
Public Enum eStringTrimming
    StringTrimmingNone = &H0
    StringTrimmingCharacter = &H1
    StringTrimmingWord = &H2
    StringTrimmingEllipsisCharacter = &H3
    StringTrimmingEllipsisWord = &H4
    StringTrimmingEllipsisPath = &H5
End Enum

Public Enum eStringFormatFlags
    StringFormatFlagsNone = &H0
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum

Public Enum eTextAlignH
    eLeft
    eCenter
    eRight
End Enum

Public Enum eTextAlignV
    eTop
    eMiddle
    eBottom
End Enum

Public Enum eVisibleType
  scNone
  scAllways
  scOnlyActive
End Enum

Public Enum eTypeBadge
  vbNone
  vbLabel
  vbIcon
End Enum

Public Enum eImageType
  eIconFont
  eImageFile
End Enum

Private Type tCell
    Text As String
    SubText As String
    Label As String
    Tag As Variant
    BackColor As Long
    ForeColor As Long
    Font As StdFont
    Height As Single
    PointLX As Single
    PointRX As Single
    PointY As Single
    Radius As Single
    LineTo As Long
    ColumnTo As Long
    CurveTo As Long
    SideTo As tSide
    Direction As tDirection
    IconChar As Long
    IconFile As String
    CellVisible As Boolean
    LineWidth As Single
    LineStyle As DashStyle
    LineStartCap As LineCap
    LineEndCap As LineCap
    LineStartColor As OLE_COLOR
    LineEndColor As OLE_COLOR
    LineOpacity As Long
    LineVisible As Boolean
End Type

Private Type tColumn
    Header As String
    Cell() As tCell
    StartX As Long
    StartY As Long
    Width As Long
    BackColor As OLE_COLOR
    ForeColor As OLE_COLOR
    Font As StdFont
    Visible As Boolean
End Type

Public Enum tDirection
  UpSide = 0
  DownSide = 1
End Enum

Public Enum tSide
  toLeft = 0
  toRight = 1
  Auto = 2
End Enum

Public Enum LineCap
   LineCapFlat = 0
   LineCapSquare = 1
   LineCapRound = 2
   LineCapTriangle = 3
   LineCapNoAnchor = &H10         ' corresponds to flat cap
   LineCapSquareAnchor = &H11     ' corresponds to square cap
   LineCapRoundAnchor = &H12      ' corresponds to round cap
   LineCapDiamondAnchor = &H13    ' corresponds to triangle cap
   LineCapArrowAnchor = &H14      ' no correspondence
End Enum

Public Enum DashStyle
   DashStyleSolid          ' 0
   DashStyleDash           ' 1
   DashStyleDot            ' 2
   DashStyleDashDot        ' 3
   DashStyleDashDotDot     ' 4
End Enum

'Constants
Private Const CombineModeExclude As Long = &H4
Private Const WrapModeTileFlipXY = &H3
Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const TLS_MINIMUM_AVAILABLE As Long = 64
Private Const IDC_HAND As Long = 32649
Private Const UnitPixel As Long = &H2&

'Define EVENTS-------------------
Public Event Click(lColumn As Long, lItem As Long, iColumnTo As Long, iLineTo As Long, iCurveTo As Long)
Public Event DblClick(lColumn As Long, lItem As Long)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Property Variables:
Private hFontCollection As Long
Private GdipToken   As Long
Private nScale      As Single
Private hCur        As Long
Private m_ImageType As eImageType

Private m_Clickable     As Boolean
Private m_Editable      As Boolean
Private m_Enabled       As Boolean

Private m_Col()         As tColumn
Private mCol            As Long
Private m_ColWidth      As Long
Private m_ActiveCol     As Long
Private m_ItemHeight    As Long
Private m_ActiveItem()  As Long
Private m_SelectItem()  As Long

Private m_BorderColor   As OLE_COLOR
Private m_BackColor     As OLE_COLOR
Private m_BoxColor      As OLE_COLOR
Private m_BorderWidth   As Long
Private m_CornerCurve   As Long
Private m_ForeColor1    As OLE_COLOR
Private m_ForeColor2    As OLE_COLOR
Private m_Font1         As StdFont
Private m_Font2         As StdFont
Private m_IconFont      As StdFont
Private m_LabelFont     As StdFont
Private m_HeaderFont    As StdFont
Private m_HeaderBackColor As OLE_COLOR
Private m_HeaderForeColor As OLE_COLOR
Private m_ColorActive   As OLE_COLOR
Private m_BorderColorActive As OLE_COLOR

Private m_CaptionAlignV As eTextAlignV
Private m_CaptionAlignH As eTextAlignH
Private m_SubTextAlignV As eTextAlignV
Private m_SubTextAlignH As eTextAlignH
Private m_FontOpacity   As Long

Private m_IconForeColor As OLE_COLOR
Private m_IconAlignV    As eTextAlignV
Private m_IconAlignH    As eTextAlignH
Private m_Bmp           As Long

Private m_SubTextVisible As eVisibleType
Private m_BadgeVisible   As eVisibleType
Private m_BadgeType      As eTypeBadge
Private m_BoxOpacity     As Long

Private m_LineWidth     As Single
Private m_LineStyle     As DashStyle
Private m_LineOpacity   As Long
Private m_LineStartCap  As LineCap
Private m_LineEndCap    As LineCap
Private m_LineStartColor As OLE_COLOR
Private m_LineEndColor  As OLE_COLOR

Private SideStart   As Long
Private SideEnd     As Long
Private InitCurve   As Long
Private EndCurve    As Long
Private mY          As Single
Private mX          As Single
Private m_X1        As Single
Private m_Y1        As Single
Private m_X2        As Single
Private m_Y2        As Single

Private m_DefaultSide  As tSide
Private m_ColsMoveable As Boolean
Private mRedraw        As Boolean

' Double Buffering Variables
Private m_hMemDC        As Long
Private m_hMemBmp       As Long
Private m_hOldBmp       As Long
Private m_MemWidth      As Long
Private m_MemHeight     As Long

' Drag and Drop variables
Private m_DragOffsetX   As Single
Private m_DragOffsetY   As Single
Private m_IsDragging    As Boolean


Public Function AddColumn(Optional sCaption As String, Optional lWidth As Long, _
                          Optional lX As Long = 0, Optional lY As Long = 0, _
                          Optional oBackColor As OLE_COLOR, Optional oForeColor As OLE_COLOR, _
                          Optional bVisible As Boolean = True) As Boolean
On Error Resume Next
Dim C As Long

C = ColCount

ReDim Preserve m_Col(C)
ReDim Preserve m_ActiveItem(C)
ReDim Preserve m_SelectItem(C)

With m_Col(C)
  .Header = sCaption
  .BackColor = IIf(oBackColor = 0, m_HeaderBackColor, oBackColor)
  .ForeColor = IIf(oForeColor = 0, m_HeaderForeColor, oForeColor)
  .Width = IIf(lWidth > 0, lWidth, m_ColWidth)
  .StartX = lX
  .StartY = lY
  .Visible = bVisible
End With

m_ActiveItem(C) = -1
m_SelectItem(C) = -1

Refresh
End Function

Public Function AddItem(ByVal lColumn As Long, ByVal eText As String, Optional eSubText As String = "", Optional eLabel As String = "", _
                        Optional eBackColor As OLE_COLOR = vbWhite, Optional eForeColor As OLE_COLOR = vbBlack, _
                        Optional lHeight As Long, Optional eTag As Variant = "", _
                        Optional eIconchar As String = "&H0", Optional eIconFile As String = "", _
                        Optional eVisible As Boolean = True) As Boolean
                        
On Error Resume Next
Dim i   As Long
  
  i = ItemCount(lColumn)
    
With m_Col(lColumn)
  ReDim Preserve .Cell(i)
  .Cell(i).Text = eText
  .Cell(i).SubText = eSubText
  .Cell(i).Label = eLabel
  .Cell(i).CellVisible = eVisible
  .Cell(i).BackColor = eBackColor
  .Cell(i).ForeColor = eForeColor
  .Cell(i).Height = IIf(lHeight > 0, lHeight, m_ItemHeight)
  .Cell(i).Tag = eTag
  .Cell(i).LineTo = -1
  .Cell(i).CurveTo = -1
  .Cell(i).LineVisible = False
  .Cell(i).IconChar = IconCharCode(eIconchar)
  .Cell(i).IconFile = eIconFile
End With

m_ActiveItem(lColumn) = -1
m_SelectItem(lColumn) = -1

If eVisible Then Refresh
End Function

Public Function AddLine(ByVal lColStart As Long, ByVal ItemStart As Long, ByVal lColEnd As Long, ByVal ItemEnd As Long, Optional lStyle As DashStyle = 0, _
                        Optional lStartCap As LineCap = 0, Optional lEndCap As LineCap = 0, _
                        Optional cStartColor As OLE_COLOR = &HFF&, Optional cEndColor As OLE_COLOR = &HC67300, _
                        Optional lLineWidth As Single = 5, Optional lOpacity As Long = 50, _
                        Optional bVisible As Boolean = True)
On Error Resume Next
Dim Side As tSide

If lColStart > lColEnd Then
    Side = toLeft
Else
    Side = toRight
End If

If m_Col(lColEnd).Cell(ItemEnd).LineTo = ItemStart And m_Col(lColEnd).Cell(ItemEnd).ColumnTo = lColStart Then
  With m_Col(lColEnd).Cell(ItemEnd)
    .LineStartCap = .LineEndCap
    .LineStartColor = .LineEndColor
  End With
Else
  With m_Col(lColStart).Cell(ItemStart)
    .ColumnTo = lColEnd
    .LineTo = ItemEnd
    .SideTo = Side
    .LineStyle = lStyle
    .LineWidth = lLineWidth
    .LineStartCap = lStartCap
    .LineEndCap = lEndCap
    .LineStartColor = cStartColor
    .LineEndColor = cEndColor
    .LineOpacity = lOpacity
    .LineVisible = bVisible
  End With
End If

If bVisible Then Refresh
End Function

Public Function AddCurve(ByVal lColumn As Long, ByVal ItemStart As Long, ByVal ItemEnd As Long, Optional lSide As tSide, _
                            Optional iRadius As Single = 50, Optional lStyle As DashStyle = 0, _
                            Optional lStartCap As LineCap = 0, Optional lEndCap As LineCap = 0, _
                            Optional cStartColor As OLE_COLOR = &HFF&, Optional cEndColor As OLE_COLOR = &HC67300, _
                            Optional lLineWidth As Single = 5, Optional lOpacity As Long = 50, _
                            Optional bVisible As Boolean = True)
On Error Resume Next

If m_Col(lColumn).Cell(ItemEnd).CurveTo = ItemStart Then
  With m_Col(lColumn).Cell(ItemEnd)
    .LineStartCap = .LineEndCap
    .LineEndColor = .LineStartColor
    .LineVisible = bVisible
  End With
Else
  With m_Col(lColumn).Cell(ItemStart)
    .CurveTo = ItemEnd
    .Direction = IIf(ItemStart > ItemEnd, 0, 1)
    .SideTo = lSide
    .Radius = iRadius
    .LineStyle = lStyle
    .LineWidth = lLineWidth
    .LineStartCap = lStartCap
    .LineEndCap = lEndCap
    .LineStartColor = cStartColor
    .LineEndColor = cEndColor
    .LineOpacity = lOpacity
    .LineVisible = bVisible
  End With
End If

If bVisible Then Refresh
End Function

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Public Sub Clear()
    Erase m_Col
    Refresh
End Sub

Public Sub GetCanvasBounds(ByRef MaxX As Long, ByRef MaxY As Long)
    Dim C As Long, bR As Long, bB As Long
    MaxX = 0
    MaxY = 0
    If ColCount <= 0 Then Exit Sub
    
    For C = 0 To ColCount - 1
        If m_Col(C).Visible Then
            bR = m_Col(C).StartX + m_Col(C).Width + 30
            bB = m_Col(C).StartY + (ItemCount(C) + 1) * m_ItemHeight + 40
            If bR > MaxX Then MaxX = bR
            If bB > MaxY Then MaxY = bB
        End If
    Next C
End Sub

Private Sub UpdateScrollBars()
    Dim MaxX As Long, MaxY As Long
    Dim vMax As Long, hMax As Long
    
    If ColCount <= 0 Then
        ucScrollV.Visible = False
        ucScrollH.Visible = False
        Exit Sub
    End If
    
    GetCanvasBounds MaxX, MaxY
    
    If MaxY > UserControl.ScaleHeight Then
        vMax = MaxY - UserControl.ScaleHeight + 30
        With ucScrollV
            .Max = vMax
            If .Value > vMax Then .Value = vMax
            .Visible = True
            .TrackMouseWheelOnHwnd UserControl.hwnd
        End With
    Else
        With ucScrollV
            .Max = 0
            .Value = 0
            .Visible = False
        End With
    End If
    
    If MaxX > UserControl.ScaleWidth Then
        hMax = MaxX - UserControl.ScaleWidth + 30
        With ucScrollH
            .Max = hMax
            If .Value > hMax Then .Value = hMax
            .Visible = True
            .TrackMouseWheelOnHwnd UserControl.hwnd
        End With
    Else
        With ucScrollH
            .Max = 0
            .Value = 0
            .Visible = False
        End With
    End If
End Sub

Public Sub Refresh()
  On Error Resume Next
  If mRedraw = False Then Exit Sub
  UpdateScrollBars
  Draw
End Sub

Public Sub RemoveItem(ByVal lColumn As Long, ByVal Item As Long)
  On Local Error Resume Next
  Dim j As Long, C As Long
  
  If lColumn < 0 Or lColumn >= ColCount Then Exit Sub
  If Item < 0 Or Item >= ItemCount(lColumn) Then Exit Sub
  
  With m_Col(lColumn)
    If ItemCount(lColumn) > 1 Then
      For j = Item To UBound(.Cell) - 1
        .Cell(j) = .Cell(j + 1)
      Next
      ReDim Preserve .Cell(UBound(.Cell) - 1)
    Else
      Erase .Cell
    End If
  End With
  
  For C = 0 To ColCount - 1
    If C <> lColumn Then
      With m_Col(C)
        For j = 0 To ItemCount(C) - 1
          If .Cell(j).ColumnTo = lColumn Then
            If .Cell(j).LineTo = Item Then
              .Cell(j).LineTo = -1
            ElseIf .Cell(j).LineTo > Item Then
              .Cell(j).LineTo = .Cell(j).LineTo - 1
            End If
          End If
        Next j
      End With
    Else
      With m_Col(C)
        For j = 0 To ItemCount(C) - 1
          If .Cell(j).CurveTo = Item Then
            .Cell(j).CurveTo = -1
          ElseIf .Cell(j).CurveTo > Item Then
            .Cell(j).CurveTo = .Cell(j).CurveTo - 1
          End If
        Next j
      End With
    End If
  Next C
  
  m_ActiveItem(lColumn) = -1
  m_SelectItem(lColumn) = -1
  
  Refresh
End Sub

Private Function MousePointerHands(ByVal NewValue As Boolean)
  If NewValue Then
    If Ambient.UserMode Then
      UserControl.MousePointer = vbCustom
      UserControl.MouseIcon = GetSystemHandCursor
    End If
  Else
    If hCur Then DestroyCursor hCur: hCur = 0
    UserControl.MousePointer = vbDefault
    UserControl.MouseIcon = Nothing
  End If
End Function

Private Function GetSystemHandCursor() As Picture
  Dim Pic As PicBmp
  Dim IPic As IPicture
  Dim GUID(0 To 3) As Long
  
  If hCur Then DestroyCursor hCur: hCur = 0
  
  hCur = LoadCursor(ByVal 0&, IDC_HAND)
   
  GUID(0) = &H7BF80980
  GUID(1) = &H101ABF32
  GUID(2) = &HAA00BB8B
  GUID(3) = &HAB0C3000
  
  With Pic
    .Size = Len(Pic)
    .type = vbPicTypeIcon
    .hBmp = hCur
    .hpal = 0
  End With
  
  Call OleCreatePictureIndirect(Pic, GUID(0), 1, IPic)
  Set GetSystemHandCursor = IPic
End Function

Private Sub CreateBuffer(ByVal lWidth As Long, ByVal lHeight As Long)
    If lWidth <= 0 Then lWidth = 1
    If lHeight <= 0 Then lHeight = 1
    
    If lWidth = m_MemWidth And lHeight = m_MemHeight And m_hMemDC <> 0 Then Exit Sub
    
    DestroyBuffer
    
    Dim hdc As Long
    hdc = UserControl.hdc
    m_hMemDC = CreateCompatibleDC(hdc)
    If m_hMemDC <> 0 Then
        m_hMemBmp = CreateCompatibleBitmap(hdc, lWidth, lHeight)
        If m_hMemBmp <> 0 Then
            m_hOldBmp = SelectObject(m_hMemDC, m_hMemBmp)
            m_MemWidth = lWidth
            m_MemHeight = lHeight
        End If
    End If
End Sub

Private Sub DestroyBuffer()
    If m_hMemDC <> 0 Then
        If m_hOldBmp <> 0 Then
            SelectObject m_hMemDC, m_hOldBmp
            m_hOldBmp = 0
        End If
        If m_hMemBmp <> 0 Then
            DeleteObject m_hMemBmp
            m_hMemBmp = 0
        End If
        DeleteDC m_hMemDC
        m_hMemDC = 0
    End If
    m_MemWidth = 0
    m_MemHeight = 0
End Sub

Private Sub Draw()
  Dim rHeader As RECTL, sHeader As RECTS
  Dim CELDA As RECTL, IcoBox As RECTS
  Dim cp1REC As RECTS, cp2REC As RECTS
  Dim rLabel As RECTL, SLabel As RECTS
  Dim IcoBoxL As RECTL
  Dim i As Long, C As Long
  Dim lY As Long, lX As Long
  Dim lMargen As Long
  Dim mBorder As Long, lBorder As Long
  Dim ucTH As Long
  Dim pGen As Long
  Dim hGraphics As Long
  
  If ColCount <= 0 Then Exit Sub
  If mRedraw = False Then Exit Sub
  
  CreateBuffer UserControl.ScaleWidth, UserControl.ScaleHeight
  If m_hMemDC = 0 Then Exit Sub
  
  GdipCreateFromHDC m_hMemDC, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
  GdipGraphicsClear hGraphics, RGBA(m_BackColor, 100)

  lBorder = m_BorderWidth * 2 * nScale
  mBorder = m_BorderWidth * nScale
  lMargen = (m_CornerCurve / Screen.TwipsPerPixelX) + 5
  
  lY = -ucScrollV.Value
  lX = -ucScrollH.Value
  
  ucTH = UserControl.TextHeight("?hj") + 10 * nScale
  
  ' PASS 0: Calculate ALL cell attachment points in virtual canvas coordinates
  For C = 0 To ColCount - 1
    With m_Col(C)
      For i = 0 To ItemCount(C) - 1
        pGen = (m_ItemHeight * i) + m_ItemHeight
        .Cell(i).PointLX = lX + .StartX + mBorder
        .Cell(i).PointRX = lX + .StartX + .Width + mBorder
        .Cell(i).PointY = lY + pGen + lBorder + (m_ItemHeight * 0.5) + .StartY
      Next i
    End With
  Next C

  ' LAYER 1: DRAW CONNECTION RELATIONS & CURVES (UNDERNEATH TABLES)
  For C = 0 To ColCount - 1
    If m_Col(C).Visible Then
      For i = 0 To ItemCount(C) - 1
        With m_Col(C).Cell(i)
          ' Inter-column relations (LineTo)
          If .LineTo > -1 And .ColumnTo > -1 And .ColumnTo < ColCount Then
            If .LineTo < ItemCount(.ColumnTo) And m_Col(.ColumnTo).Visible Then
              DrawRelationBezier hGraphics, C, i, .ColumnTo, .LineTo, _
                                 .LineStyle, .LineWidth, .LineStartCap, .LineEndCap, _
                                 .LineStartColor, .LineEndColor, .LineOpacity
            End If
          End If
          
          ' Intra-column curves (CurveTo)
          If .CurveTo > -1 And .CurveTo < ItemCount(C) Then
            m_X1 = IIf(.SideTo = toLeft, .PointLX, .PointRX)
            m_Y1 = .PointY
            m_X2 = IIf(.SideTo = toLeft, m_Col(C).Cell(.CurveTo).PointLX, m_Col(C).Cell(.CurveTo).PointRX)
            m_Y2 = m_Col(C).Cell(.CurveTo).PointY
            DrawCurve hGraphics, m_X1, m_Y1, m_X2, m_Y2, .Radius, .SideTo, .Direction, _
                      .LineStyle, .LineWidth, .LineStartCap, .LineEndCap, _
                      .LineStartColor, .LineEndColor, .LineOpacity
          End If
        End With
      Next i
    End If
  Next C

  ' LAYER 2: DRAW COLUMNS, HEADERS, CELLS & BADGES (ON TOP)
  For C = 0 To ColCount - 1
    With m_Col(C)
      If .Visible Then
        ' Draw Header
        SetRECL rHeader, lX + mBorder + .StartX, lY + lBorder + .StartY, .Width, m_ItemHeight - lBorder
        SetRECS sHeader, rHeader.Left, rHeader.Top, rHeader.Width, rHeader.Height
        
        If rHeader.Left + rHeader.Width >= 0 And rHeader.Left <= UserControl.ScaleWidth And _
           rHeader.Top + rHeader.Height >= 0 And rHeader.Top <= UserControl.ScaleHeight Then
            DrawRoundRect hGraphics, rHeader, RGBA(.BackColor, SafeRange(m_BoxOpacity + 10, 0, 100)), RGBA(m_BorderColor, m_BoxOpacity), m_BorderWidth, m_CornerCurve, True
            DrawCaption hGraphics, .Header, m_HeaderFont, sHeader, RGBA(.ForeColor, m_FontOpacity), 0, eCenter, eMiddle, 0, 0, False
        End If
        
        ' Draw Cells
        For i = 0 To ItemCount(C) - 1
          pGen = (m_ItemHeight * i) + m_ItemHeight
          SetRECL CELDA, lX + mBorder + .StartX, lY + pGen + lBorder + .StartY, .Width, m_ItemHeight - lBorder
          
          If CELDA.Left + CELDA.Width >= 0 And CELDA.Left <= UserControl.ScaleWidth And _
             CELDA.Top + CELDA.Height >= 0 And CELDA.Top <= UserControl.ScaleHeight Then
             
              SetRECL rLabel, CELDA.Left + ((CELDA.Width / 3) * 2), CELDA.Top + 5, (CELDA.Width / 3) - 5, ucTH - 5
              SetRECS SLabel, CELDA.Left + ((CELDA.Width / 3) * 2), CELDA.Top + 5, (CELDA.Width / 3) - 5, ucTH - 5
              SetRECS IcoBox, CELDA.Left + ((CELDA.Width / 3) * 2), CELDA.Top + 5, (CELDA.Width / 3) - 5, CELDA.Height - 10
              SetRECS cp1REC, CELDA.Left + lMargen, rLabel.Top + mBorder, CELDA.Width - lMargen, CELDA.Height / 2
              SetRECS cp2REC, CELDA.Left + lMargen, rLabel.Top + rLabel.Height, CELDA.Width - lMargen, CELDA.Height / 2
              
              DrawRoundRect hGraphics, CELDA, RGBA(IIf(i = m_ActiveItem(C), m_ColorActive, m_BoxColor), SafeRange(m_BoxOpacity + 10, 0, 100)), RGBA(IIf(i = m_SelectItem(C), m_BorderColorActive, m_BorderColor), m_BoxOpacity), m_BorderWidth, m_CornerCurve, True
              DrawCaption hGraphics, .Cell(i).Text, m_Font1, cp1REC, RGBA(m_ForeColor1, m_FontOpacity), 0, m_CaptionAlignH, m_CaptionAlignV, 0, 0, False
              
              Dim bShowSub As Boolean
              bShowSub = False
              Select Case m_SubTextVisible
                Case Is = scAllways
                  bShowSub = True
                Case Is = scOnlyActive
                  If i = m_SelectItem(C) Then bShowSub = True
              End Select
              If bShowSub Then
                DrawCaption hGraphics, .Cell(i).SubText, m_Font2, cp2REC, RGBA(m_ForeColor2, m_FontOpacity), 0, m_SubTextAlignH, m_SubTextAlignV, 0, 0, False
              End If
              
              Dim bShowBadge As Boolean
              bShowBadge = False
              Select Case m_BadgeVisible
                Case Is = scAllways
                  bShowBadge = True
                Case Is = scOnlyActive
                  If i = m_SelectItem(C) Then bShowBadge = True
              End Select
              
              If bShowBadge Then
                If m_BadgeType = vbLabel Then
                  DrawRoundRect hGraphics, rLabel, RGBA(m_BorderColor, m_BoxOpacity), RGBA(m_BorderColor, m_BoxOpacity), 1, m_CornerCurve, True
                  DrawCaption hGraphics, .Cell(i).Label, m_LabelFont, SLabel, RGBA(m_ForeColor1, m_FontOpacity), 0, eCenter, eMiddle, 0, 0, False
                ElseIf m_BadgeType = vbIcon Then
                  If m_ImageType = eIconFont Then
                    DrawCaption hGraphics, .Cell(i).IconChar, IconFont, IcoBox, RGBA(m_IconForeColor, 100), 0, eCenter, eMiddle, 0, 0, True
                  Else
                    SetRECL IcoBoxL, CLng(IcoBox.Left), CLng(IcoBox.Top), CLng(IcoBox.Width), CLng(IcoBox.Height)
                    LoadPictureFromFile .Cell(i).IconFile
                    Call GdipSetInterpolationMode(hGraphics, 7&)
                    Call GdipSetPixelOffsetMode(hGraphics, 4&)
                    GdipDrawImageRectI hGraphics, m_Bmp, IcoBoxL.Left, IcoBoxL.Top, IcoBoxL.Width, IcoBoxL.Height
                  End If
                End If
              End If
              
          End If
        Next i
      End If
    End With
  Next C
  
  GdipDeleteGraphics hGraphics
  
  BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_hMemDC, 0, 0, vbSrcCopy
  UserControl.Refresh
End Sub

Private Sub DrawRelationBezier(ByVal hGraphics As Long, ByVal Col1 As Long, ByVal Item1 As Long, _
                               ByVal Col2 As Long, ByVal Item2 As Long, _
                               ByVal fLineStyle As DashStyle, ByVal fLineWidth As Single, _
                               ByVal fStartCap As LineCap, ByVal fEndCap As LineCap, _
                               ByVal fStartColor As OLE_COLOR, ByVal fEndColor As OLE_COLOR, _
                               ByVal fOpacity As Long)
    Dim P1 As POINTS, P2 As POINTS, P3 As POINTS, P4 As POINTS
    Dim col1Left As Single, col1Right As Single, Y1 As Single
    Dim col2Left As Single, col2Right As Single, Y2 As Single
    Dim dx As Single
    Dim hPen As Long, hBrush As Long
    
    col1Left = m_Col(Col1).Cell(Item1).PointLX
    col1Right = m_Col(Col1).Cell(Item1).PointRX
    Y1 = m_Col(Col1).Cell(Item1).PointY
    
    col2Left = m_Col(Col2).Cell(Item2).PointLX
    col2Right = m_Col(Col2).Cell(Item2).PointRX
    Y2 = m_Col(Col2).Cell(Item2).PointY
    
    If col2Left >= col1Right - 10 Then
        ' Target column is to the right of source column -> S-Curve from Right to Left
        P1.X = col1Right
        P1.Y = Y1
        P4.X = col2Left
        P4.Y = Y2
        
        dx = (P4.X - P1.X) * 0.5
        If dx < 35 * nScale Then dx = 35 * nScale
        
        P2.X = P1.X + dx
        P2.Y = P1.Y
        P3.X = P4.X - dx
        P3.Y = P4.Y
        
    ElseIf col2Right <= col1Left + 10 Then
        ' Target column is to the left of source column -> S-Curve from Left to Right
        P1.X = col1Left
        P1.Y = Y1
        P4.X = col2Right
        P4.Y = Y2
        
        dx = (P1.X - P4.X) * 0.5
        If dx < 35 * nScale Then dx = 35 * nScale
        
        P2.X = P1.X - dx
        P2.Y = P1.Y
        P3.X = P4.X + dx
        P3.Y = P4.Y
        
    Else
        ' Columns overlap horizontally (stacked or partially overlapping) -> C-Curve around side
        Dim mid1 As Single, mid2 As Single
        mid1 = (col1Left + col1Right) * 0.5
        mid2 = (col2Left + col2Right) * 0.5
        
        If mid2 >= mid1 Then
            P1.X = col1Right
            P1.Y = Y1
            P4.X = col2Right
            P4.Y = Y2
            
            Dim rMax As Single
            rMax = IIf(P1.X > P4.X, P1.X, P4.X) + 40 * nScale
            P2.X = rMax
            P2.Y = P1.Y
            P3.X = rMax
            P3.Y = P4.Y
        Else
            P1.X = col1Left
            P1.Y = Y1
            P4.X = col2Left
            P4.Y = Y2
            
            Dim lMin As Single
            lMin = IIf(P1.X < P4.X, P1.X, P4.X) - 40 * nScale
            P2.X = lMin
            P2.Y = P1.Y
            P3.X = lMin
            P3.Y = P4.Y
        End If
    End If
    
    GdipCreatePen1 RGBA(fStartColor, fOpacity), fLineWidth * nScale, UnitPixel, hPen
    GdipCreateLineBrush P1, P4, RGBA(fStartColor, fOpacity), RGBA(fEndColor, fOpacity), WrapModeTileFlipXY, hBrush
    
    GdipSetPenBrushFill hPen, hBrush
    GdipSetPenDashStyle hPen, fLineStyle
    GdipSetPenStartCap hPen, fStartCap
    GdipSetPenEndCap hPen, fEndCap
    
    GdipDrawBezier hGraphics, hPen, P1.X, P1.Y, P2.X, P2.Y, P3.X, P3.Y, P4.X, P4.Y
    
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
End Sub

Private Function DrawBubble(ByVal hGraphics As Long, RCT As RECTL, BorderColor As Long, BorderWidth As Long, BackColor As Long, lCurve As Long, coWidth As Long, coLen As Long, COPos As CallOutPosition) As Long
    Dim mpath As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mRound As Long
    Dim Xx As Long, Yy As Long
    Dim lMax As Long
    Dim coAngle  As Long

With RCT
        
    coAngle = coWidth / 2

    mRound = GetSafeRound(lCurve * nScale, .Width, .Height)
    
    Select Case COPos
        Case coLeft
            .Left = .Left + coLen
            .Width = .Width - coLen
            lMax = .Height - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coTop
            .Top = .Top + coLen
            .Height = .Height - coLen
            lMax = .Width - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coRight
            .Width = .Width - coLen
            lMax = .Height - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coBottom
            .Height = .Height - coLen
            lMax = .Width - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
    End Select

    If BorderWidth >= 1 Then GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, hPen
    GdipCreateSolidFill BackColor, hBrush

    GdipCreatePath &H0, mpath

    Select Case COPos
        Case coLeft
            Yy = .Top + (.Height / 2)
            Xx = .Left
            GdipAddPathLineI mpath, Xx, Yy - coAngle, Xx - coLen, Yy
            GdipAddPathLineI mpath, Xx - coLen, Yy, Xx, Yy + coAngle
        Case coTop
            Yy = .Top
            Xx = .Left + (.Width / 2)
            GdipAddPathLineI mpath, Xx + coAngle, Yy, Xx, Yy - coLen
            GdipAddPathLineI mpath, Xx, Yy - coLen, Xx - coAngle, Yy
        Case coRight
            Yy = .Top + (.Height / 2)
            Xx = .Left + .Width
            GdipAddPathLineI mpath, Xx, Yy + coAngle, Xx + coLen, Yy
            GdipAddPathLineI mpath, Xx + coLen, Yy, Xx, Yy - coAngle
        Case coBottom
            Yy = .Top + .Height
            Xx = .Left + (.Width / 2)
            GdipAddPathLineI mpath, Xx - coAngle, Yy, Xx, Yy + coLen
            GdipAddPathLineI mpath, Xx, Yy + coLen, Xx + coAngle, Yy
    End Select
    
    'Corners
    GdipAddPathArcI mpath, .Left, .Top, mRound, mRound, 180, 90
    GdipAddPathArcI mpath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
    GdipAddPathArcI mpath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
    GdipAddPathArcI mpath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90

    GdipClosePathFigures mpath

    GdipFillPath hGraphics, hBrush, mpath
    If BorderWidth >= 1 Then GdipDrawPath hGraphics, hPen, mpath

    Call GdipDeletePath(mpath)
    Call GdipDeleteBrush(hBrush)
    If BorderWidth >= 1 Then Call GdipDeletePen(hPen)

End With
End Function

Private Function DrawCaption(ByVal hGraphics As Long, sString As Variant, oFont As StdFont, layoutRect As RECTS, _
                             ByVal TextColor As Long, ByVal mAngle As Single, _
                             ByVal H_Align As eTextAlignH, ByVal V_Align As eTextAlignV, _
                             ByVal CapX As Long, ByVal CapY As Long, Optional Icon As Boolean) As Long

    Dim hFormat      As Long
    Dim hBrush       As Long
    Dim hFontFamily  As Long
    Dim lFontSize    As Long
    Dim lFontStyle   As Long
    Dim hPath        As Long
    Dim newX         As Single
    Dim newY         As Single
    Dim Ret          As Long
    
    If Trim$(sString) <> vbNullString Then
      
      GdipCreatePath &H0, hPath
      Call GetFontStyleAndSize(oFont, lFontStyle, lFontSize)
      Call GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily)
      
      If hFontFamily = 0 Then
          GdipGetGenericFontFamilySansSerif hFontFamily
      End If
      
      GdipCreateStringFormat 0, 0, hFormat
      GdipSetStringFormatAlign hFormat, H_Align
      GdipSetStringFormatLineAlign hFormat, V_Align
      GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisCharacter
      
        If mAngle <> 0 Then
            newY = (layoutRect.Height / 2)
            newX = (layoutRect.Width / 2)
            Call GdipTranslateWorldTransform(hGraphics, newX, newY, 0)
            Call GdipRotateWorldTransform(hGraphics, mAngle, 0)
            Call GdipTranslateWorldTransform(hGraphics, -newX, -newY, 0)
        End If
        
         layoutRect.Left = layoutRect.Left + CapX
         layoutRect.Top = layoutRect.Top + CapY
         
      If Icon Then
        GdipAddPathString hPath, StrPtr(ChrW2(sString)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
      Else
        GdipAddPathString hPath, StrPtr(sString), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
      End If
      
        GdipDeleteStringFormat hFormat
        GdipCreateSolidFill TextColor, hBrush
        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush
        If mAngle <> 0 Then GdipResetWorldTransform hGraphics
        GdipDeleteFontFamily hFontFamily
        GdipDeletePath hPath
    End If

End Function

Private Sub DrawLine(hGraphics As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                     fLineStyle As DashStyle, fLineWidth As Single, _
                     fStarCap As LineCap, fEndCap As LineCap, _
                     fStartColor As OLE_COLOR, fEndColor As OLE_COLOR, fOpacity As Long)
  Dim hPen As Long
  Dim hBrush As Long
  Dim lP1 As POINTS
  Dim lP2 As POINTS
      
  lP1.X = X1: lP1.Y = Y1
  lP2.X = X2: lP2.Y = Y2

  GdipCreatePen1 RGBA(fStartColor, fOpacity), fLineWidth * nScale, UnitPixel, hPen
  GdipCreateLineBrush lP1, lP2, RGBA(fStartColor, fOpacity), RGBA(fEndColor, fOpacity), WrapModeTileFlipXY, hBrush
  
  GdipSetPenBrushFill hPen, hBrush
  GdipSetPenDashStyle hPen, fLineStyle
  GdipSetPenStartCap hPen, fStarCap
  GdipSetPenEndCap hPen, fEndCap
  
  GdipDrawLine hGraphics, hPen, lP1.X, lP1.Y, lP2.X, lP2.Y
  
  Call GdipDeleteBrush(hBrush)
  Call GdipDeletePen(hPen)
End Sub

Private Sub DrawCurve(hGraphics As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                     iRadius As Single, lSide As tSide, lDirection As tDirection, _
                     fLineStyle As DashStyle, fLineWidth As Single, _
                     fStarCap As LineCap, fEndCap As LineCap, _
                     fStartColor As OLE_COLOR, fEndColor As OLE_COLOR, fOpacity As Long)
  Dim hPen As Long
  Dim hBrush As Long
  Dim lP1 As POINTS
  Dim lP2 As POINTS
  Dim lP3 As POINTS
  Dim lP4 As POINTS
      
  lP1.X = X1: lP1.Y = Y1
  lP4.X = X2: lP4.Y = Y2
  
  Dim xSide As tSide
  
  If lSide = Auto Then
    If X1 >= 0 And X1 <= UserControl.ScaleWidth - (m_ColWidth + (iRadius * 1.5)) Then
      xSide = toRight
    ElseIf X1 >= m_ColWidth + (iRadius * 1.5) And X1 <= UserControl.ScaleWidth - m_ColWidth Then
      xSide = toLeft
    End If
  Else
    xSide = lSide
  End If
  
  lP2.X = X1 + IIf(xSide = toRight, iRadius, -iRadius)
  lP2.Y = Y1 + IIf(lDirection = 0, -(iRadius / 2), (iRadius / 2))
  lP3.X = X2 + IIf(xSide = toRight, iRadius, -iRadius)
  lP3.Y = Y2 + IIf(lDirection = 0, (iRadius / 2), -(iRadius / 2))

  GdipCreatePen1 RGBA(fStartColor, fOpacity), fLineWidth * nScale, UnitPixel, hPen
  GdipCreateLineBrush lP1, lP2, RGBA(fStartColor, fOpacity), RGBA(fEndColor, fOpacity), WrapModeTileFlipXY, hBrush
  
  GdipSetPenBrushFill hPen, hBrush
  GdipSetPenDashStyle hPen, fLineStyle
  GdipSetPenStartCap hPen, fStarCap
  GdipSetPenEndCap hPen, fEndCap
  
  GdipDrawBezier hGraphics, hPen, lP1.X, lP1.Y, lP2.X, lP2.Y, lP3.X, lP3.Y, lP4.X, lP4.Y
  
  Call GdipDeleteBrush(hBrush)
  Call GdipDeletePen(hPen)
End Sub

Private Function DrawRoundRect(ByVal hGraphics As Long, Rect As RECTL, ByVal BackColor As Long, _
                               ByVal BorderColor As Long, ByVal BorderWidth As Long, _
                               ByVal Round As Long, Filled As Boolean) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mpath As Long
    Dim mRound As Long
    
    If m_BorderWidth > 0 Then GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPen
    If Filled Then GdipCreateSolidFill BackColor, hBrush
    
    GdipCreatePath &H0, mpath
    
    With Rect
        mRound = GetSafeRound((Round * nScale), .Width * 2, .Height * 2)
        If mRound = 0 Then mRound = 1
            GdipAddPathArcI mpath, .Left, .Top, mRound, mRound, 180, 90
            GdipAddPathArcI mpath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
            GdipAddPathArcI mpath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
            GdipAddPathArcI mpath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90
    End With
    
    GdipClosePathFigures mpath
    GdipFillPath hGraphics, hBrush, mpath
    If m_BorderWidth > 0 Then GdipDrawPath hGraphics, hPen, mpath
    
    Call GdipDeletePath(mpath)
    If Filled Then Call GdipDeleteBrush(hBrush)
    If m_BorderWidth > 0 Then Call GdipDeletePen(hPen)
End Function

Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long)
On Error GoTo ErrO
    Dim hdc As Long
    lFontStyle = 0
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
    
    hdc = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)
    ReleaseDC 0&, hdc
ErrO:
End Function

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

Private Function GetItem(lColumn As Long, ByVal Y As Single) As Long
    Dim absY As Long
    If lColumn < 0 Or lColumn >= ColCount Then
        GetItem = -1
        Exit Function
    End If
    
    absY = (Y + ucScrollV.Value) - (m_Col(lColumn).StartY + m_ItemHeight)
    If absY < 0 Then
        GetItem = -1
    Else
        GetItem = (absY \ m_ItemHeight)
        If GetItem >= ItemCount(lColumn) Then GetItem = -1
    End If
End Function

Private Function GetColumn(ByVal X As Single, ByVal Y As Single) As Long
  Dim C As Long
  Dim absX As Long, absY As Long
  
  absX = X + ucScrollH.Value
  absY = Y + ucScrollV.Value
  
  GetColumn = -1
  If ColCount <= 0 Then Exit Function
  
  For C = 0 To ColCount - 1
    With m_Col(C)
      If .Visible Then
        If absX >= .StartX And absX <= (.StartX + .Width) And _
           absY >= .StartY And absY <= (.StartY + (ItemCount(C) + 1) * m_ItemHeight) Then
          GetColumn = C
          Exit Function
        End If
      End If
    End With
  Next C
End Function

Private Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double, LPY As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    LPY = CDbl(GetDeviceCaps(hdc, LOGPIXELSY))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Function IconCharCode(ByVal New_IconCharCode As String) As Long
  If Trim$(New_IconCharCode) <> "" Or New_IconCharCode <> vbNullString Then
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        IconCharCode = "&H" & New_IconCharCode
    Else
        IconCharCode = New_IconCharCode
    End If
  End If
End Function

Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

Private Function LoadPictureFromFile(ByVal FileName As String) As Boolean
  Dim imgW As Long, imgH As Long
  Dim BmpW As Single, BmpH As Single
  Dim Bmp  As Long, Grph As Long

  If m_Bmp Then
      Call GdipDisposeImage(m_Bmp)
      m_Bmp = 0
  End If
  
  If Len(FileName) = 0 Then Exit Function

  Call GdipLoadImageFromFile(StrPtr(FileName), Bmp)
  If Bmp = 0 Then Exit Function

  Call GdipGetImageDimension(Bmp, BmpW, BmpH)
  imgW = CLng(BmpW)
  imgH = CLng(BmpH)

  Call GdipCreateBitmapFromScan0(imgW, imgH, 0, &H26200A, 0, m_Bmp)
  Call GdipGetImageGraphicsContext(m_Bmp, Grph)
  Call GdipDrawImageRectRectI(Grph, Bmp, 0, 0, imgW, imgH, 0, 0, imgW, imgH, &H2, 0, 0, 0)
  Call GdipDisposeImage(Bmp)
  Call GdipDeleteGraphics(Grph)

  LoadPictureFromFile = True
End Function

Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
    Dim i       As Long
    For i = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(i) = lProp Then
            ReadValue = TlsGetValue(i + 1)
            Exit Function
        End If
    Next
    ReadValue = Default
End Function

Private Function RGBA(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
  If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
  RGBA = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
  Opacity = CByte((Abs(Opacity) / 100) * 255)
  If Opacity < 128 Then
      If Opacity < 0& Then Opacity = 0&
      RGBA = RGBA Or Opacity * &H1000000
  Else
      If Opacity > 255& Then Opacity = 255&
      RGBA = RGBA Or (Opacity - 128&) * &H1000000 Or &H80000000
  End If
End Function

Private Function SafeRange(Value, Min, Max) As Long
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
    SafeRange = Value
End Function

Private Function SetRECS(lpRect As RECTS, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long
  lpRect.Left = X
  lpRect.Top = Y
  lpRect.Width = W
  lpRect.Height = H
End Function

Private Sub TerminateGDI()
    DestroyBuffer
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub ucScrollH_Change()
  Refresh
End Sub

Private Sub ucScrollH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ucScrollV.TrackMouseWheelOnHwndStop
  ucScrollH.TrackMouseWheelOnHwnd UserControl.hwnd
End Sub

Private Sub ucScrollH_Scroll()
  Refresh
End Sub

Private Sub ucScrollV_Change()
  Refresh
End Sub

Private Sub ucScrollV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ucScrollH.TrackMouseWheelOnHwndStop
  ucScrollV.TrackMouseWheelOnHwnd UserControl.hwnd
End Sub

Private Sub ucScrollV_Scroll()
  Refresh
End Sub

Private Sub UserControl_Initialize()
    InitGDI
    nScale = GetWindowsDPI
    SideStart = -1
End Sub

Private Sub UserControl_InitProperties()
  hFontCollection = ReadValue(&HFC)
  
  m_Enabled = True
  m_Clickable = False
  mRedraw = True
  m_Editable = False
  
  m_BorderColor = &HFF8080
  m_BorderColorActive = vbRed
  m_BackColor = &H8000000F
  m_BoxColor = vbRed
  m_BorderWidth = 1
  m_ForeColor1 = &HFFFFFF
  m_ForeColor2 = &HFFFFFF
  m_IconForeColor = &HFFFFFF
  Set m_Font1 = UserControl.Font
  Set m_Font2 = UserControl.Font
  Set m_IconFont = UserControl.Font
  Set m_LabelFont = UserControl.Font
  m_CaptionAlignV = 1
  m_CaptionAlignH = 1
  m_SubTextAlignV = 1
  m_SubTextAlignH = 1
  m_ItemHeight = 60
  m_ColWidth = 150
  m_ColorActive = vbRed
  m_BadgeType = False
  m_CornerCurve = 5
  m_BoxOpacity = 90
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 37, 38
        ucScrollV.Value = IIf(ucScrollV.Value > (m_ItemHeight / 2), ucScrollV.Value - (m_ItemHeight / 2), 0)
    Case 39, 40
        ucScrollV.Value = IIf(ucScrollV.Value + (m_ItemHeight / 2) < ucScrollV.Max, ucScrollV.Value + (m_ItemHeight / 2), ucScrollV.Max)
  End Select
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_DblClick()
  On Error Resume Next
  If m_Editable = True Then
    m_ActiveCol = GetColumn(mX, mY)
    m_ActiveItem(m_ActiveCol) = GetItem(m_ActiveCol, mY)

    If SideStart = -1 Then
      SideStart = m_ActiveCol
      InitCurve = m_ActiveItem(m_ActiveCol)
    Else
      SideEnd = m_ActiveCol
      If SideStart = SideEnd Then
        EndCurve = m_ActiveItem(m_ActiveCol)
        AddCurve SideStart, InitCurve, EndCurve, IIf(m_DefaultSide = Auto, IIf(m_ActiveCol > 0, toLeft, toRight), m_DefaultSide), 50, m_LineStyle, m_LineStartCap, m_LineEndCap, m_LineStartColor, m_LineEndColor, m_LineWidth, m_LineOpacity, True
        SideStart = -1
      Else
        If m_ActiveItem(SideStart) > -1 And m_ActiveItem(m_ActiveCol) > -1 Then
          AddLine SideStart, m_ActiveItem(SideStart), SideEnd, m_ActiveItem(m_ActiveCol), m_LineStyle, m_LineStartCap, m_LineEndCap, m_LineStartColor, m_LineEndColor, m_LineWidth, m_LineOpacity, True
          SideStart = -1
        End If
      End If
    End If
  Else
    m_ActiveItem(SideStart) = -1
    m_ActiveItem(SideEnd) = -1
  End If

  Refresh
  RaiseEvent DblClick(m_ActiveCol, m_SelectItem(m_ActiveCol))
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mY = Y
  mX = X
  mCol = GetColumn(X, Y)

  If mCol > -1 Then
    m_ActiveCol = mCol
    m_SelectItem(mCol) = -1
    m_ActiveItem(mCol) = -1

    If m_ColsMoveable And Button = 1 Then
      m_IsDragging = True
      m_DragOffsetX = (X + ucScrollH.Value) - m_Col(mCol).StartX
      m_DragOffsetY = (Y + ucScrollV.Value) - m_Col(mCol).StartY
    End If

    If m_Clickable Then
      m_SelectItem(mCol) = GetItem(mCol, Y)
      RaiseEvent MouseDown(Button, Shift, X, Y)
      With m_Col(mCol)
        If m_SelectItem(mCol) <> -1 Then RaiseEvent Click(mCol, GetItem(mCol, Y), .Cell(GetItem(mCol, Y)).ColumnTo, .Cell(GetItem(mCol, Y)).LineTo, .Cell(GetItem(mCol, Y)).CurveTo)
      End With
      Refresh
    Else
      RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
  Else
    m_IsDragging = False
    RaiseEvent MouseDown(Button, Shift, X, Y)
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If m_Editable Then
    MousePointerHands True
  Else
    MousePointerHands False
  End If

  If m_ColsMoveable And m_IsDragging And mCol > -1 And Button = 1 Then
    Dim newX As Long, newY As Long
    newX = (X + ucScrollH.Value) - m_DragOffsetX
    newY = (Y + ucScrollV.Value) - m_DragOffsetY
    If newX < 0 Then newX = 0
    If newY < 0 Then newY = 0
    m_Col(mCol).StartX = newX
    m_Col(mCol).StartY = newY
    Refresh
  End If

  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  m_IsDragging = False
  If m_ColsMoveable Then mCol = -1
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    m_Enabled = .ReadProperty("Enabled", True)
    m_Clickable = .ReadProperty("Clickable", False)
    m_Editable = .ReadProperty("Editable", False)
    
    m_BorderColor = .ReadProperty("BorderColor", &HFF8080)
    m_BorderColorActive = .ReadProperty("BorderColorActive", vbRed)
    m_BackColor = .ReadProperty("BackColor", &H8000000F)
    m_BoxColor = .ReadProperty("BoxColor", vbRed)
    m_BorderWidth = .ReadProperty("BorderWidth", 1)
    m_ForeColor1 = .ReadProperty("TextColor", &HFFFFFF)
    m_ForeColor2 = .ReadProperty("SubTextColor", &HFFFFFF)
    m_HeaderBackColor = .ReadProperty("HeaderBackColor", vbBlack)
    m_HeaderForeColor = .ReadProperty("HeaderForeColor", vbWhite)
    Set m_Font1 = .ReadProperty("TextFont", UserControl.Font)
    Set m_Font2 = .ReadProperty("SubTextFont", UserControl.Font)
    Set m_LabelFont = .ReadProperty("LabelFont", UserControl.Font)
    Set m_HeaderFont = .ReadProperty("HeaderFont", UserControl.Font)
    m_FontOpacity = .ReadProperty("FontOpacity", 100)
    m_CaptionAlignV = .ReadProperty("TextAlignV", 1)
    m_CaptionAlignH = .ReadProperty("TextAlignH", 1)
    m_SubTextAlignV = .ReadProperty("SubTextAlignV", 1)
    m_SubTextAlignH = .ReadProperty("SubTextAlignH", 1)
    m_IconForeColor = .ReadProperty("IconForeColor", &HFFFFFF)
    Set m_IconFont = .ReadProperty("IconFont", UserControl.Font)
    m_ImageType = .ReadProperty("ImageType", 0)
    m_BadgeVisible = .ReadProperty("BadgeVisible", 1)
    m_SubTextVisible = .ReadProperty("SubTextVisible", 1)
    m_ItemHeight = .ReadProperty("ItemHeight", 60)
    m_ColWidth = .ReadProperty("ColWidth", 150)
    m_ColsMoveable = .ReadProperty("ColsMoveable", False)
    m_ColorActive = .ReadProperty("BackColorActive", vbRed)
    m_BadgeType = .ReadProperty("BadgeType", 0)
    m_CornerCurve = .ReadProperty("CornerCurve", 5)
    m_BoxOpacity = .ReadProperty("BoxOpacity", 90)
    m_DefaultSide = .ReadProperty("LineSideDefault", 1)
    m_LineWidth = .ReadProperty("LineWidth", 5)
    m_LineStyle = .ReadProperty("LineStyle", 0)
    m_LineOpacity = .ReadProperty("LineOpacity", 100)
    m_LineStartCap = .ReadProperty("LineStartCap", 0)
    m_LineEndCap = .ReadProperty("LineEndCap", 0)
    m_LineStartColor = .ReadProperty("LineStartColor", &HFF&)
    m_LineEndColor = .ReadProperty("LineEndColor", &HC67300)
  End With
End Sub

Private Sub UserControl_Resize()
    ucScrollV.Move UserControl.ScaleWidth - 11, 0, 11, UserControl.ScaleHeight - 11
    ucScrollH.Move 0, UserControl.ScaleHeight - 11, UserControl.ScaleWidth - 11, 11
    UpdateScrollBars
    If ColCount > 0 Then Refresh
End Sub

Private Sub UserControl_Show()
  If ColCount > 0 Then Refresh
End Sub

Private Sub UserControl_Terminate()
  DestroyBuffer
  TerminateGDI
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty("Enabled", m_Enabled)
    Call .WriteProperty("Clickable", m_Clickable)
    Call .WriteProperty("Editable", m_Editable)
    Call .WriteProperty("BorderColor", m_BorderColor)
    Call .WriteProperty("BorderColorActive", m_BorderColorActive)
    Call .WriteProperty("BackColor", m_BackColor)
    Call .WriteProperty("BorderWidth", m_BorderWidth)
    Call .WriteProperty("BoxColor", m_BoxColor)
    Call .WriteProperty("TextColor", m_ForeColor1)
    Call .WriteProperty("SubTextColor", m_ForeColor2)
    Call .WriteProperty("HeaderBackColor", m_HeaderBackColor)
    Call .WriteProperty("HeaderForeColor", m_HeaderForeColor)
    Call .WriteProperty("TextFont", m_Font1)
    Call .WriteProperty("SubTextFont", m_Font2)
    Call .WriteProperty("LabelFont", m_LabelFont)
    Call .WriteProperty("HeaderFont", m_HeaderFont)
    Call .WriteProperty("FontOpacity", m_FontOpacity)
    Call .WriteProperty("TextAlignV", m_CaptionAlignV)
    Call .WriteProperty("TextAlignH", m_CaptionAlignH)
    Call .WriteProperty("SubTextAlignV", m_SubTextAlignV)
    Call .WriteProperty("SubTextAlignH", m_SubTextAlignH)
    Call .WriteProperty("IconForeColor", m_IconForeColor)
    Call .WriteProperty("IconFont", m_IconFont)
    Call .WriteProperty("ImageType", m_ImageType)
    Call .WriteProperty("BadgeVisible", m_BadgeVisible)
    Call .WriteProperty("SubTextVisible", m_SubTextVisible)
    Call .WriteProperty("ItemHeight", m_ItemHeight)
    Call .WriteProperty("ColWidth", m_ColWidth)
    Call .WriteProperty("ColsMoveable", m_ColsMoveable)
    Call .WriteProperty("BackColorActive", m_ColorActive)
    Call .WriteProperty("BadgeType", m_BadgeType)
    Call .WriteProperty("CornerCurve", m_CornerCurve)
    Call .WriteProperty("BoxOpacity", m_BoxOpacity)
    Call .WriteProperty("LineSideDefault", m_DefaultSide)
    Call .WriteProperty("LineWidth", m_LineWidth)
    Call .WriteProperty("LineStyle", m_LineStyle)
    Call .WriteProperty("LineOpacity", m_LineOpacity)
    Call .WriteProperty("LineStartCap", m_LineStartCap)
    Call .WriteProperty("LineEndCap", m_LineEndCap)
    Call .WriteProperty("LineStartColor", m_LineStartColor)
    Call .WriteProperty("LineEndColor", m_LineEndColor)
  End With
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
    m_BackColor = New_Color
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get BackColorActive() As OLE_COLOR
    BackColorActive = m_ColorActive
End Property

Public Property Let BackColorActive(ByVal NewColorActive As OLE_COLOR)
    m_ColorActive = NewColorActive
    PropertyChanged "BackColorActive"
    Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
    m_BorderColor = NewBorderColor
    PropertyChanged "BorderColor"
    Refresh
End Property

Public Property Get BorderColorActive() As OLE_COLOR
    BorderColorActive = m_BorderColorActive
End Property

Public Property Let BorderColorActive(ByVal NewColorActive As OLE_COLOR)
    m_BorderColorActive = NewColorActive
    PropertyChanged "BorderColorActive"
    Refresh
End Property

Public Property Get BorderWidth() As Long
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
    m_BorderWidth = NewBorderWidth
    PropertyChanged "BorderWidth"
    Refresh
End Property

Public Property Get BoxColor() As OLE_COLOR
    BoxColor = m_BoxColor
End Property

Public Property Let BoxColor(ByVal New_Color As OLE_COLOR)
    m_BoxColor = New_Color
    PropertyChanged "BoxColor"
    Refresh
End Property

Public Property Get BoxOpacity() As Long
    BoxOpacity = m_BoxOpacity
End Property

Public Property Let BoxOpacity(ByVal newBoxOpacity As Long)
    m_BoxOpacity = newBoxOpacity
    PropertyChanged "BoxOpacity"
    Refresh
End Property

Public Property Get Clickable() As Boolean
    Clickable = m_Clickable
End Property

Public Property Let Clickable(ByVal NewClickable As Boolean)
    m_Clickable = NewClickable
    PropertyChanged "Clickable"
    Refresh
End Property

Public Property Get ColWidth() As Long
    ColWidth = m_ColWidth
End Property

Public Property Let ColWidth(ByVal NewColWidth As Long)
    m_ColWidth = NewColWidth
    PropertyChanged "ColWidth"
    Refresh
End Property

Public Property Get CornerCurve() As Long
    CornerCurve = m_CornerCurve
End Property

Public Property Let CornerCurve(ByVal NewCornerCurve As Long)
    m_CornerCurve = NewCornerCurve
    PropertyChanged "CornerCurve"
    Refresh
End Property

Public Property Get Editable() As Boolean
    Editable = m_Editable
End Property

Public Property Let Editable(ByVal NEditable As Boolean)
    m_Editable = NEditable
    PropertyChanged "Editable"
    Refresh
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    Refresh
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get HeaderBackColor() As OLE_COLOR
    HeaderBackColor = m_HeaderBackColor
End Property

Public Property Let HeaderBackColor(ByVal NewHeaderBackColor As OLE_COLOR)
    m_HeaderBackColor = NewHeaderBackColor
    PropertyChanged "HeaderBackColor"
    Refresh
End Property

Public Property Get HeaderFont() As StdFont
    Set HeaderFont = m_HeaderFont
End Property

Public Property Set HeaderFont(ByVal NewHeaderFont As StdFont)
    Set m_HeaderFont = NewHeaderFont
    PropertyChanged "HeaderFont"
    Refresh
End Property

Public Property Get HeaderForeColor() As OLE_COLOR
    HeaderForeColor = m_HeaderForeColor
End Property

Public Property Let HeaderForeColor(ByVal NewHeaderForeColor As OLE_COLOR)
    m_HeaderForeColor = NewHeaderForeColor
    PropertyChanged "HeaderForeColor"
    Refresh
End Property

Public Property Let Header(ByVal lColumn As Long, ByVal sHeader As String)
  On Error Resume Next
  With m_Col(lColumn)
    .Header = sHeader
  End With
  Refresh
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property

Public Property Set IconFont(New_Font As StdFont)
    Set m_IconFont = New_Font
    PropertyChanged "IconFont"
    Refresh
End Property

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property

Public Property Let IconForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "IconForeColor"
    Refresh
End Property

Public Property Get BadgeVisible() As eVisibleType
    BadgeVisible = m_BadgeVisible
End Property

Public Property Let BadgeVisible(ByVal newVisible As eVisibleType)
    m_BadgeVisible = newVisible
    PropertyChanged "BadgeVisible"
    Refresh
End Property

Public Property Get ImageType() As eImageType
    ImageType = m_ImageType
End Property

Public Property Let ImageType(ByVal NewImageType As eImageType)
    m_ImageType = NewImageType
    PropertyChanged "ImageType"
    Refresh
End Property

Property Get ColCount() As Long
On Error GoTo ErrC
    ColCount = UBound(m_Col) + 1
    Exit Property
ErrC:
    ColCount = 0
End Property

Property Get ItemCount(lColumn As Long) As Long
On Error Resume Next
    ItemCount = UBound(m_Col(lColumn).Cell) + 1
End Property

Public Property Get ItemHeight() As Long
    ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal NewSectionSpace As Long)
    m_ItemHeight = NewSectionSpace
    PropertyChanged "ItemHeight"
    Refresh
End Property

Public Property Get LabelFont() As StdFont
    Set LabelFont = m_LabelFont
End Property

Public Property Set LabelFont(ByVal New_Font As StdFont)
    Set m_LabelFont = New_Font
    PropertyChanged "LabelFont"
    Refresh
End Property

Public Property Get BadgeType() As eTypeBadge
    BadgeType = m_BadgeType
End Property

Public Property Let BadgeType(ByVal NewBadgeType As eTypeBadge)
    m_BadgeType = NewBadgeType
    PropertyChanged "BadgeType"
    Refresh
End Property

Property Get LinesCount() As Long
    LinesCount = 0
End Property

Property Get MaxItemCount() As Long
    Dim Count As Long, CountF As Long, C As Long
    CountF = 0
    On Error Resume Next
    For C = 0 To UBound(m_Col)
        Count = ItemCount(C)
        If Count > CountF Then CountF = Count
    Next C
    MaxItemCount = CountF
End Property

Public Property Get Redraw() As Boolean
    Redraw = mRedraw
End Property

Public Property Let Redraw(ByVal bRedraw As Boolean)
    mRedraw = bRedraw
    If mRedraw Then Refresh
End Property

Public Property Get RunMode() As Boolean
    On Error Resume Next
    RunMode = True
    RunMode = Ambient.UserMode
    RunMode = Extender.Parent.RunMode
End Property

Public Property Get SubTextAlignH() As eTextAlignH
    SubTextAlignH = m_SubTextAlignH
End Property

Public Property Let SubTextAlignH(ByVal NewCaptionAlignH As eTextAlignH)
    m_SubTextAlignH = NewCaptionAlignH
    PropertyChanged "SubTextAlignH"
    Refresh
End Property

Public Property Get SubTextAlignV() As eTextAlignV
    SubTextAlignV = m_SubTextAlignV
End Property

Public Property Let SubTextAlignV(ByVal NewCaptionAlignV As eTextAlignV)
    m_SubTextAlignV = NewCaptionAlignV
    PropertyChanged "SubTextAlignV"
    Refresh
End Property

Public Property Get SubTextColor() As OLE_COLOR
    SubTextColor = m_ForeColor2
End Property

Public Property Let SubTextColor(ByVal NewForeColor As OLE_COLOR)
    m_ForeColor2 = NewForeColor
    PropertyChanged "SubTextColor"
    Refresh
End Property

Public Property Get SubTextFont() As StdFont
    Set SubTextFont = m_Font2
End Property

Public Property Set SubTextFont(ByVal New_Font As StdFont)
    Set m_Font2 = New_Font
    PropertyChanged "SubTextFont"
    Refresh
End Property

Public Property Get SubTextVisible() As eVisibleType
    SubTextVisible = m_SubTextVisible
End Property

Public Property Let SubTextVisible(ByVal newVisible As eVisibleType)
    m_SubTextVisible = newVisible
    PropertyChanged "SubTextVisible"
    Refresh
End Property

Public Property Get TextAlignH() As eTextAlignH
    TextAlignH = m_CaptionAlignH
End Property

Public Property Let TextAlignH(ByVal NewCaptionAlignH As eTextAlignH)
    m_CaptionAlignH = NewCaptionAlignH
    PropertyChanged "TextAlignH"
    Refresh
End Property

Public Property Get TextAlignV() As eTextAlignV
    TextAlignV = m_CaptionAlignV
End Property

Public Property Let TextAlignV(ByVal NewCaptionAlignV As eTextAlignV)
    m_CaptionAlignV = NewCaptionAlignV
    PropertyChanged "TextAlignV"
    Refresh
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_ForeColor1
End Property

Public Property Let TextColor(ByVal NewForeColor As OLE_COLOR)
    m_ForeColor1 = NewForeColor
    PropertyChanged "TextColor"
    Refresh
End Property

Public Property Get TextFont() As StdFont
    Set TextFont = m_Font1
End Property

Public Property Set TextFont(ByVal New_Font As StdFont)
    Set m_Font1 = New_Font
    PropertyChanged "TextFont"
    Refresh
End Property

Public Property Get Version() As String
    Version = sVersion
End Property

Public Property Get Visible() As Boolean
    Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal newVisible As Boolean)
    Extender.Visible = newVisible
End Property

Public Property Get LineWidth() As Single
    LineWidth = m_LineWidth
End Property

Public Property Let LineWidth(ByVal NewLineWidth As Single)
    m_LineWidth = NewLineWidth
    PropertyChanged "LineWidth"
    Refresh
End Property

Public Property Get LineStyle() As DashStyle
    LineStyle = m_LineStyle
End Property

Public Property Let LineStyle(ByVal NewLineStyle As DashStyle)
    m_LineStyle = NewLineStyle
    PropertyChanged "LineStyle"
    Refresh
End Property

Public Property Get LineOpacity() As Long
    LineOpacity = m_LineOpacity
End Property

Public Property Let LineOpacity(ByVal NewLineOpacity As Long)
    m_LineOpacity = NewLineOpacity
    PropertyChanged "LineOpacity"
    Refresh
End Property

Public Property Get LineStartCap() As LineCap
    LineStartCap = m_LineStartCap
End Property

Public Property Let LineStartCap(ByVal NewLineStartCap As LineCap)
    m_LineStartCap = NewLineStartCap
    PropertyChanged "LineStartCap"
    Refresh
End Property

Public Property Get LineEndCap() As LineCap
    LineEndCap = m_LineEndCap
End Property

Public Property Let LineEndCap(ByVal NewLineEndCap As LineCap)
    m_LineEndCap = NewLineEndCap
    PropertyChanged "LineEndCap"
    Refresh
End Property

Public Property Get LineStartColor() As OLE_COLOR
    LineStartColor = m_LineStartColor
End Property

Public Property Let LineStartColor(ByVal NewLineStartColor As OLE_COLOR)
    m_LineStartColor = NewLineStartColor
    PropertyChanged "LineStartColor"
    Refresh
End Property

Public Property Get LineEndColor() As OLE_COLOR
    LineEndColor = m_LineEndColor
End Property

Public Property Let LineEndColor(ByVal NewLineEndColor As OLE_COLOR)
    m_LineEndColor = NewLineEndColor
    PropertyChanged "LineEndColor"
    Refresh
End Property

Public Property Get LineSideDefault() As tSide
    LineSideDefault = m_DefaultSide
End Property

Public Property Let LineSideDefault(ByVal NewDefaultSide As tSide)
    m_DefaultSide = NewDefaultSide
    PropertyChanged "LineSideDefault"
    Refresh
End Property

Public Property Get ColsMoveable() As Boolean
    ColsMoveable = m_ColsMoveable
End Property

Public Property Let ColsMoveable(ByVal NewColsMoveable As Boolean)
    m_ColsMoveable = NewColsMoveable
    PropertyChanged "ColsMoveable"
    Refresh
End Property

Public Property Get FontOpacity() As Long
    FontOpacity = m_FontOpacity
End Property

Public Property Let FontOpacity(ByVal NewFontOpacity As Long)
    m_FontOpacity = NewFontOpacity
    PropertyChanged "FontOpacity"
    Refresh
End Property
