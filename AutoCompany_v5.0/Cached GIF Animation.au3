#include-once
#include <GDIPlus.au3>

; #INDEX# =======================================================================================================================
; Title .........: Cached GIF Animation
; AutoIt Version : 3.3.14.5
; Language ..... : English
; Description ...: Functions to manage GIF animation
; Author ........: Nine
; Version .......: 0.2.1.4
; ===============================================================================================================================

; #GLOBALS# =====================================================================================================================
Global $aGIF_Animation[0][8]
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _GUICtrlCreateAnimGIF
; _GUICtrlDeleteAnimGIF
; _GIF_Animation_Quit
; ===============================================================================================================================

; #INTERNAL_USE_ONLY#============================================================================================================
; __GIF_Animation_DrawTimer
; __GIF_Animation_DrawFrame
; __GDIPlus_GraphicsDrawCachedBitmap
; __GDIPlus_CachedBitmapCreate
; __GDIPlus_CachedBitmapDispose
; ===============================================================================================================================

_GDIPlus_Startup()

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlCreateAnimGIF
; Description ...: Create a GIF control inside a previously GUI created window
; Syntax.........: _GUICtrlCreateAnimGIF($vSource, $iLeft, $iTop, $iWidth, $iHeight, [$iStyle = -1, [$iExStyle = -1, [$bHandle = False]]])
; Parameters ....: $vSource           - Either a GIF file name or a handle to a GIF image create by GDI+ (handle would be used as a resource)
;                  $$iLeft            - Left position of the control
;                  $iTop              - Top position of the control
;                  $iWidth            - Width of the control (must be the actual size of the GIF)
;                  $iHeight           - Height of the control (must be the actual size of the GIF)
;                  $iStyle            - Optional: Style of the control (by default : $SS_NOTIFY forced style : $SS_BITMAP)
;                  $iExStyle          - Optional: Extented style of the control (by default : null)
;                  $bHandle           - Optional: True if a GDI+ image handle of a GIF, False if a GIF file name (by default : False)
; Return values .: Success - Id of the control
;                  Failure - Returns 0 and sets @error
; Author ........: Nine
; Modified ......:
; Remarks .......: Width and Height are mandatory because of an erronous display without them
;                  As discussed here :https://www.autoitscript.com/forum/topic/153782-help-filedocumentation-issues-discussion-only/page/30/?tab=comments#comment-1438857
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlCreateAnimGIF($vSource, $iLeft, $iTop, $iWidth, $iHeight, $iStyle = -1, $iExStyle = -1, $bHandle = False)
  Local Const $GDIP_PROPERTYTAGFRAMEDELAY = 0x5100
  Local $iIdx, $idPic, $hImage, $aTime, $iNumberOfFrames, $iType

  If $bHandle Then
    $iType = _GDIPlus_ImageGetType($vSource)
    If @error Or $iType <> $GDIP_IMAGETYPE_BITMAP Then Return SetError(1, 0, 0)
  Else
    If Not FileExists($vSource) Then Return SetError(2, 0, 0)
  EndIf
  $idPic = GUICtrlCreatePic("", $iLeft, $iTop, $iWidth, $iHeight, $iStyle, $iExStyle)
  If Not $idPic Then Return SetError(3, 0, 0)
  $hImage = $bHandle ? $vSource : _GDIPlus_ImageLoadFromFile($vSource)
  If @error Then Return SetError(10 + @error, 0, 0)
  $iNumberOfFrames = _GDIPlus_ImageGetFrameCount($hImage, $GDIP_FRAMEDIMENSION_TIME)
  If @error Then Return SetError(20 + @error, 0, 0)
  $aTime = _GDIPlus_ImageGetPropertyItem($hImage, $GDIP_PROPERTYTAGFRAMEDELAY)
  If @error Then Return SetError(30 + @error, 0, 0)
  If UBound($aTime) - 1 <> $iNumberOfFrames Then Return SetError(4, 0, 0)
  For $i = 0 To UBound($aTime) - 1
    If Not $aTime[$i] Then $aTime[$i] = 5
  Next
  ; search for an empty slot left after deletion
  For $iIdx = 0 To UBound($aGIF_Animation) - 1
    If Not $aGIF_Animation[$iIdx][0] Then ExitLoop
  Next

  ; gather all pertinent informations

  If $iIdx = UBound($aGIF_Animation) Then ReDim $aGIF_Animation[$iIdx + 1][UBound($aGIF_Animation, 2)]
  $aGIF_Animation[$iIdx][0] = $idPic
  $aGIF_Animation[$iIdx][1] = $hImage
  $aGIF_Animation[$iIdx][2] = $iNumberOfFrames
  $aGIF_Animation[$iIdx][3] = $aTime  ; 1-base array
  $aGIF_Animation[$iIdx][4] = 0       ; current Frame number
  $aGIF_Animation[$iIdx][5] = TimerInit()
  $aGIF_Animation[$iIdx][6] = GUICtrlGetHandle($idPic)
  $aGIF_Animation[$iIdx][7] = _GDIPlus_GraphicsCreateFromHWND($aGIF_Animation[$iIdx][6])

  ; if first GIF, start the timer
  If UBound($aGIF_Animation) = 1 Then AdlibRegister(__GIF_Animation_DrawTimer, 10)
  Return $idPic

EndFunc   ;==>_GUICtrlCreateAnimGIF

; #FUNCTION# ====================================================================================================================
; Name...........: _GIF_Animation_Quit
; Description ...: Dispose of all object, unregister draw timer and shutdown GDI+
; Syntax.........: _GIF_Animation_Quit()
; Parameters ....: None
; Return values .: None
; Author ........: Nine
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GIF_Animation_Quit()
  AdlibUnRegister(__GIF_Animation_DrawTimer)
  For $i = 0 To UBound($aGIF_Animation) - 1
    If Not $aGIF_Animation[$i][0] Then ContinueLoop
    _GDIPlus_ImageDispose($aGIF_Animation[$i][1])
    _GDIPlus_GraphicsDispose($aGIF_Animation[$i][7])
  Next
  _GDIPlus_Shutdown()
EndFunc   ;==>_GIF_Animation_Quit

; #FUNCTION# ====================================================================================================================
; Name...........: _GUICtrlDeleteAnimGIF
; Description ...: Delete one GIF control
; Syntax.........: _GUICtrlDeleteAnimGIF($idCtrl)
; Parameters ....: $idCtrl              - Control id return from _GUICtrlCreateAnimGIF
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error (unfound control id)
; Author ........: Nine
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlDeleteAnimGIF($idPic)
  Local $iNum = 0, $bFound = False
  For $iIdx = 0 To UBound($aGIF_Animation) - 1
    If $aGIF_Animation[$iIdx][0] = $idPic Then
      $aGIF_Animation[$iIdx][0] = 0
      _GDIPlus_ImageDispose($aGIF_Animation[$iIdx][1])
      _GDIPlus_GraphicsDispose($aGIF_Animation[$iIdx][7])
      GUICtrlDelete($idPic)
      $bFound = True
    ElseIf $aGIF_Animation[$iIdx][0] Then
      $iNum += 1
    EndIf
  Next
  If Not $bFound Then Return SetError(1, 0, 0)
  If Not $iNum Then
    AdlibUnRegister(__GIF_Animation_DrawTimer)
    ReDim $aGIF_Animation[0][UBound($aGIF_Animation,2)]
  EndIf
  Return 1
EndFunc   ;==>_GUICtrlDeleteAnimGIF

; #INTERNAL_USE_ONLY#============================================================================================================
Func __GIF_Animation_DrawTimer()
  Local $aTime
  For $i = 0 To UBound($aGIF_Animation) - 1
    If Not $aGIF_Animation[$i][0] Then ContinueLoop
    $aTime = $aGIF_Animation[$i][3]
    If TimerDiff($aGIF_Animation[$i][5]) >= $aTime[$aGIF_Animation[$i][4] + 1] * 10 Then
      __GIF_Animation_DrawFrame($i)
      $aGIF_Animation[$i][4] += 1
      If $aGIF_Animation[$i][4] = $aGIF_Animation[$i][2] Then $aGIF_Animation[$i][4] = 0 ; If $iFrame = $iFrameCount then reset $iFrame to 0
      $aGIF_Animation[$i][5] = TimerInit()
    EndIf
  Next
EndFunc   ;==>__GIF_Animation_DrawTimer

Func __GIF_Animation_DrawFrame($iGIF)
  _GDIPlus_ImageSelectActiveFrame($aGIF_Animation[$iGIF][1], $GDIP_FRAMEDIMENSION_TIME, $aGIF_Animation[$iGIF][4])
  Local $hCachedBmp = __GDIPlus_CachedBitmapCreate($aGIF_Animation[$iGIF][7], $aGIF_Animation[$iGIF][1]) ; (hGraphics, $hBitmap)
  __GDIPlus_GraphicsDrawCachedBitmap($aGIF_Animation[$iGIF][7], $hCachedBmp, 0, 0) ;(hGraphics, hCachedBmp, X, Y)
  __GDIPlus_CachedBitmapDispose($hCachedBmp)
EndFunc   ;==>__GIF_Animation_DrawFrame

Func __GDIPlus_GraphicsDrawCachedBitmap($hGraphics, $hCachedBitmap, $iX, $iY)
  Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDrawCachedBitmap", "handle", $hGraphics, "handle", $hCachedBitmap, "int", $iX, "int", $iY)
  If @error Then Return SetError(@error, @extended, False)
  If $aResult[0] Then Return SetError(10, $aResult[0], False)
  Return True
EndFunc   ;==>__GDIPlus_GraphicsDrawCachedBitmap

Func __GDIPlus_CachedBitmapCreate($hGraphics, $hBitmap)
  Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipCreateCachedBitmap", "handle", $hBitmap, "handle", $hGraphics, "handle*", 0)
  If @error Then Return SetError(@error, @extended, 0)
  If $aResult[0] Then Return SetError(10, $aResult[0], 0)
  Return $aResult[3]
EndFunc   ;==>__GDIPlus_CachedBitmapCreate

Func __GDIPlus_CachedBitmapDispose($hCachedBitmap)
  Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipDeleteCachedBitmap", "handle", $hCachedBitmap)
  If @error Then Return SetError(@error, @extended, False)
  If $aResult[0] Then Return SetError(10, $aResult[0], False)
  Return True
EndFunc   ;==>__GDIPlus_CachedBitmapDispose
; ===============================================================================================================================
