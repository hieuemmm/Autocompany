#include <GUIConstantsEx.au3>
#include <file.au3>
#include <array.au3>
#Include <GuiListView.au3>
TCPStartup()
global $PROXY_LIST[1][5]
Global $hWs2_32 = -1
Global Const $__TCP_WINDOW = GUICreate("Async Sockets UDF")
;CheckProxy(True) ;Lọc file 600 proxy
;CheckProxy(False) ;Lọc file Con proxy
Func CheckProxy($TrangThai=True)
   ; Đặt trạng thái của từng proxy.
	If $TrangThai Then
		Readin_Proxies(@ScriptDir & "/Data/Proxy_List.txt")
	Else
		Readin_Proxies(@ScriptDir & "/Data/ProxyChecked.txt")
	EndIf
	;Set trạng thai cho Proxy
	For $x = UBound($PROXY_LIST)-1 to 0 step -1
		 Test_Connection( $x)
		 Sleep(100)
	Next
	While 1
		Local $Count = 0
		For $x = UBound($PROXY_LIST)-1 to 0 step -1
			if	$PROXY_LIST[$x][4] == -1 Then
				$Count+=1
				ExitLoop
			EndIf
		Next
		If $Count == 0 Then
			For $x = UBound($PROXY_LIST)-1 to 0 step -1
				If	$PROXY_LIST[$x][4] <> 1 Then
					_ArrayDelete($PROXY_LIST, $x)
				EndIf
			Next
			SaveProxyChecked()
			If $TrangThai Then
				Readin_Proxies(@ScriptDir & "/Data/Proxy_List.txt")
			Else
				Readin_Proxies(@ScriptDir & "/Data/ProxyChecked.txt")
			EndIf
			ExitLoop
		EndIf
	WEnd
EndFunc
Func SaveProxyChecked()
	Local $FileProxy = ""
	For $x = UBound($PROXY_LIST)-1 to 0 step -1
		$FileProxy &= $PROXY_LIST[$x][0] & ":" & $PROXY_LIST[$x][1]
		If $x <> 0 Then
			$FileProxy &= @CRLF
		EndIf
	Next
	Local $hFileOpen = FileOpen(@ScriptDir & "\Data\ProxyChecked.txt", 130) ; 2+128 =>ghi đè và UTF8
	FileWrite($hFileOpen, $FileProxy)
	FileClose($hFileOpen)
EndFunc

Func Readin_Proxies($Path)
    local $linecount = _FileCountLines($Path)
    ReDim $PROXY_LIST[$linecount][5]
    $filehnd = FileOpen($Path)
    For $x = 0 to $linecount step 1
        $Line = ""
        While 1                 ;Collect the entire line into a variable.
            $character = FileRead( $filehnd, 1)
            if @error = -1 then ExitLoop 2
            if $character = @CR Then ExitLoop
            if $character = @LF then ContinueLoop
            $Line &= $character
        WEnd

        $spl = StringSplit( $Line, ":", 1)
        $PROXY_LIST[$x][0] = $spl[1]
        if $spl[0] >= 2 Then
            $PROXY_LIST[$x][1] = $spl[2]
        Else
            $PROXY_LIST[$x][1] = 80
		  EndIf
    Next
    FileClose( $filehnd)
EndFunc

Func Test_Connection( $arrayslot)
    local $SocketID = ___ASocket()
    ___ASockSelect( $SocketID, $__TCP_WINDOW, 0x401 + $arrayslot,  BitOR( 1, 2, 16, 32))
    GUIRegisterMsg( 0x401 + $arrayslot, "Opensocket_data_" )
    ___ASockConnect( $SocketID, $PROXY_LIST[$arrayslot][0], $PROXY_LIST[$arrayslot][1])
    $PROXY_LIST[$arrayslot][3] = $SocketID
EndFunc

Func Opensocket_data_( $hWnd, $iMsgID, $WParam, $LParam )
    Local $iError = ___HiWord( $LParam )
    Local $iEvent = ___LoWord( $LParam )
    Abs($hWnd)
    local $Array_Slot = $iMsgID-0x0401 ;No more loops to slow down message delievery!
    local $x = $Array_Slot
    Switch $iEvent

        Case 16
            If $iError Then ;OFFLINE
					$PROXY_LIST[$x][4] = 0
            Else ;ONLINE
					 $PROXY_LIST[$x][4] = 1
            EndIf
            ___ASockShutdown($PROXY_LIST[$x][3])
            TCPCloseSocket($PROXY_LIST[$x][2])
    EndSwitch
EndFunc

;==================================================================================================================
;
; UDF BẮT ĐẦU TỪ ĐÂY
;
;==================================================================================================================
Func ___ASocket($iAddressFamily = 2, $iType = 1, $iProtocol = 6)
    If $hWs2_32 = -1 Then $hWs2_32 = DllOpen( "Ws2_32.dll" )
    Local $hSocket = DllCall($hWs2_32, "uint", "socket", "int", $iAddressFamily, "int", $iType, "int", $iProtocol)
    If @error Then
        SetError(1, @error)
        Return -1
    EndIf
    If $hSocket[ 0 ] = -1 Then
        SetError(2, ___WSAGetLastError())
        Return -1
    EndIf
    Return $hSocket[ 0 ]
EndFunc   ;==>_ASocket

Func ___ASockShutdown($hSocket)
    If $hWs2_32 = -1 Then $hWs2_32 = DllOpen( "Ws2_32.dll" )
    Local $iRet = DllCall($hWs2_32, "int", "shutdown", "uint", $hSocket, "int", 2)
    If @error Then
        SetError(1, @error)
        Return False
    EndIf
    If $iRet[ 0 ] <> 0 Then
        SetError(2, ___WSAGetLastError())
        Return False
    EndIf
    Return True
EndFunc   ;==>_ASockShutdown

Func ___ASockClose($hSocket)
    If $hWs2_32 = -1 Then $hWs2_32 = DllOpen( "Ws2_32.dll" )
    Local $iRet = DllCall($hWs2_32, "int", "closesocket", "uint", $hSocket)
    If @error Then
        SetError(1, @error)
        Return False
    EndIf
    If $iRet[ 0 ] <> 0 Then
        SetError(2, ___WSAGetLastError())
        Return False
    EndIf
    Return True
EndFunc   ;==>_ASockClose

Func ___ASockSelect($hSocket, $hWnd, $uiMsg, $iEvent)
    If $hWs2_32 = -1 Then $hWs2_32 = DllOpen( "Ws2_32.dll" )
    Local $iRet = DllCall( _
            $hWs2_32, _
            "int", "WSAAsyncSelect", _
            "uint", $hSocket, _
            "hwnd", $hWnd, _
            "uint", $uiMsg, _
            "int", $iEvent _
            )
    If @error Then
        SetError(1, @error)
        Return False
    EndIf
    If $iRet[ 0 ] <> 0 Then
        SetError(2, ___WSAGetLastError())
        Return False
    EndIf
    Return True
EndFunc   ;==>_ASockSelect

; Note: you can see that $iMaxPending is set to 5 by default.
; IT DOES NOT MEAN THAT DEFAULT = 5 PENDING CONNECTIONS
; 5 == SOMAXCONN, so don't worry be happy
Func ___ASockListen($hSocket, $sIP, $uiPort, $iMaxPending = 5); 5 == SOMAXCONN => No need to change it.
    Local $iRet
    Local $stAddress

    If $hWs2_32 = -1 Then $hWs2_32 = DllOpen( "Ws2_32.dll" )

    $stAddress = ___SockAddr($sIP, $uiPort)
    If @error Then
        SetError(@error, @extended)
        Return False
    EndIf

    $iRet = DllCall($hWs2_32, "int", "bind", "uint", $hSocket, "ptr", DllStructGetPtr($stAddress), "int", DllStructGetSize($stAddress))
    If @error Then
        SetError(3, @error)
        Return False
    EndIf
    If $iRet[ 0 ] <> 0 Then
        $stAddress = 0; Deallocate
        SetError(4, ___WSAGetLastError())
        Return False
    EndIf

    $iRet = DllCall($hWs2_32, "int", "listen", "uint", $hSocket, "int", $iMaxPending)
    If @error Then
        SetError(5, @error)
        Return False
    EndIf
    If $iRet[ 0 ] <> 0 Then
        $stAddress = 0; Deallocate
        SetError(6, ___WSAGetLastError())
        Return False
    EndIf

    Return True
EndFunc   ;==>_ASockListen

Func ___ASockConnect($hSocket, $sIP, $uiPort)
    Local $iRet
    Local $stAddress

    If $hWs2_32 = -1 Then $hWs2_32 = DllOpen( "Ws2_32.dll" )

    $stAddress = ___SockAddr($sIP, $uiPort)
    If @error Then
        SetError(@error, @extended)
        Return False
    EndIf

    $iRet = DllCall($hWs2_32, "int", "connect", "uint", $hSocket, "ptr", DllStructGetPtr($stAddress), "int", DllStructGetSize($stAddress))
    If @error Then
        SetError(3, @error)
        Return False
    EndIf

    $iRet = ___WSAGetLastError()
    If $iRet = 10035 Then; WSAEWOULDBLOCK
        Return True; Asynchronous connect attempt has been started.
    EndIf
    SetExtended(1); Connected immediately
    Return True
EndFunc   ;==>_ASockConnect

; A wrapper function to ease all the pain in creating and filling the sockaddr struct
Func ___SockAddr($sIP, $iPort, $iAddressFamily = 2)
    Local $iRet
    Local $stAddress

    If $hWs2_32 = -1 Then $hWs2_32 = DllOpen( "Ws2_32.dll" )

    $stAddress = DllStructCreate("short; ushort; uint; char[8]")
    If @error Then
        SetError(1, @error)
        Return False
    EndIf

    DllStructSetData($stAddress, 1, $iAddressFamily)
    $iRet = DllCall($hWs2_32, "ushort", "htons", "ushort", $iPort)
    DllStructSetData($stAddress, 2, $iRet[ 0 ])
    $iRet = DllCall($hWs2_32, "uint", "inet_addr", "str", $sIP)
    If $iRet[ 0 ] = 0xffffffff Then; INADDR_NONE
        $stAddress = 0; Deallocate
        SetError(2, ___WSAGetLastError())
        Return False
    EndIf
    DllStructSetData($stAddress, 3, $iRet[ 0 ])

    Return $stAddress
EndFunc   ;==>__SockAddr

Func ___WSAGetLastError()
    If $hWs2_32 = -1 Then $hWs2_32 = DllOpen( "Ws2_32.dll" )
    Local $iRet = DllCall($hWs2_32, "int", "WSAGetLastError")
    If @error Then
        ;if $console_out = True then ConsoleWrite("+> _WSAGetLastError(): WSAGetLastError() failed. Script line number: " & @ScriptLineNumber & @CRLF)
        SetExtended(1)
        Return 0
    EndIf
    Return $iRet[ 0 ]
EndFunc   ;==>_WSAGetLastError
; Got these here:
; http://www.autoitscript.com/forum/index.php?showtopic=5620&hl=MAKELONG
Func ___MakeLong($LoWord, $HiWord)
    Return BitOR($HiWord * 0x10000, BitAND($LoWord, 0xFFFF)); Thanks Larry
EndFunc   ;==>_MakeLong

Func ___HiWord($Long)
    Return BitShift($Long, 16); Thanks Valik
EndFunc   ;==>_HiWord

Func ___LoWord($Long)
    Return BitAND($Long, 0xFFFF); Thanks Valik
EndFunc   ;==>_LoWord
;----------------------------------OTHER AUTOIT INBUILT FUNCS----##

;===============================================================================
;
; Function Name:    _GetIP()
; Description:      Get public IP address of a network/computer.
; Parameter(s):     None
; Requirement(s):   Internet access.
; Return Value(s):  On Success - Returns the public IP Address
;                   On Failure - -1  and sets @ERROR = 1
; Author(s):        Larry/Ezzetabi & Jarvis Stubblefield
;
;===============================================================================
Func _Get_IP()
    Local $ip, $t_ip
    If InetGet("http://checkip.dyndns.org/?rnd1=" & Random(1, 65536) & "&rnd2=" & Random(1, 65536), @TempDir & "\~ip.tmp") Then
        $ip = FileRead(@TempDir & "\~ip.tmp", FileGetSize(@TempDir & "\~ip.tmp"))
        FileDelete(@TempDir & "\~ip.tmp")
        $ip = StringTrimLeft($ip, StringInStr($ip, ":") + 1)
        $ip = StringTrimRight($ip, StringLen($ip) - StringInStr($ip, "/") + 2)
        $t_ip = StringSplit($ip, '.')
        If $t_ip[0] = 4 And StringIsDigit($t_ip[1]) And StringIsDigit($t_ip[2]) And StringIsDigit($t_ip[3]) And StringIsDigit($t_ip[4]) Then
            Return $ip
        EndIf
    EndIf
    If InetGet("http://www.whatismyip.com/?rnd1=" & Random(1, 65536) & "&rnd2=" & Random(1, 65536), @TempDir & "\~ip.tmp") Then
        $ip = FileRead(@TempDir & "\~ip.tmp", FileGetSize(@TempDir & "\~ip.tmp"))
        FileDelete(@TempDir & "\~ip.tmp")
        $ip = StringTrimLeft($ip, StringInStr($ip, "Your ip is") + 10)
        $ip = StringLeft($ip, StringInStr($ip, " ") - 1)
        $ip = StringStripWS($ip, 8)
        $t_ip = StringSplit($ip, '.')
        If $t_ip[0] = 4 And StringIsDigit($t_ip[1]) And StringIsDigit($t_ip[2]) And StringIsDigit($t_ip[3]) And StringIsDigit($t_ip[4]) Then
            Return $ip
        EndIf
    EndIf
    SetError(1)
    Return -1
EndFunc   ;==>_Get_IP