#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Image\icon.ico
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=Auto lấy thông tin công ty tại                   https://www.thongtincongty.com/
#AutoIt3Wrapper_Res_Fileversion=4.0.0.0
#AutoIt3Wrapper_Res_Language=1066
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
Global $DemSoCompanyRequets = 0
Global $DemSoCompanyCopied = 0
Global $DemSoCompanyLuuExcel = 0
Global $MaxPage = 0
Global $BiChan = "CHƯA BỊ CHẶN"
Global $Handle = ""
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <StructureConstants.au3>
#include <GuiListView.au3>
#include <Array.au3>
#include <_HttpRequest.au3>
#include "LayDanhSachNgayCapPhep.au3"
#include <HandleImgSearch.au3>
#include <ExcelCOM_UDF.au3>
#include <Date.au3>
#Region GIAO DIỆN
$FormMain = GUICreate("AutoCompany", 521, 436, 192, 124)
$Handle = WinGetHandle("AutoCompany")
GUISetIcon(@ScriptDir & "\Image\icon.ico", -1)
GUISetBkColor(0xFFFFFF)
$ListView1 = GUICtrlCreateListView("            DANH SÁCH           ", 0, 0, 154, 435)
GUICtrlSetCursor (-1, 0)
$ListView2 = GUICtrlCreateListView("              ĐÃ CHỌN              ", 155, 0, 154, 435-150)
GUICtrlSetCursor (-1, 0)

$InputDaTa = GUICtrlCreateInput("", 310, 1, 130, 28, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER,$WS_BORDER), 0)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
$ButtonSave = GUICtrlCreateButton("Save", 440, 0, 80, 30)
GUICtrlSetCursor (-1, 0)
GUICtrlSetFont(-1, 12, 400, 0, "Tahoma")
$Label1 = GUICtrlCreateLabel("Xem hướng dẫn (Ctrl + M)", 311, 413-20, 1000, 17)
GUICtrlSetColor(-1, 0xFF0000)
$Picture = GUICtrlCreatePic(@ScriptDir & "\Image\Image.bmp", 311, 413, 207, 22)
$Label = GUICtrlCreateLabel("", 310, 30, 209, 145, $WS_BORDER)
GUICtrlSetCursor (-1, 7)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
$ButtonPlay = GUICtrlCreateButton("Play", 440, 175, 80, 30)
GUICtrlSetCursor (-1, 0)
GUICtrlSetFont(-1, 12, 400, 0, "Tahoma")

$Label1 = GUICtrlCreateLabel("-----------------------------------------------", 309, 198-2, 131, 10)
$RadioMacDinh = GUICtrlCreateRadio("Mặc Đinh", 313, 208, 100, 17)
GUICtrlSetCursor (-1, 0)
GUICtrlSetState(-1, $GUI_CHECKED)
$Label1 = GUICtrlCreateLabel("[Ngày Cấp Phép => Hiện Tại] :", 312+20, 246-19, 148, 17)
$InputDay = GUICtrlCreateInput("", 312+20, 246, 128, 28, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER,$WS_BORDER), 0)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")

$Label1 = GUICtrlCreateLabel("------------------------------------------------------------------------------------", 309, 208+70-1, 1000, 17)
$RadioNangCao = GUICtrlCreateRadio("Nâng Cao", 313, 208+83, 100, 17)
GUICtrlSetCursor (-1, 0)
$Label1 = GUICtrlCreateLabel("Ngày Bắt Đầu :", 312+20, 208+83+17, 128, 17)
$InputBD = GUICtrlCreateInput("", 312+20, 208+83+17+17, 124, 24, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER,$WS_BORDER), 0)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
$Label1 = GUICtrlCreateLabel("Ngày Kết Thúc :", 312+20, 208+83+17+17+24, 128, 17)
$InputKT = GUICtrlCreateInput("", 312+20, 208+83+17+17+24+17, 124, 24, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER,$WS_BORDER), 0)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")

$PictureRobot = GUICtrlCreatePic(@ScriptDir & "\Image\RoBotXanh.bmp", 155, 280-2, 154, 157,BitOR($SS_NOTIFY,$WS_GROUP,$WS_CLIPSIBLINGS,$WS_BORDER))
GUICtrlSetCursor (-1, 0)
; SET mặc định hiển thị, mặc định không hiển thị
GUICtrlSetState ($InputBD, $GUI_DISABLE)
GUICtrlSetState ($InputKT, $GUI_DISABLE)
GUICtrlSetState ($InputDay, $GUI_ENABLE)
;======================HOT KEY
HotKeySet("^s","SaveConfig")
HotKeySet("^d","SetDataInPut")
HotKeySet("^m","Tooltips")
HotKeySet("^{BS}","DeleteListView")
HotKeySet("^{PAUSE}","Thoat")
HotKeySet("^{SPACE}","ExcelFile")
HotKeySet("{F5}","RUNNING")
HotKeySet('{F6}', 'HideGUI')
Global $iRun = True
Global $iM = True
Global $show = 1 ; ẩn hiện GUI
Global $RoBotStatus = 0
;====================== GIẢ LẬP SỰ KIỆN DOUBLE CLICK
$DoubeClickListView1 = GUICtrlCreateDummy()
$DoubeClickListView2 = GUICtrlCreateDummy()
GUIRegisterMsg($WM_NOTIFY, 'WM_NOTIFY')
;=======================
LoadKey()
GUICtrlSetData($InputDay,@MDAY&"/"&@MON&"/"&@YEAR)
GUICtrlSetData($InputBD,@MDAY&"/"&@MON&"/"&@YEAR)
GUICtrlSetData($InputKT,@MDAY&"/"&@MON&"/"&@YEAR)
;=======================
GUISetState(@SW_SHOW)
#EndRegion
#Region SỰ KIỆN
While True
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonPlay
			If GUICtrlRead($RadioMacDinh) = 1 Then ; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
				If StringLen(GUICtrlRead($InputDay)) = 10 Then
					RUNNING()
				Else
					GUICtrlSetData ($Label,@CRLF &"  Vui lòng điền ngày" & @CRLF & "  Đúng định dạng: dd/mm/yyyy"& @CRLF & "  Ví dụ: 05/05/2020"& @CRLF &"  Để lấy kết quả từ ngày đó đến "& @CRLF &"  ngày hiện tại")
				EndIf

			Else
				If StringLen(GUICtrlRead($InputBD)) = 10 And StringLen(GUICtrlRead($InputKT)) = 10 Then
					Global $NgayBatDauDanhSachNgayCapPhep = GUICtrlRead($InputBD)
					Global $NgayKetThucDanhSachNgayCapPhep = GUICtrlRead($InputKT)
					RUNNINGNANGCAO()
				Else
					GUICtrlSetData ($Label,@CRLF &"  Vui lòng điền đủ hai trường" & @CRLF & "  Đúng định dạng: dd/mm/yyyy"& @CRLF & "  Ví dụ: "& @CRLF &"  Bắt đầu :05/05/2020"& @CRLF & "  Kết thúc :05/05/2021"& @CRLF &"  Để lấy kết quả của 365 ngày")
				EndIf
			EndIf
		Case $ButtonSave
			SaveConfig()
		Case $DoubeClickListView1
			XuLyDoubelCLickListView1()
		Case $DoubeClickListView2
			XuLyDoubelCLickListView2()
		Case $RadioMacDinh
			If GUICtrlRead($RadioMacDinh) = 1 Then ; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
				GUICtrlSetState ($InputBD, $GUI_DISABLE)
				GUICtrlSetState ($InputKT, $GUI_DISABLE)
				GUICtrlSetState ($InputDay, $GUI_ENABLE)
			Else
				GUICtrlSetState ($InputBD, $GUI_ENABLE)
				GUICtrlSetState ($InputKT, $GUI_ENABLE)
				GUICtrlSetState ($InputDay, $GUI_DISABLE)
			EndIf
		Case $RadioNangCao
			If GUICtrlRead($RadioMacDinh) = 1 Then ; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
				GUICtrlSetState ($InputBD, $GUI_DISABLE)
				GUICtrlSetState ($InputKT, $GUI_DISABLE)
				GUICtrlSetState ($InputDay, $GUI_ENABLE)
			Else
				GUICtrlSetState ($InputBD, $GUI_ENABLE)
				GUICtrlSetState ($InputKT, $GUI_ENABLE)
				GUICtrlSetState ($InputDay, $GUI_DISABLE)
			EndIf
		case $PictureRobot
			Tooltips()
	EndSwitch
WEnd
#EndRegion
#Region FUNCTION
Func RUNNING();lấy một trang
	if $iRun Then
		GUICtrlSetData($ButtonPlay,"Stop")
		$iRun = False
		Local $LenthDay = _DateDiff ( "D", DaoNguocNgay(GUICtrlRead($InputDay)), @YEAR&"/"&@MON&"/"&@MDAY)
		Local $Key
		Local $String,$DanhsachCongTy
		$list_count = _GUICtrlListView_GetItemCount ($ListView2)
		If $list_count < 1 Then
			GUICtrlSetData ($Label,@CRLF &"  Vui lòng chọn !!!")
		Else
			For $i=0 To $list_count-1
				_GUICtrlListView_SetItemSelected($ListView2,$i)
				$Read = GUICtrlRead(GUICtrlRead($ListView2))
				$String &= $Read
			Next
			$Key = StringSplit($String,"|")
			Local $Lenth = UBound($Key)
			$Read = ""
			For $i = 1 to $Lenth-2
				$Read &= IniRead(@ScriptDir & "\Data\Config.ini","Data",$Key[$i],"") & "|"
			Next
			Local $DanhsachLink = StringSplit($Read,"|")
			;==================
			Local $Lenth = $DanhsachLink[0]-1
			For $i = 1 to $Lenth ;;; LINK TONG
				Local $KetQua[1][6]
				GUICtrlSetData ($Label,@CRLF &"  Đang xử lý... ["&$Key[$i]&"]")
				Local $DanhsachCongTy = LayDanhSachCongTy($DanhsachLink[$i])
				Local $Len = UBound($DanhsachCongTy)
				For $j = 0 to $Len - 1
					Local $ThongTin = LayThongTinCongTy($DanhsachCongTy[$j])
					If _DateDiff("D",DaoNguocNgay(StringStripWS($ThongTin[5],8)),@YEAR&"/"&@MON&"/"&@MDAY) <= $LenthDay Then
						_ArrayAdd($KetQua, $ThongTin[1]&"|"&$ThongTin[5]&"|"&$ThongTin[0]&"|"&$ThongTin[2]&"|"&$ThongTin[4]&"|"&$ThongTin[3])
						GUICtrlSetData ($Label,@CRLF &"  Đã xử lý... ["&$Key[$i]&"]" & @CRLF & "  Đã Copy "&$j &" Company.")
					Else
						ExitLoop
					EndIf
				Next
				If (UBound($KetQua)-1)>0 Then
					CreateExcel($Key[$i],$KetQua,$key,$i,$j)
				Else
					Local $KetQua[1][6]
					CreateExcel($Key[$i],$KetQua,$key,$i,$j)
				EndIf
			Next
			GUICtrlSetData ($Label,@CRLF &"  Tiến trình đã hoàn tất !!!")
		EndIf
		$iRun = True
		GUICtrlSetData($ButtonPlay,"Play")
	EndIf
EndFunc
Func RUNNINGNANGCAO();Lấy tất cả
	Global $DemSoCompanyRequets = 0
	Global $DemSoCompanyCopied = 0
	Global $DemSoCompanyLuuExcel = 0
	Global $BiChan = "CHƯA"
	if $iRun Then
		GUICtrlSetData($ButtonPlay,"Stop")
		$iRun = False
		Local $Key
		Local $String,$DanhsachCongTy
		$list_count = _GUICtrlListView_GetItemCount ($ListView2)
		If $list_count < 1 Then
			GUICtrlSetData ($Label,@CRLF &"  Vui lòng chọn !!!")
		Else
			For $i=0 To $list_count-1
				_GUICtrlListView_SetItemSelected($ListView2,$i)
				$Read = GUICtrlRead(GUICtrlRead($ListView2))
				$String &= $Read
			Next
			$Key = StringSplit($String,"|")
			Local $Lenth = UBound($Key)
			$Read = ""
			For $i = 1 to $Lenth-2
				$Read &= IniRead(@ScriptDir & "\Data\Config.ini","Data",$Key[$i],"") & "|"
			Next
			Local $DanhsachLink = StringSplit($Read,"|")
			;==================
			Local $Lenth = $DanhsachLink[0]-1
			For $i = 1 to $Lenth ;;; LINK TONG
				Local $KetQua[1][6]
				ThongBao()
				Local $DanhsachCongTy = LayDanhSachCongTyNangCao($DanhsachLink[$i])
				Local $Len = UBound($DanhsachCongTy)
				For $j = 0 to $Len - 1
					Local $ThongTin = LayThongTinCongTyNangCao($DanhsachCongTy[$j])
					$DemSoCompanyRequets += 1
						If $ThongTin == 1 Then
							$BiChan = "RỒI"
							ExitLoop
						EndIf
					ThongBao()
					If KiemTraNgayCapPhep(StringStripWS($ThongTin[5],8)) Then
						_ArrayAdd($KetQua, $ThongTin[1]&"|"&$ThongTin[5]&"|"&$ThongTin[0]&"|"&$ThongTin[2]&"|"&$ThongTin[4]&"|"&$ThongTin[3])
						$DemSoCompanyCopied += 1
						ThongBao()
					Else
						ExitLoop
					EndIf
				Next
				If (UBound($KetQua)-1)> 0 Then
					CreateExcel($Key[$i],$KetQua,$key,$i,$j)
				Else
					Local $KetQua[1][6]
					CreateExcel($Key[$i],$KetQua,$key,$i,$j)
				EndIf
			Next
			If GUICtrlRead($RadioMacDinh) = 1 Then
				GUICtrlSetData ($Label,@CRLF &"  Tiến trình đã hoàn tất !!!")
			Else
				ThongBao(1)
			EndIf
		EndIf
		$iRun = True
		GUICtrlSetData($ButtonPlay,"Play")
	EndIf
EndFunc
Func DeleteListView()
	If $iRun Then
	Local $DataListView1 = StringMid(GUICtrlRead(GUICtrlRead($ListView1)),1,StringLen(GUICtrlRead(GUICtrlRead($ListView1)))-1)
	Local $DataKey = FileReadToArray(@ScriptDir & "\Data\Key.txt")
	Local $Lenth = UBound($DataKey)-1
	For $i = $Lenth to 0 step -1
		If $DataKey[$i] == $DataListView1 Then
			_ArrayDelete($DataKey, $i)
		EndIf
	Next
	Local $String = ""
	Local $str
	For $i = 0 to UBound($DataKey)-1
		$String &= $DataKey[$i] & @CRLF
	Next
	FileWrite(FileOpen(@ScriptDir & "\Data\Key.txt",2),$String)
	IniDelete(@ScriptDir & "\Data\Config.ini","Data",$DataListView1)
	LoadKey()
	GUICtrlSetData ($Label,@CRLF &"  Đã xóa ["&$DataListView1&"]")
	EndIf
EndFunc
Func XuLyDoubelCLickListView1()
	Local $DataListView1 = StringMid(GUICtrlRead(GUICtrlRead($ListView1)),1,StringLen(GUICtrlRead(GUICtrlRead($ListView1)))-1)
	GUICtrlCreateListViewItem($DataListView1,$ListView2)
	_GUICtrlListView_DeleteItemsSelected ($ListView1)
	GUICtrlSetData ($Label,"")
EndFunc
Func XuLyDoubelCLickListView2()
	Local $DataListView2 = StringMid(GUICtrlRead(GUICtrlRead($ListView2)),1,StringLen(GUICtrlRead(GUICtrlRead($ListView2)))-1)
	GUICtrlCreateListViewItem($DataListView2,$ListView1)
	_GUICtrlListView_DeleteItemsSelected ($ListView2)
	GUICtrlSetData ($Label,"")
EndFunc
Func WM_NOTIFY($hWnd, $MsgID, $wParam, $lParam)
	Local $stNMHDR, $iCode
	$stNMHDR = DllStructCreate($tagNMHDR, $lParam)
	If @error Then Return 0
	$iCode = DllStructGetData($stNMHDR, 'Code')
	Switch $wParam
		Case $ListView1
			Switch $iCode
				Case $NM_DBLCLK
					$bDblClck_Event = True
					GUICtrlSendToDummy($DoubeClickListView1) ; double
			EndSwitch
		Case $ListView2
			Switch $iCode
				Case $NM_DBLCLK
					$bDblClck_Event = True
					GUICtrlSendToDummy($DoubeClickListView2) ; double
			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc
Func LoadKey()
	_GUICtrlListView_DeleteAllItems($ListView1)
	_GUICtrlListView_DeleteAllItems($ListView2)
	Local $String = FileReadToArray(@ScriptDir & "\Data\Key.txt")
	Local $Lenth = UBound($String)-1
	For $i = 0 to $Lenth
		GUICtrlCreateListViewItem($String[$i],$ListView1)
	Next
EndFunc
Func SetDataInPut()
	If $iRun Then
	GUICtrlSetData($InputDaTa,"")
	GUICtrlSetData($InputDaTa,ClipGet())
	SaveConfig()
	EndIf
EndFunc
Func ExcelFile()
	If $iRun Then
	ShellExecute (@ScriptDir & "\Excel\Data.xlsx")
	EndIf
EndFunc
Func SaveConfig()
	If $iRun Then
		Local $Key = XuLyKey(GUICtrlRead($InputDaTa))
		If StringLen($Key) > 5 Then
			If Not StringInStr(FileRead(@ScriptDir & "\Data\Key.txt"),$Key) Then
				FileWrite(@ScriptDir & "\Data\Key.txt",$Key & @CRLF)
				GUICtrlSetData ($Label,@CRLF &"                   Đã Thêm.")
			Else
				IniDelete(@ScriptDir & "\Data\Config.ini","Data",$Key)
				GUICtrlSetData ($Label,@CRLF &"  Đã Cập Nhật Thành công!")
			EndIf
			IniWrite(@ScriptDir & "\Data\Config.ini","Data",$Key,GUICtrlRead($InputDaTa))
		Else
			GUICtrlSetData ($Label,"")
		EndIf
		GUICtrlSetData($InputDaTa,"")
		LoadKey()
	EndIf
EndFunc
Func XuLyKey($String)
	$String = StringSplit($String,"/")
	Return StringReplace(StringUpper($String[$String[0]-1]),"-"," ")
EndFunc
Func DaoNguocNgay($String)
	$String = StringSplit($String,"/")
	Return $String[3]&"/"&$String[2]&"/"&$String[1]
EndFunc
Func Thoat()
	Exit
EndFunc
Func Tooltips()
	Local $pos = WinGetPos($Handle)
	If $iM Then
		Local $Data = "AutoCompany là phần mềm đơn giản, mục đích: Auto lấy dữ liệu từ http://thongtincongty.com về lưu vào Excel"&@CRLF&"Phần mềm có 2 chức năng chính: "&@CRLF&"1.Mặc Định : Lấy dữ liệu của tỉnh ở trang đầu tiên và có lọc theo ngày cấp phép"&@CRLF&"2.Nâng Cao: Lấy toàn bộ dữ liệu của tỉnh, có lọc theo ngày cấp phép"&@CRLF&""&@CRLF&""&@CRLF&"Các phím tắt: "&@CRLF&"Ctrl + S : Lưu Link vào [danh sách] và đặt Form về trạng thái ban đầu"&@CRLF&"Ctrl + D : Paste Link vào TextBox và (Ctrl + S)"&@CRLF&"Ctrl + BackSpase : Delete Link đã {CHỌN} trong [Danh Sách], chỉ xóa trong [Danh Sách] không xóa được trong [Đã chọn]"&@CRLF&"Ctrl + Spase(Phím cách) : Mở File Excel chứa dữ liệu lấy được"&@CRLF&"Ctrl + M : Bật/Tắt hướng dẫn (Đây chính là hướng dẫn)"&@CRLF&"F5 : Chạy Chương trình"&@CRLF&"F6 : Ẩn/hiện giao diện"&@CRLF&"Ctrl + Pause Break : Thoát"&@CRLF&"Lưu ý :"&@CRLF&" 1. Khi đã F5 [Chạy chương trình], tất cả các thao tác lên Tool này đều bị vô hiệu hóa"&@CRLF&""&@CRLF&""&@CRLF&"Application: AutoCompany"&@CRLF&"Version: 4.0"&@CRLF&"Author: Hiếu EM"&@CRLF&"FB: https://www.facebook.com/hieuemmm"&@CRLF&"Zalo: 0398503361"&@CRLF&"GitHub: github.com/hieuemmm"&@CRLF&"Copyright: TP.Đà Nẵng 10/05/2021 04:15 PM"
		ToolTip(FileRead(@ScriptDir &"\Data\ThongTinPhanMem.md"),$pos[0]+$pos[2],$pos[1])

		$iM = False
	Else
		$iM = True
		ToolTip("")
	EndIf
EndFunc
Func ThongBao($Ketthuc = 0)
	Global $DemSoCompanyRequets
	Global $DemSoCompanyCopied
	Global $DemSoCompanyLuuExcel
	Global $MaxPage
	Global $BiChan
	If $Ketthuc = 1 Then
		GUICtrlSetData ($Label,"  TRANG WEB CHẶN : "& $BiChan &@CRLF & "  Đã Requet : " & $DemSoCompanyRequets& "/"& $MaxPage*50 & @CRLF & "  Đã Copied : " & $DemSoCompanyCopied& @CRLF & "  Đã Lưu vào Excel : " & $DemSoCompanyLuuExcel& @CRLF & "  Tiến trình hoàn tất!!!")
	Else
		GUICtrlSetData ($Label,"  TRANG WEB CHẶN : "& $BiChan &@CRLF & "  Đã Requet : " & $DemSoCompanyRequets& "/"& $MaxPage*50 & @CRLF & "  Đã Copied : " & $DemSoCompanyCopied& @CRLF & "  Đã Lưu vào Excel : " & $DemSoCompanyLuuExcel)
	EndIf
EndFunc
Func HideGUI()
	 If $show = 1 Then
        WinSetState($FormMain, '', @SW_HIDE)
        $show = 0
    Else ;If $show = 0 Then
        WinSetState($FormMain, '', @SW_SHOW )
        $show = 1
    EndIf
EndFunc
#EndRegion
#Region FUNCTION CHÍNH
Func LayDanhSachCongTy($Link)
	Local 	$rq = _HttpRequest(2, $Link)
			$rq = StringRegExp($rq,'href="(.*?)"',3)
	Local 	$Lenth=UBound($rq)
	For $i = $Lenth -1 to 0 Step -1
		If  StringInStr($rq[$i],"https://www.thongtincongty.com/company") == 0 Then
			_ArrayDelete($rq, $i)
		EndIf
	Next
	Local $Lenth=UBound($rq)
	For $i = $Lenth -1 to 0 Step -2
		_ArrayDelete($rq, $i)
	Next
	Return $rq
EndFunc
Func LayThongTinCongTy($Link)
	Local $rq, $TenCongTy, $LinkMaSoThue, $DiaChi, $DaiDienPhapLuat = "", $NgayCapPhep = "", $LinkSoDienThoai, $MaSoThue, $SoDienThoai = "", $ThongTin[6] = ["", "", "", "", "", ""]
	$rq = _HttpRequest(2,$Link)
	$rq = StringMid($rq,StringInStr($rq,"<h4>"),StringInStr($rq,"<div>Doanh nghiệp mới cập nhật:</h4>")-1)
	$rq = StringMid($rq,1,StringInStr($rq,"</div>")-1) ; Lấy được chuỗi trong ô Sẫm màu
	$TenCongTy = StringRegExp($rq,'<span title="(.*?)"',3)
	$LinkMaSoThue = StringRegExp($rq,'<img src="data:image/png;base64,(.*?)"',3)
	$rq = StringMid($rq,StringInStr($rq,"Địa chỉ:"),StringLen($rq))
	$DiaChi = StringMid($rq,10,StringInStr($rq,"<br/>")-10)
	$DaiDienPhapLuat = StringRegExp($rq,'Đại diện pháp luật: (.*?)<br/>',3)
	$LinkSoDienThoai = StringRegExp($rq,'<img src="data:image/png;base64,(.*?)"',3)
	$NgayCapPhep = StringRegExp($rq,'Ngày cấp giấy phép: (.*?)<br/>',3)
	SaveImage("MST",$LinkMaSoThue[0])
	If IsArray($LinkSoDienThoai) Then
		SaveImage("SDT",$LinkSoDienThoai[0])
		$SoDienThoai = GiaiCapCha("SDT")
	EndIf
	$MaSoThue = GiaiCapCha("MST")
	_ArrayInsert($ThongTin, 0, $TenCongTy[0])
	_ArrayInsert($ThongTin, 1, $MaSoThue)
	_ArrayInsert($ThongTin, 2, $DiaChi)
	If IsArray($DaiDienPhapLuat) Then
		_ArrayInsert($ThongTin, 3, $DaiDienPhapLuat[0])
	EndIf
	_ArrayInsert($ThongTin, 4, $SoDienThoai)
	If IsArray($NgayCapPhep) Then
		_ArrayInsert($ThongTin, 5, $NgayCapPhep[0])
	EndIf
	Return $ThongTin
EndFunc
Func LayDanhSachCongTyNangCao($Link,$Sleep = 200)
	Global $BiChan
	Global $MaxPage = 0
	Local $Lenth, $MangKetquaDSCompany, $MaNguonTrang, $LinkCompany, $LinkPage
	#Region Xử lý lấy phần $MaxPage
		$MaNguonTrang = _HttpRequest(2, $Link)
		If CheckBlock($MaNguonTrang) Then
			$BiChan = "RỒI"
		EndIf
		$LinkCompany = StringRegExp($MaNguonTrang,'href="(.*?)"',3)
		$LinkPage = $LinkCompany ; Nhân đôi mảng request được
		For $k = UBound($LinkPage) -1 to 0 Step -1 ; xóa các phần tử != chuỗi
			If  StringInStr($LinkPage[$k],"https://www.thongtincongty.com/tinh-binh-dinh/?page=") == 0 Then
				_ArrayDelete($LinkPage, $k)
			EndIf
		Next
		If IsArray($LinkPage) Then
			Local $MaxPage = StringReplace($LinkPage[UBound($LinkPage) -1], "https://www.thongtincongty.com/tinh-binh-dinh/?page=", "")
		EndIf

	#EndRegion
	#Region Bắt đầu lấy danh sách công ty từ trang 1 đến trang $MaxPage
		For $i = 1 to $MaxPage
			SetError(0)
			$MaNguonTrang = _HttpRequest(2, $Link & "?page=" & $i)
			If CheckBlock($MaNguonTrang) Then
				$BiChan = "RỒI"
				Return $MangKetquaDSCompany
			EndIf
			$LinkCompany = StringRegExp($MaNguonTrang,'href="(.*?)"',3)
			If $i == 1 Then
				$MangKetquaDSCompany = $LinkCompany
			Else
				_ArrayAdd($MangKetquaDSCompany, $LinkCompany)
			EndIf
			Sleep($Sleep)
		Next
		#Region Loại Bỏ link không phải dạng "https://www.thongtincongty.com/company" và trùng nhau
			$Lenth = UBound($MangKetquaDSCompany)
			For $j = $Lenth -1 to 0 Step -1
				If  StringInStr($MangKetquaDSCompany[$j],"https://www.thongtincongty.com/company") == 0 Then
					_ArrayDelete($MangKetquaDSCompany, $j)
				EndIf
			Next
			Local $Lenth=UBound($MangKetquaDSCompany)
			For $m = $Lenth -1 to 0 Step -2
				_ArrayDelete($MangKetquaDSCompany, $m)
			Next
		#EndRegion
	#EndRegion
	Return $MangKetquaDSCompany
EndFunc
Func LayThongTinCongTyNangCao($Link,$Sleep = 200)
	Global $BiChan
	Local $rq, $TenCongTy, $LinkMaSoThue, $DiaChi, $DaiDienPhapLuat = "", $NgayCapPhep = "", $LinkSoDienThoai, $MaSoThue, $SoDienThoai = "", $ThongTin[6] = ["", "", "", "", "", ""]
	$rq = _HttpRequest(2,$Link)
	If CheckBlock($rq) Then
		$BiChan = "RỒI"
		Return 1
	EndIf
	$rq = StringMid($rq,StringInStr($rq,"<h4>"),StringInStr($rq,"<div>Doanh nghiệp mới cập nhật:</h4>")-1)
	$rq = StringMid($rq,1,StringInStr($rq,"</div>")-1) ; Lấy được chuỗi trong ô Sẫm màu
	$TenCongTy = StringRegExp($rq,'<span title="(.*?)"',3)
	$LinkMaSoThue = StringRegExp($rq,'<img src="data:image/png;base64,(.*?)"',3)
	$rq = StringMid($rq,StringInStr($rq,"Địa chỉ:"),StringLen($rq))
	$DiaChi = StringMid($rq,10,StringInStr($rq,"<br/>")-10)
	$DaiDienPhapLuat = StringRegExp($rq,'Đại diện pháp luật: (.*?)<br/>',3)
	$LinkSoDienThoai = StringRegExp($rq,'<img src="data:image/png;base64,(.*?)"',3)
	$NgayCapPhep = StringRegExp($rq,'Ngày cấp giấy phép: (.*?)<br/>',3)
	SaveImage("MST",$LinkMaSoThue[0])
	If IsArray($LinkSoDienThoai) Then
		SaveImage("SDT",$LinkSoDienThoai[0])
		$SoDienThoai = GiaiCapCha("SDT")
	EndIf
	$MaSoThue = GiaiCapCha("MST")
	_ArrayInsert($ThongTin, 0, $TenCongTy[0])
	_ArrayInsert($ThongTin, 1, $MaSoThue)
	_ArrayInsert($ThongTin, 2, $DiaChi)
	If IsArray($DaiDienPhapLuat) Then
		_ArrayInsert($ThongTin, 3, $DaiDienPhapLuat[0])
	EndIf
	_ArrayInsert($ThongTin, 4, $SoDienThoai)
	If IsArray($NgayCapPhep) Then
		_ArrayInsert($ThongTin, 5, $NgayCapPhep[0])
	EndIf
	Sleep($Sleep)
	Return $ThongTin
EndFunc
Func KiemTraNgayCapPhep($NgayCapPhep)
	Global $MangNgayCapPhep
	Global $NgayBatDauDanhSachNgayCapPhep
	Global $NgayKetThucDanhSachNgayCapPhep
	RUNDANHSACHNGAYCAPPHEP()
	For $i = 0 To UBound($MangNgayCapPhep)
		If StringReplace($MangNgayCapPhep[$i],$NgayCapPhep,"") == "" Then
			Return 1
		EndIf
	Next
	Return 0
EndFunc
Func CheckBlock($String)
	If $String == "Cannot Connect To MySQL Server" Then
		Return True
	EndIf
	 Return False
EndFunc
Func SaveImage($FileName,$Base64)
	Local $dBinary = Binary(_B64Decode($Base64))
	Local $hFile = FileOpen(@ScriptDir & "\Image\" & $FileName &".bmp", 16 + 2)
	FileWrite($hFile, $dBinary)
	FileClose($hFile)
EndFunc
Func GiaiCapCha($y)
	Local $DayToaDoMaSoThue[2][2]
	Local $toado[0][0]
	;SO 0
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\0.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "0" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;SO 1
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\1.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "1" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;SO 2
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\2.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "2" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;SO 3
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\3.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "3" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;SO 4
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\4.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "4" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;SO 5
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\5.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "5" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;SO 6
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\6.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "6" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;SO 7
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\7.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "7" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;SO 8
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\8.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "8" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;SO 9
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\9.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "9" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;Dau +
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\+.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "+" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;Dau -
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\-.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "-" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	;Dau - loại 2
	Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $y & ".bmp", @ScriptDir & "\Image\-1.bmp")
	For $i = 1 to $toado[0][0]
		Local $data = "-" & "|"& Number($toado[$i][0])
		_ArrayAdd($DayToaDoMaSoThue, $data )
	Next
	Local $Lenth = UBound($DayToaDoMaSoThue)
	for $i = 2 to $Lenth-1
		$DayToaDoMaSoThue[$i-2][0] = String($DayToaDoMaSoThue[$i][0])
		$DayToaDoMaSoThue[$i-2][1] = Number($DayToaDoMaSoThue[$i][1])
	Next
	_ArrayDelete($DayToaDoMaSoThue,$Lenth-1)
	_ArrayDelete($DayToaDoMaSoThue,$Lenth-2)
	_ArraySort($DayToaDoMaSoThue,0,0,0,1)
	Local $String= ""
	$Lenth = UBound($DayToaDoMaSoThue)
	for $i = 0 to $Lenth-1
		$String &= $DayToaDoMaSoThue[$i][0]
	Next
	$KiemtraTRU = 0 ; nếu sdt có dấu "-" thì cắt đoạn sau dấu "-" đi
	Local $KiemtraTRU = StringInStr($String,"-")
	If $KiemtraTRU <> 0 Then
		$String = StringMid($String,1,$KiemtraTRU-1)
	EndIf
	Return $String
EndFunc
Func CreateExcel($Name,$ThongTin,$Key,$x,$j)
	Global $DemSoCompanyRequets
	Global $DemSoCompanyCopied
	Global $DemSoCompanyLuuExcel
	If GUICtrlRead($RadioMacDinh) = 1 Then
		GUICtrlSetData ($Label,@CRLF &"  Đã xử lý... ["&$Key[$x]&"]" & @CRLF & "  Đã Copy "&$j &" Company."& @CRLF & "  Đang Lưu vào Excel...")
	EndIf
	Local $Lenth = UBound($ThongTin)
	Local $oWorkbook = _ExcelBookOpen(@ScriptDir & "\Excel\Data.xlsx",0)
	_ExcelSheetDelete($oWorkbook, $Name)
	_ExcelSheetAddNew($oWorkbook,$Name)
	_ExcelSheetActivate($oWorkbook, $Name)
	For $i = 1 to $Lenth-1
		_ExcelWriteCell($oWorkbook, $ThongTin[$i][0], $i,1)
		_ExcelWriteCell($oWorkbook, $ThongTin[$i][1], $i,2)
		_ExcelWriteCell($oWorkbook, $ThongTin[$i][2], $i,3)
		_ExcelWriteCell($oWorkbook, $ThongTin[$i][3], $i,4)
		_ExcelWriteCell($oWorkbook, $ThongTin[$i][4], $i,5)
		_ExcelWriteCell($oWorkbook, $ThongTin[$i][5], $i,6)
		$DemSoCompanyLuuExcel += 1
		If GUICtrlRead($RadioMacDinh) = 1 Then
			GUICtrlSetData ($Label,@CRLF &"  Đã xử lý... ["&$Key[$x]&"]" & @CRLF & "  Đã Copy "&$j &" Company."& @CRLF & "  Đã lưu: " & $i)
		Else
			ThongBao()
		EndIf
	Next
	_ExcelBookSave($oWorkbook)
	_ExcelBookClose($oWorkbook)
EndFunc
#EndRegion