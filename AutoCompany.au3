#RequireAdmin

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <StructureConstants.au3>
#include <GuiListView.au3>
#include <Array.au3>
#include <Date.au3>
#include <Sound.au3>

#include <_HttpRequest.au3>
#include <CheckProxy.au3>
#include <LayDanhSachNgayCapPhep.au3>
#include <HandleImgSearch.au3>
#include <ExcelCOM_UDF.au3>

;BIẾN TOÀN CỤC
Global $iRun = True
Global $iM = True
Global $show = 1 ; ẩn hiện GUI
Global $RoBotStatus = 0
Global $DemSoCompanyRequets = 0
Global $DemSoLanRequets = 0
Global $DemSoCompanyCopied = 0
Global $DemSoCompanyLuuExcel = 0
Global $MaxPage = 0
Global $Handle = ""
Global $ItemProxy = 0
Global $ItemProxyGia = Round(Random(0,50)) ; tăng lần đổi IP [Mang tính chất trang trí]

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
;GIẢ LẬP SỰ KIỆN DOUBLE CLICK
$DoubeClickListView1 = GUICtrlCreateDummy()
$DoubeClickListView2 = GUICtrlCreateDummy()
GUIRegisterMsg($WM_NOTIFY, 'WM_NOTIFY')
;SET HOTKEY
HotKeySet("^s","SaveConfig")
HotKeySet("^d","SetDataInPut")
HotKeySet("^{BS}","DeleteListView")
HotKeySet("^{SPACE}","ExcelFile")
HotKeySet("{F5}","RUNNING")
HotKeySet('{F6}', 'HideGUI')
;LOADING MẶC ĐỊNH
GUICtrlSetState ($InputBD, $GUI_DISABLE)
GUICtrlSetState ($InputKT, $GUI_DISABLE)
GUICtrlSetState ($InputDay, $GUI_ENABLE)
LoadKey()
GUICtrlSetData($InputDay,@MDAY&"/"&@MON&"/"&@YEAR)
GUICtrlSetData($InputBD,@MDAY&"/"&@MON&"/"&@YEAR)
GUICtrlSetData($InputKT,@MDAY&"/"&@MON&"/"&@YEAR)

GUISetState(@SW_SHOW)
#EndRegion

#Region BẮT SỰ KIỆN GUI
	While True
		Switch GUIGetMsg()
			Case $DoubeClickListView1
				;XuLyDoubelCLickListView1()
				Local $DataListView1 = StringMid(GUICtrlRead(GUICtrlRead($ListView1)),1,StringLen(GUICtrlRead(GUICtrlRead($ListView1)))-1)
				GUICtrlCreateListViewItem($DataListView1,$ListView2)
				_GUICtrlListView_DeleteItemsSelected ($ListView1)
				GUICtrlSetData ($Label,"")
			Case $DoubeClickListView2
				;XuLyDoubelCLickListView2()
				Local $DataListView2 = StringMid(GUICtrlRead(GUICtrlRead($ListView2)),1,StringLen(GUICtrlRead(GUICtrlRead($ListView2)))-1)
				GUICtrlCreateListViewItem($DataListView2,$ListView1)
				_GUICtrlListView_DeleteItemsSelected ($ListView2)
				GUICtrlSetData ($Label,"")
			Case $ButtonSave
				SaveConfig()
			Case $PictureRobot
					Local $pos = WinGetPos($Handle)
					If $iM Then
						ToolTip(FileRead(@ScriptDir &"\Data\ThongTinPhanMem.md"),$pos[0]+$pos[2],$pos[1])

						$iM = False
					Else
						$iM = True
						ToolTip("")
					EndIf
			Case $ButtonPlay
				If GUICtrlRead($RadioMacDinh) = 1 Then
					If StringLen(GUICtrlRead($InputDay)) = 10 Then
						RUNNING()
						_SoundPlay(@ScriptDir & "\TienTrinhHoanTat.mp3",1)
					Else
						GUICtrlSetData ($Label,@CRLF &"  Vui lòng điền ngày" & @CRLF & "  Đúng định dạng: dd/mm/yyyy"& @CRLF & "  Ví dụ: 05/05/2020"& @CRLF &"  Để lấy kết quả từ ngày đó đến "& @CRLF &"  ngày hiện tại")
					EndIf
				Else
					If StringLen(GUICtrlRead($InputBD)) = 10 And StringLen(GUICtrlRead($InputKT)) = 10 Then
						Global $NgayBatDauDanhSachNgayCapPhep = GUICtrlRead($InputBD)
						Global $NgayKetThucDanhSachNgayCapPhep = GUICtrlRead($InputKT)
						RUNNINGNANGCAO()
						_SoundPlay(@ScriptDir & "\TienTrinhHoanTat.mp3",1)
					Else
						GUICtrlSetData ($Label,@CRLF &"  Vui lòng điền đủ hai trường" & @CRLF & "  Đúng định dạng: dd/mm/yyyy"& @CRLF & "  Ví dụ: "& @CRLF &"  Bắt đầu :05/05/2020"& @CRLF & "  Kết thúc :05/05/2021"& @CRLF &"  Để lấy kết quả của 365 ngày")
					EndIf
				EndIf
			Case $RadioMacDinh
				If GUICtrlRead($RadioMacDinh) = 1 Then
					GUICtrlSetState ($InputBD, $GUI_DISABLE)
					GUICtrlSetState ($InputKT, $GUI_DISABLE)
					GUICtrlSetState ($InputDay, $GUI_ENABLE)
				Else
					GUICtrlSetState ($InputBD, $GUI_ENABLE)
					GUICtrlSetState ($InputKT, $GUI_ENABLE)
					GUICtrlSetState ($InputDay, $GUI_DISABLE)
				EndIf
			Case $RadioNangCao
				If GUICtrlRead($RadioMacDinh) = 1 Then
					GUICtrlSetState ($InputBD, $GUI_DISABLE)
					GUICtrlSetState ($InputKT, $GUI_DISABLE)
					GUICtrlSetState ($InputDay, $GUI_ENABLE)
				Else
					GUICtrlSetState ($InputBD, $GUI_ENABLE)
					GUICtrlSetState ($InputKT, $GUI_ENABLE)
					GUICtrlSetState ($InputDay, $GUI_DISABLE)
				EndIf
			Case $GUI_EVENT_CLOSE
				Exit
		EndSwitch
	WEnd
#EndRegion
#Region FUNCTION MAIN
Func ON_OFF_PROGRAM()

EndFunc
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
	if $iRun Then
		GUICtrlSetData($ButtonPlay,"Stop")
		$iRun = False
		Local $Key, $String, $DanhsachCongTy
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

			;Kiểm tra và chuẩn bị Proxy trong file ProxyChecked.txt trước khi bắt đầu chạy
			GUICtrlSetData ($Label,"  Đang chuẩn bị IP để Requet...")
			CheckProxy(False)
			For $i = 1 to $DanhsachLink[0]-1
				Local $KetQua[1][6]
				Local $DanhsachCongTy = LayDanhSachCongTyNangCao($DanhsachLink[$i])
				ThongBao()
				For $j = 0 to UBound($DanhsachCongTy) - 1
					Local $ThongTin = LayThongTinCongTy($DanhsachCongTy[$j],True);chạy ở chế độ nâng cao
					ThongBao()
					If KiemTraNgayCapPhep(StringStripWS($ThongTin[5],8)) Then
						_ArrayAdd($KetQua, $ThongTin[1]&"|"&$ThongTin[5]&"|"&$ThongTin[0]&"|"&$ThongTin[2]&"|"&$ThongTin[4]&"|"&$ThongTin[3])
						ThongBao()
					EndIf
				Next
				If (UBound($KetQua)-1)> 0 Then
					CreateExcel($Key[$i],$KetQua,$key,$i,$j)
				Else
					Local $KetQua[1][6]
					CreateExcel($Key[$i],$KetQua,$key,$i,$j)
				EndIf
				MsgBox(0,'',"sắp hiển thị mảng")
				_ArrayDisplay($KetQua)
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
#EndRegion
#Region FUNCTION XỬ LÝ THÔNG TIN CÔNG TY
Func LayDanhSachCongTy($Link)
	Local $rq = _HttpRequest(2, $Link)
			$rq = StringRegExp($rq,'href="(.*?)"',3)
	Local 	$Lenth=UBound($rq)
	For $i = $Lenth -1 to 0 Step -1
		If  StringInStr($rq[$i],"https://www.tratencongty.com/company") == 0 Then
			_ArrayDelete($rq, $i)
		EndIf
	Next
	Local $Lenth=UBound($rq)
	For $i = $Lenth -1 to 0 Step -2
		_ArrayDelete($rq, $i)
	Next
	Return $rq
EndFunc
Func LayThongTinCongTy($Link,$NangCao = False)
	Local $rq="", $TenCongTy="",$DiaChi="", $DaiDienPhapLuat = "", $NgayCapPhep = "", $LinkSoDienThoai="", $MaSoThue="", $SoDienThoai = "", $ThongTin[6] = ["", "", "", "", "", ""]
	If ($NangCao) Then
		$rq = GetHTMLNangCao($link)
		$DemSoCompanyRequets +=1
	Else
		$rq = _HttpRequest(2, $link)
	EndIf
	$MaSoThue = StringRegExp($rq,'<title>(.*?)</title>',3);OK
	; Lấy được chuỗi trong ô Sẫm màu
	$rq = StringMid($rq,StringInStr($rq,'<div class="jumbotron">'),StringInStr($rq,"<div>Doanh nghiệp mới cập nhật:</h4>")-1)
	$TenCongTy = StringRegExp($rq,'<span title="(.*?)"',3) ;OK
	;thu hẹp phạm vi source vì bị trùng
	$rq = StringMid($rq,StringInStr($rq,"Địa chỉ:"),StringLen($rq))
	$DiaChi = StringMid($rq,10,StringInStr($rq,"<br/>")-10)
	$DaiDienPhapLuat = StringRegExp($rq,'Đại diện pháp luật: (.*?)<br/>',3)
	$LinkSoDienThoai = StringRegExp($rq,'<img src="data:image/png;base64,(.*?)"',3)
	$NgayCapPhep = StringRegExp($rq,'Ngày cấp giấy phép: (.*?)<br/>',3)

	If IsArray($TenCongTy) Then
		_ArrayInsert($ThongTin, 0, $TenCongTy[0])
	Else
		_ArrayInsert($ThongTin, 0, " ")
		_HttpRequest_ConsoleWrite("Không tồn tại $TenCongTy : "& $Link)
	EndIf ;OK

	If IsArray($MaSoThue) Then
		_ArrayInsert($ThongTin, 1, $MaSoThue[0])
	Else
		_ArrayInsert($ThongTin, 1, " ")
		_HttpRequest_ConsoleWrite("Không tồn tại $MaSoThue : "& $Link)
	EndIf ;OK

	_ArrayInsert($ThongTin, 2, $DiaChi) ;OK

	If IsArray($DaiDienPhapLuat) Then
		_ArrayInsert($ThongTin, 3, $DaiDienPhapLuat[0])
	Else
		_ArrayInsert($ThongTin, 3, " ")
		_HttpRequest_ConsoleWrite("Không tồn tại $DaiDienPhapLuat : "& $Link)
	EndIf;OK
	If IsArray($LinkSoDienThoai) Then
		If StringLen($LinkSoDienThoai[0]) > 200 Then
			SaveImage("SDT",$LinkSoDienThoai[0])
			$SoDienThoai = GiaiCapCha("SDT")
		Else
			$SoDienThoai = " "
			_HttpRequest_ConsoleWrite("Loi Sai Link $SoDienThoai : "& $Link)
		EndIf
	Else
		$SoDienThoai = " "
		_HttpRequest_ConsoleWrite("Không tồn tại $SoDienThoai : "& $Link)
	EndIf
	_ArrayInsert($ThongTin, 4, $SoDienThoai);OK

	If IsArray($NgayCapPhep) Then
		_ArrayInsert($ThongTin, 5, $NgayCapPhep[0])
	Else
		_ArrayInsert($ThongTin, 5, " ")
		_HttpRequest_ConsoleWrite("Không tồn tại $NgayCapPhep : "& $Link)
	EndIf
	Return $ThongTin
EndFunc
Func LayDanhSachCongTyNangCao($Link)
	$MaxPage = 0
	Local $Dem = 0
	Local $Lenth, $MangKetquaDSCompany, $MaNguonTrang, $LinkCompany, $LinkPage
	;Xử lý lấy phần $MaxPage
	$MaNguonTrang = GetHTMLNangCao($Link)
	$LinkCompany = StringRegExp($MaNguonTrang,'href="(.*?)"',3)
	$LinkPage = $LinkCompany ; Nhân đôi mảng request được
	For $k = UBound($LinkPage) -1 to 0 Step -1 ; xóa các phần tử != chuỗi
		If  StringInStr($LinkPage[$k],$Link&"?page=") == 0 Then
			_ArrayDelete($LinkPage, $k)
		EndIf
	Next
	If IsArray($LinkPage) Then
		$MaxPage = StringReplace($LinkPage[UBound($LinkPage) -1], $Link&"?page=", "")
	EndIf
	;Bắt đầu lấy danh sách link công ty từ trang 1 đến trang $MaxPage
	For $i = 1 to $MaxPage
		$MaNguonTrang = GetHTMLNangCao($Link & "?page=" & $i)
		$LinkCompany = StringRegExp($MaNguonTrang,'href="(.*?)"',3)
		If $i == 1 Then
			$MangKetquaDSCompany = $LinkCompany
		Else
			_ArrayAdd($MangKetquaDSCompany, $LinkCompany)
		EndIf
		If $i == $MaxPage Then
			ExitLoop
		EndIf
		$Dem += 1
		GUICtrlSetData ($Label,"  Đang lấy danh sách công ty. "&@CRLF &"  Đã lấy: "&$Dem &"/"& $MaxPage& " Trang" & @CRLF &"  Đây là quá trình lấy số lượng lớn "& @CRLF &"  Mất rất nhều thời gian"& @CRLF &"  Nhấn F6 để ẩn giao diện này" & @CRLF &"  Tôi sẽ nhắc bạn khi hoàn tất")
	Next
	GUICtrlSetData ($Label,"  Đã lấy danh sách công ty. "&@CRLF &"  Đã lấy: "&$MaxPage &"/"& $MaxPage & @CRLF &"  Đây là quá trình lấy số lượng lớn "& @CRLF &"  Mất rất nhều thời gian"& @CRLF &"  Nhấn F6 để ẩn giao diện này" & @CRLF &"  Tôi sẽ nhắc bạn khi hoàn tất")
	;Loại Bỏ link không phải dạng "https://www.tratencongty.com/company" và trùng nhau
	$Lenth = UBound($MangKetquaDSCompany)
	For $j = $Lenth -1 to 0 Step -1
		If  StringInStr($MangKetquaDSCompany[$j],"https://www.tratencongty.com/company") == 0 Then
			_ArrayDelete($MangKetquaDSCompany, $j)
		EndIf
		GUICtrlSetData ($Label,"  Xử lý danh sách công ty. "&@CRLF &"  Đã xử lý: "&$j &"/"& $Lenth& " URL" & @CRLF &"  Đây là quá trình lấy số lượng lớn "& @CRLF &"  Mất rất nhều thời gian"& @CRLF &"  Nhấn F6 để ẩn giao diện này" & @CRLF &"  Tôi sẽ nhắc bạn khi hoàn tất")
	Next
	Local $Lenth=UBound($MangKetquaDSCompany)
	For $m = $Lenth -1 to 0 Step -2
		_ArrayDelete($MangKetquaDSCompany, $m)
		GUICtrlSetData ($Label,"  Xử lý danh sách công ty. "&@CRLF &"  Đã xử lý: "&$m &"/"& $Lenth& " URL" & @CRLF &"  Đây là quá trình lấy số lượng lớn "& @CRLF &"  Mất rất nhều thời gian"& @CRLF &"  Nhấn F6 để ẩn giao diện này" & @CRLF &"  Tôi sẽ nhắc bạn khi hoàn tất")
	Next
	Return $MangKetquaDSCompany
EndFunc
Func KiemTraNgayCapPhep($NgayCapPhep)
	Global $MangNgayCapPhep
	Global $NgayBatDauDanhSachNgayCapPhep
	Global $NgayKetThucDanhSachNgayCapPhep
	LayDanhSachNgayCapPhep()
	For $i = 0 To UBound($MangNgayCapPhep)-1
		;~ Nếu Replace ra rỗng thì đúng
		If StringReplace($MangNgayCapPhep[$i],$NgayCapPhep,"") == "" Then
			$DemSoCompanyCopied += 1
			Return True
		EndIf
	Next
	Return False
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
Func GiaiCapCha($LoaiCapCha)
	Local $DayToaDoMaSoThue[2][2]
	Local $toado[0][0]
	Local $MangHinhAnhDoiChieu[13] = ['0','1','2','3','4','5','6','7','8','9','+','-','-V2']
	Local $String = ""
	For $k = 0 to UBound($MangHinhAnhDoiChieu)-1
			Local $toado = _BmpImgSearch(@ScriptDir & "\Image\" & $LoaiCapCha & ".bmp", @ScriptDir & "\Image\"&$MangHinhAnhDoiChieu[$k]&".bmp")
			For $i = 1 to $toado[0][0]
				Local $data = $MangHinhAnhDoiChieu[$k] & "|"& Number($toado[$i][0])
				_ArrayAdd($DayToaDoMaSoThue, $data )
			Next
	Next
	;Dư 2 phần tử đầu tiên do lúc đầu khởi tạo [2][2] rỗng
	Local $Lenth = UBound($DayToaDoMaSoThue)
	; Đưa tất cả phần tử dịch lùi 2 đơn vị từ 2=>0
	for $i = 2 to $Lenth - 1
		$DayToaDoMaSoThue[$i-2][0] = String($DayToaDoMaSoThue[$i][0])
		$DayToaDoMaSoThue[$i-2][1] = Number($DayToaDoMaSoThue[$i][1])
	Next
	; Nếu không xóa lùi thì sẽ lỗi
	_ArrayDelete($DayToaDoMaSoThue,$Lenth-1)
	_ArrayDelete($DayToaDoMaSoThue,$Lenth-2)
	_ArraySort($DayToaDoMaSoThue,0,0,0,1)
	For $i = 0 to UBound($DayToaDoMaSoThue)-1
		$String &= $DayToaDoMaSoThue[$i][0]
	Next
	; Nếu SDT có dấu "-" thì cắt đoạn sau dấu "-" đi  ; ví dụ: 0988312732-0933 => 0988312732
	If StringInStr($String,"-") <> 0 Then
		$String = StringMid($String,1,StringInStr($String,"-")-1)
	EndIf
	Return $String
EndFunc
Func CreateExcel($Name,$ThongTin,$Key,$x,$j)
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
Func GetHTMLNangCao($Link)
	$LinkHienTai = $Link
	$ItemProxyGia += 1
	_HttpRequest_ReduceMem() ;giảm tài nguyên hệ thống
	While True
		$SourceHTML = _HttpRequest(2, $Link)
		If CheckBlock($SourceHTML) Then
			While Not(_HttpRequest_CheckProxyLive($PROXY_LIST[$ItemProxy][0]&":"&$PROXY_LIST[$ItemProxy][1]))
				$ItemProxy+=1
			WEnd
			_HttpRequest_SetProxy($PROXY_LIST[$ItemProxy][0]&":"&$PROXY_LIST[$ItemProxy][1])
		Else ;Không Bị Block
			$DemSoLanRequets +=1
			ExitLoop
		EndIf
	WEnd
	return $SourceHTML
EndFunc
#EndRegion
#Region FUNCTION
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
Func ThongBao($Ketthuc = 0)
	If $Ketthuc = 1 Then
		GUICtrlSetData ($Label,"  LẦN ĐỔI IP : " & $ItemProxyGia & @CRLF & "  Lần Requet : " & $DemSoLanRequets& @CRLF &"  Công ty đã yêu cầu : " & $DemSoCompanyRequets& "/"& $MaxPage*50 & @CRLF & "  Đã Copied : " & $DemSoCompanyCopied& @CRLF & "  Đã Lưu vào Excel : " & $DemSoCompanyLuuExcel & @CRLF & "  Tiến trình hoàn tất!!!")
	Else
		GUICtrlSetData ($Label,"  LẦN ĐỔI IP : " & $ItemProxyGia & @CRLF & "  Lần Requet : " & $DemSoLanRequets & @CRLF &"  Công ty đã yêu cầu : " & $DemSoCompanyRequets & "/" & $MaxPage*50 & @CRLF & "  Đã Copied : " & $DemSoCompanyCopied & @CRLF & "  Đã Lưu vào Excel : " & $DemSoCompanyLuuExcel)
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