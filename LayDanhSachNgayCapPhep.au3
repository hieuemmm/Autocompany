#include <Array.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>

Global $NgayBatDauDanhSachNgayCapPhep = "01/01/2021"
Global $NgayKetThucDanhSachNgayCapPhep = "31/12/2021"
Global $MangNgayCapPhep[1] = [$NgayBatDauDanhSachNgayCapPhep]

;RUNDANHSACHNGAYCAPPHEP()

Func RUNDANHSACHNGAYCAPPHEP()
	Local $NgayStart = StringSplit($NgayBatDauDanhSachNgayCapPhep,"/")
	Global $MangNgayCapPhep[1] = [$NgayBatDauDanhSachNgayCapPhep]
	TangMotNgay($NgayStart[1], $NgayStart[2], $NgayStart[3])
EndFunc
Func TangMotNgay( ByRef $Ngay, ByRef $Thang, ByRef $Nam)
	While True
		If NgayToiDaTrongThang($Thang,$Nam) > $Ngay Then
			$Ngay += 1
		Else
			$Ngay = 1
			If (12 > $Thang) Then
				$Thang += 1
			Else
				$Thang = 1
				$Nam += 1
			EndIf
		EndIf
		$Ngay = NumberToString($Ngay)
		$Thang = NumberToString($Thang)
		$Nam = NumberToString($Nam)
		If StringReplace($NgayKetThucDanhSachNgayCapPhep,$Ngay & "/" & $Thang & "/" & $Nam,"") == "" Then ;Dừng Tăng
			_ArrayAdd($MangNgayCapPhep,$NgayKetThucDanhSachNgayCapPhep)
			return -1
		EndIf
		_ArrayAdd($MangNgayCapPhep,$Ngay & "/" & $Thang & "/" & $Nam)
	WEnd
EndFunc
Func NgayToiDaTrongThang($Thang,$Nam)
	Switch $Thang
		Case 1,3,5,7,8,10,12
			Return 31
		Case 4,6,9,11
			Return 30
		Case 2
			If KiemTraNamNhuan($Nam) Then
				Return 29
			Else
				Return 28
			EndIf
	EndSwitch
EndFunc
Func KiemTraNamNhuan($Nam)
	If Mod($Nam,400) == 0 Or (Mod($Nam,4) == 0 And Mod($Nam,100) <> 0) Then
		Return 1
	EndIf
	Return 0
EndFunc
Func NumberToString($Number)
	Switch $Number
		Case "01"
			Return "01"
		Case "02"
			Return "02"
		Case "03"
			Return "03"
		Case "04"
			Return "04"
		Case "05"
			Return "05"
		Case "06"
			Return "06"
		Case "07"
			Return "07"
		Case "08"
			Return "08"
		Case "09"
			Return "09"
		Case 1 To 9
			Return "0" & $Number
		Case Else
			Return $Number
	EndSwitch
EndFunc
