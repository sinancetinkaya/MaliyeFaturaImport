#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
AeroFix:=0
IniRead, ilk_fatura, FaturaImport.ini, son, FaturaNo
InputBox, ilk_fatura, Fatura Numarası, Lütfen import edilecek İLK fatura numarasını gir., , 150, 170,,,,,%ilk_fatura%
if ErrorLevel
		exit
InputBox, son_fatura, Fatura Numarası, Lütfen import edilecek SON fatura numarasını gir., , 150, 170,,,,,
if ErrorLevel
		exit

InputBox, FaturaSeri, Fatura Seri, Fatura üzerinde Seri Kodunu girin, , 150, 170,,,,,D
if ErrorLevel
		exit



WinActivate e:\ceren\firin.exe
WinWaitActive e:\ceren\firin.exe
Color=
PixelGetColor, Color, 25, 50
if  Color <> 0x000080
{
	ToolTip , Fırın programı fatura bilgilerinin alınacağı ekranda değil!!!
	sleep,4000
	exit
}

	
SiraNo=0
Send,%ilk_fatura%{ENTER}{UP}{F5}

Color =
While Color <> 0x800000
	PixelGetColor, Color,268,268
Ekran:=EkranKopyala()

While sayi_al(Ekran,5, 22, 11) >= ilk_fatura && sayi_al(Ekran,5, 22, 11) <= son_fatura
{
	pos:=16

	temp:=metin_al(Ekran,pos, 8, 24)

	While SubStr(temp,1,1) != "" and SubStr(temp,1,1) != "─"
	{
		SiraNo++
		FaturaNo%SiraNo%:=sayi_al(Ekran,5, 22, 11)
		Tarih%SiraNo%:=metin_al(Ekran,7, 22, 10)
		BayiNo%SiraNo%:=sayi_al(Ekran,9, 22, 5)
		Ad%SiraNo%:=metin_al(Ekran,9,28,30) "." metin_al(Ekran,10,28,30)
		temp:=metin_al(Ekran,pos, 8, 24)
		Aciklama%SiraNo%=%temp%
		Miktar%SiraNo%:=sayi_al(Ekran,pos, 33, 7)
		Birim%SiraNo%=Adet
		Matrah%SiraNo%:=sayi_al(Ekran,pos, 49, 11)
		KDVTutar%SiraNo%:=sayi_al(Ekran,pos, 61, 7)
		StringReplace, t_Matrah, Matrah%SiraNo%,`,,., All
		StringReplace, t_KDVTutar, KDVTutar%SiraNo%,`,,., All
		KDVOran%SiraNo%:= round(t_KDVTutar/(t_Matrah/100))
		pos++
		temp:=metin_al(Ekran,pos, 8, 24)
	}
	Send,{ESC}

	color=
	While color != "0x000000"
		PixelGetColor, color, 280+AeroFix, 268

	Send,{PgDn}{F5}

	color=
	While color != "0x800000"
		PixelGetColor, color, 280+AeroFix, 268

	Ekran:=EkranKopyala()

}
Send,{ESC 2}{UP 2}{ENTER}

color=
While color != "0x000080"
	PixelGetColor, color, 540+AeroFix, 400

Ekran:=

loop,%SiraNo%
{
	Temp:=BayiNo%A_Index%
	if Temp is not space
	{
		if StrLen(VergiDairesi%Temp%) = 0	;kayıt okunmuşsa aynı kayıt için bir daha işlem yapma
		{
			Send,%Temp%{ENTER}{UP}
			Sleep,200
			While sayi_al(Ekran,4,19,5) != Temp
				Ekran:=EkranKopyala()
			Adres%Temp%:=metin_al(Ekran,13,19,30) " " metin_al(Ekran,14,19,30) " " metin_al(Ekran,15,25,14) "/" metin_al(Ekran,16,19,20)
			VergiDairesi%Temp%:=metin_al(Ekran,19,19,20)
			VergiNumarasi%Temp%:=metin_al(Ekran,19,60,15)
		}
	}
}
StringSplit, word_array, Tarih1, /
YilAy := word_array3 . word_array2 . word_array1 . "000000"
FormatTime, YilAy, %YilAy%,yyyy.MM MMMM
InputBox, YilAy, Ay, Yıl.Ay'ı Girin:, , 150, 170,,,,,%YilAy%
if ErrorLevel
		exit

IfWinNotExist, Microsoft Excel
	Excel := ComObjCreate("Excel.Application")
else
	Excel := ComObjActive("Excel.Application")


Excel.Workbooks.Add
Excel.Visible:=True
Kitap:=Excel.ActiveWorkbook
Excel.EnableEvents := False
Excel.ScreenUpdating := False

Kitap.ActiveSheet.Range("A1").Offset(0,2).Value:="XXXXXXXX Ekmek limited Şirketi " YilAy " ayı satış faturaları listesi"
Kitap.ActiveSheet.Range("A1").Offset(2,0).Value:="Sıra No"
Kitap.ActiveSheet.Range("A1").Offset(2,1).Value:="Satış faturasını düzenleyenin adı soyadı ve ünvanı"
Kitap.ActiveSheet.Range("A1").Offset(2,2).Value:="Adresi"
Kitap.ActiveSheet.Range("A1").Offset(2,3).Value:="Faturanın seri numarası"
Kitap.ActiveSheet.Range("A1").Offset(2,4).Value:="Satış faturasının tarihi"
Kitap.ActiveSheet.Range("A1").Offset(2,5).Value:="Satış faturasının nosu"
Kitap.ActiveSheet.Range("A1").Offset(2,6).Value:="Vergi Dairesi"
Kitap.ActiveSheet.Range("A1").Offset(2,7).Value:="Vergi No"
Kitap.ActiveSheet.Range("A1").Offset(2,8).Value:="Açıklama"
Kitap.ActiveSheet.Range("A1").Offset(2,9).Value:="Miktar"
Kitap.ActiveSheet.Range("A1").Offset(2,10).Value:="Ölçü birimi"
Kitap.ActiveSheet.Range("A1").Offset(2,11).Value:="Matrah"
Kitap.ActiveSheet.Range("A1").Offset(2,12).Value:="KDV TUTARI(%18)"
Kitap.ActiveSheet.Range("A1").Offset(2,13).Value:="KDV TUTARI(%8)"
Kitap.ActiveSheet.Range("A1").Offset(2,14).Value:="KDV TUTARI(%1)"
Kitap.ActiveSheet.Range("A1").Offset(2,15).Value:="TOPLAM KDV"

SatirNo:=0
Loop, %SiraNo%
{ 					
	if(FaturaNo%A_Index%="")
						break
	e_Index:=A_Index-1
	FaturaFark:=FaturaNo%A_Index% - FaturaNo%e_Index% - 1
	loop,%FaturaFark%
	{
		Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,0).Value:=SatirNo+1
		Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,3).Value:=FaturaSeri
		Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,5).Value:=FaturaNo%e_Index%+A_Index
		SatirNo++
	}
	_BayiNo:=BayiNo%A_Index%
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,0).Value:=SatirNo+1
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,1).Value:=Ad%A_Index%
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,2).Value:=Adres%_BayiNo%
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,3).Value:=FaturaSeri
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,4).Value:=Tarih%A_Index%
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,5).Value:=FaturaNo%A_Index%
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,6).Value:=VergiDairesi%_BayiNo%
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,7).Value:=VergiNumarasi%_BayiNo%
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,8).Value:=Aciklama%A_Index%
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,9).Value:=Miktar%A_Index%
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,10).Value:=Birim%A_Index%

	StringReplace, Matrah%A_Index%, Matrah%A_Index%,.,`,, All
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,11).Value:=Matrah%A_Index%
	
	StringReplace, KDVTutar%A_Index%, KDVTutar%A_Index%,.,`,, All
	if(KDVOran%A_Index%=18) 
		Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,12).Value:=KDVTutar%A_Index%
	if(KDVOran%A_Index%=8)  
		Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,13).Value:=KDVTutar%A_Index%
	if(KDVOran%A_Index%=1)  
		Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,14).Value:=KDVTutar%A_Index%
	Kitap.ActiveSheet.Range("A1").Offset(SatirNo+3,15).Value:=KDVTutar%A_Index%
	SatirNo++
}

Kitap.ActiveSheet.Columns.EntireColumn.AutoFit 
Kitap.SaveAs(Filename:="e:\Belgelerim\Excel\Maliye Fatura Listesi\" . YilAy . ".xls")
Excel.EnableEvents := True
Excel.ScreenUpdating := True

son_fatura++
IniWrite, %son_fatura%, FaturaImport.ini, son, FaturaNo

metin_al(Ekran,x,y,byte)
{
	word_array := StrSplit(Ekran, "`n")  
	deger:=SubStr(word_array[x], y, byte)	
	deger=%deger%
	;OutputDebug, % "metin_al x:" x " y:" y " deger:" deger
	return %deger%
}

sayi_al(Ekran,x,y,byte)
{
	word_array := StrSplit(Ekran, "`n")  
	deger:=SubStr(word_array[x], y, byte)	
	deger=%deger%
	StringReplace, deger, deger,`,,,, All
	StringReplace, deger, deger, `. , `, ,, All
	;OutputDebug, % "sayi_al x:" x " y:" y " deger:" deger
	return %deger%
}

EkranKopyala()
{

	;MouseClickDrag L,15,35,805+AeroFix,480,0 ;aero
	WinGetTitle, title, A
	if (A != "e:\ceren\firin.exe") {
		WinActivate e:\ceren\firin.exe
		WinWaitActive e:\ceren\firin.exe
	}
	MouseClickDrag L,15,35,800,470,0
	Clipboard := ""
	Color := ""
;	While PixelColor(800, 470, "e:\ceren\firin.exe") != "0xFFFFFF"
;		Sleep, 50
	While (Color <> "0xFFFFFF")
		PixelGetColor, Color, 800 , 470
	Click, right
	ClipWait
	return %Clipboard%
}
