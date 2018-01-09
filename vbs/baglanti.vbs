
Public vt_yolu

'vt_yolu kontrolü
vt_yolu="db\eop.mdb"
tablo_adi="ilaclar"									'|
'***************************************************'|
Set conn = CreateObject("ADODB.Connection")			'|
Set rs = CreateObject("ADODB.Recordset")			'|
													'|
sProviderName        = "Microsoft.Jet.OLEDB.4.0"	'|
iCursorType          = 3							'|
iLockType            = 3							'|
sDataSource          = vt_yolu						'|
													'|
conn.Provider = sProviderName						'|
conn.Properties("Data Source") = sDataSource		'|
conn.Open											'|
													'|
rs.CursorType = iCursorType							'|
rs.LockType = iLockType								'|
rs.Source = tablo_adi								'|
rs.ActiveConnection = conn	
'==================================================================================================

'***************************************************'|
' Bu Kýsýmda VT Baðlantýsý Açýp-Kapama Yapýlýyor	'|
													'|
'***************************************************'|
Function copen()									'|
'vt_yolu="db\eop.mdb"								'|
tablo_adi="ilaclar"									'|
'***************************************************'|
Set conn = CreateObject("ADODB.Connection")			'|
Set rs = CreateObject("ADODB.Recordset")			'|
													'|
sProviderName        = "Microsoft.Jet.OLEDB.4.0"	'|
iCursorType          = 3							'|
iLockType            = 3							'|
sDataSource          = vt_yolu						'|
													'|
conn.Provider = sProviderName						'|
conn.Properties("Data Source") = sDataSource		'|
conn.Open											'|
													'|
rs.CursorType = iCursorType							'|
rs.LockType = iLockType								'|
rs.Source = tablo_adi								'|
rs.ActiveConnection = conn							'|
End Function										'|
'|||||||||||||||||||||||||||||||||||||||||||||||||||||
Function rskapatici()								'|
		If rs.state = 1 then						'|
			rs.close								'|
			Set rs=Nothing							'|
		Elseif rs.state = 0 then					'|
			Set rs=Nothing							'|
		End if										'|
Call cclose()										'|
Call copen()										'|
End Function										'|
													'|
Function cclose()
Call copen()										'|
        If conn.state=1 then						'|
			conn.close								'|
			Set conn=Nothing						'|
        Elseif conn.state=0 then					'|
            Set conn=Nothing						'|
        End if										'|
End Function						'|
'|||||||||||||||||||||||||||||||||||||||||||||||||||||
'==================================================================================================


	function degistir(degeri)
		if degeri=1 then
			window.event.srcElement.className="ic_hover"
		Elseif degeri=2 then
			window.event.srcElement.className="ic_mousedown"
		Else
			window.event.srcElement.className="ic"
		End if
End Function

'||||||||||||||||||||||||||||||||||||||||||||||||||||||
function noktala(girdi,koyulacaksimge)
girdi = Replace(girdi,koyulacaksimge,"")
	if girdi="" or girdi=Null OR not isnumeric(girdi)then
	sonhal="Hata oluþtu..."
	exit function
		Else
		varmi=Instr(1, girdi, ",")
		'msgbox(varmi)
		yeniduzen=Split(girdi,",")
			if UBound(yeniduzen)>0 then
				girdi=yeniduzen(0)
			End if

uzunluk = Len(girdi)
			if uzunluk<>0 then
			sonhal=""
			gecici=""
			sayac=1
							For x = uzunluk to 1 STEP -1
							gecici= mid(girdi,x,1)

										if sayac MOD 3 = 0 then
										sonhal=koyulacaksimge + gecici + sonhal
										Else
										sonhal = gecici + sonhal
										End if
							sayac=sayac+1
							Next

						If len(girdi) MOD 3 =0 then
						sonhal = right(sonhal,(len(sonhal)-1))
						End if
			Else
			sonhal="Yine hata"
			End if
End if



			if UBound(yeniduzen)>0 then
			sonhal=sonhal & "," & yeniduzen(1)
			End if
noktala=sonhal
End Function
'---------------------------------
Function atesleme(hedef)

' div oluþturuluyor
Set tt = document.createElement("DIV")
tt.id="takvim_tasiyici"
tt.style.width="300 px"
tt.style.position="absolute"
' ekran geniþliðini alalým
tt.style.top=window.event.ClientY+20
tt.style.left=window.event.ClientX-500
tt.style.backGroundColor="black"
document.body.appendChild(tt)


'ay_gezer, ay degeri taþiyici oluþturuluyor
Set ag = document.createElement("INPUT")
ag.type="TEXT"
ag.id="ay_gezer"
ag.style.visibility="hidden"
ag.value="0"
ag.style.position="absolute"
ag.style.top=0
ag.style.left=0
ag.style.zIndex=-1
document.body.appendChild(ag)


'hedef_id degeri taþiyici oluþturuluyor
Set hi = document.createElement("INPUT")
hi.type="TEXT"
hi.id="hedef_id"
hi.style.visibility="hidden"
hi.value=hedef
hi.style.position="absolute"
hi.style.top=0
hi.style.left=0
hi.style.zIndex=-2
document.body.appendChild(hi)


'ilk kez takvim çaðrýlýyor
Call takvimle(0)
End function



Function isinlan(uzaklik)
document.getElementById("ay_gezer").value = CInt(document.getElementById("ay_gezer").value) + CInt(uzaklik)
Call takvimle(document.getElementById("ay_gezer").value)
End Function


' hedef_id ye týklanan deger atanacak -artý yaratýlan tüm elementler yok edilecek-temizlik
Function hedefe_tarihver(tarih)
document.getElementById(document.getElementById("hedef_id").value).value=tarih
document.body.removeChild(document.getElementById("takvim_tasiyici"))
document.body.removeChild(document.getElementById("ay_gezer"))
document.body.removeChild(document.getElementById("hedef_id"))
End function



Function takvimle(ay)
if ay="" OR ay=Null then
ay=0
End if
ytarih=DateAdd("m",ay,now)
ay=month(ytarih)
yil=Year(ytarih)
bugun=day(ytarih) ' yeni tarih deðeri alýnýyor bu üstünde durduðumuz ay oluyor...
ilkgunhaftaninkacincisi=weekday((Dateadd("d",-(bugun-1),ytarih)),2) ' haftanin hangi günü , 3. gibi bir deðer verecek
ilkgun=weekdayname(ilkgunhaftaninkacincisi,false,2) ' çarþamba gibi bir deðer verecek
sonrakiay=DateSerial(yil,(ay+1),1)
simdikiay=DateAdd("d",-1,sonrakiay)
kacceker=Day(simdikiay)

icerik="<table width=100% cellpadding=0 cellspacing=1 border=0 style=border-style:solid;border-width:1px;>"
icerik=icerik &"<tr><td class=""ic"" colspan=""4"">"
 icerik=icerik &"<input type=""button"" onclick=""isinlan('-12')"" title=""Önceki Yýl"" class=""textbox1"" value=""&lt;&lt;"">"
 icerik=icerik &"<input type=""button"" onclick=""isinlan('-1')"" title=""Önceki Ay"" class=""textbox1"" value=""&lt;"">"
 icerik=icerik &"&nbsp;"
 icerik=icerik &"<input type=""button"" onclick=""takvimle('')"" class=""textbox1"" value=""Bu Gün'e Gel"">"
 icerik=icerik &"&nbsp;"
 icerik=icerik &"<input type=""button"" onclick=""isinlan('1')"" title=""Sonraki Ay"" class=""textbox1"" value=""&gt;"">"
 icerik=icerik &"<input type=""button"" onclick=""isinlan('+12')"" title=""Sonraki Yýl"" class=""textbox1"" value=""&gt;&gt;"">"
icerik=icerik &"</td>"
icerik=icerik &"<td colspan=""3"" class=""ic""><b>"& MonthName(ay) &","&Year(simdikiay)&"</b>"
icerik=icerik &"</td></tr>"


icerik=icerik & "<tr> <td class=""ic_gunler"">Ptesi</td><td class=""ic_gunler"">Salý</td><td  class=""ic_gunler"" >Çarþ.</td><td class=""ic_gunler"">Perþ.</td><td class=""ic_gunler"">Cuma</td><td  class=""ic_gunler"">Ctesi</td><td class=""ic_gunler"">Pazar</td> </tr><tr>"
saybakim=0
ilkgunhaftaninkacincisi=ilkgunhaftaninkacincisi-1 ' burada haftanin gunu TR uyarlamasi farký eklendi yani 1
For gunumuz=1 to kacceker + ilkgunhaftaninkacincisi

if saybakim < ilkgunhaftaninkacincisi then
tarihx=dateserial(yil,ay,gunumuz-ilkgunhaftaninkacincisi)
		saybakim=(saybakim+1)
oncekiayingunu=Dateadd("d",-(ilkgunhaftaninkacincisi-saybakim+1),(DateSerial(yil,ay,1)))
icerik = icerik & "<td class=""ic"" onclick=isinlan(-1) title=Bu&nbsp;aya&nbsp;git style=""cursor:hand"">"&MonthName(Month(oncekiayingunu)) &"</td>"
Else
		tarih=dateserial(yil,ay,gunumuz-ilkgunhaftaninkacincisi)
	Select case ilkgun
		case "Pazar"
		kac=1
		case "Pazartesi"
		kac=2
		case "Salý"
		kac=3
		case "Çarþamba"
		kac=4
		case "Perþembe"
		kac=5
		case "Cuma"
		kac=6
		case "Cumartesi"
		kac=7
	End Select
		kacincigun=weekday(tarih)

		Select Case kacincigun
			case 1
klass="haftasonu"
			case 7
klass="haftasonu"
			Case Else
klass="gun"
		End Select
if Day(tarih)=Day(now) then
klass="bugun"
Elseif tarih<now AND kacincigun<>1 AND kacincigun<>7 then
klass="oncekigun"
End if
' tarih=Replace(tarih,".",".")
icerik = icerik & "<td><a href=""#"" class="& klass &" id="&tarih&" onclick=""hedefe_tarihver(this.id)"">"&Day(tarih) & "&nbsp;</a></td>"
End if
if gunumuz mod 7 = 0 then
		icerik=icerik & "</tr><tr>"
End if

Next
gunsayici=0
if kacceker+ilkgunhaftaninkacincisi MOD 7 <> 0 then
	degergun = weekday(simdikiay,2) +1 
Else	
	degergun=1
End if
	For xcv = degergun to 7
	gunsayici = gunsayici + 1
sonrakiayingunu=Dateadd("d",gunsayici,(DateSerial(yil,ay,kacceker)))
	icerik = icerik & "<td align=right valign=bottom class=ic onclick=isinlan(1) title=Bu&nbsp;aya&nbsp;git style=cursor:hand>"&MonthName(Month(sonrakiayingunu)) &"</td>"
	Next
icerik =icerik &"</tr></table>"
document.getElementbyId("takvim_tasiyici").innerHTML=icerik
 End Function

'----------- Bu Fonksiyon Bitti -----------------------
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


document.title= "EpicReady "&  document.title


Function buyukharf()
window.event.srcElement.value=UCase(window.event.srcElement.value)
End Function

Function sadecerakamgir()
	if not isnumeric(window.event.srcElement.value) then
		msgbox("Girmeye çalýþtýðýnýz deðer bir rakam deðil, bu kýsma sadece rakam girebilirsiniz!")
		window.event.srcElement.value=""
		window.event.srcElement.focus()
	End if
End Function


' tarih formatýný ISO sormata dönüþtüren hede höd
Function JXIsoDate(dteDate)
'Version 1.0
   If IsDate(dteDate) = True Then
      DIM dteDay, dteMonth, dteYear
      dteDay = Day(dteDate)
      dteMonth = Month(dteDate)
      dteYear   = Year(dteDate)
      JXIsoDate =  dteYear & "-" & Right(Cstr(dteMonth + 100),2) & "-" & Right(Cstr(dteDay + 100),2)    
   Else
      JXIsoDate = Null
   End If
End Function


' TARÝH FORMATINI ýso FORMATA DÖNÜÞTÜREN HEDE HÖD 2 VE SON
Function JXIsoDate2(dteDate)
dteDate=Replace(dteDate,".","/")
'Version 2.0
   If IsDate(dteDate) = True Then
      DIM dteDay, dteMonth, dteYear
      dteDay = Day(dteDate)
      dteMonth = Month(dteDate)
      dteYear   = Year(dteDate)
      JXIsoDate =  dteYear & "-" & Right(Cstr(dteMonth + 100),2) & "-" & Right(Cstr(dteDay + 100),2)    
   Else
      JXIsoDate = Null
   End If
End Function





' herhangi bir sayfadan içeriði verilen (icHTML) fonksiyonla IE açarak dokum aldýrma FONKSÝYONU
Sub InitIE(icHTML,e,b,x,y)

Set objShell = CreateObject("WScript.Shell")
' Subroutine to initialize the IE display box.
  Dim intWidth, intHeight, intWidthW, intHeightW
  Set objIE = CreateObject("InternetExplorer.Application")

  With objIE
    .ToolBar = False
    .StatusBar = False
    .Resizable = False
    .Navigate("about:blank")
    With .document
      With .ParentWindow
if e="" then e=.Screen.AvailWidth
if b="" then b=.Screen.AvailHeight
        intWidth = .Screen.AvailWidth
        intHeight = .Screen.AvailHeight
        intWidthW = e
        intHeightW = b
        .resizeto intWidthW, intHeightW
        .moveto x,y
      End With
      .Write icHTML &""
      .Title = strIETitle
      objIE.Visible = True
	  objIE.ToolBar = 0
		objIE.StatusBar = 0

	  'objIE.Quit
    End With
  End With
End Sub

