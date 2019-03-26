#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <GuiMenu.au3>
#include <Timers.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include "Toast.au3"
#Include "Json.au3"

Global $app_name = "Catchment Thief 6000"


$kml = 	'<?xml version="1.0" encoding="UTF-8"?>' & @CRLF & _
		'<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">' & @CRLF & _
		'<Document>' & @CRLF

Local $catchment_kml = FileRead(@ScriptDir & "\Stretton State College catchment placemark.txt")
$kml = $kml & $catchment_kml & @CRLF


;InetGet("https://www.realestate.com.au/buy/with-5-bedrooms-in-calamvale,+qld+4116%3b/list-1?includeSurrounding=false", @ScriptDir & "\fred2.html")

Local $num_listings = get_node("count(//span[@class='property-price '])")
ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $num_listings = ' & $num_listings & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

for $listing_num = 1 to $num_listings

	Local $price = get_node("(//span[@class='property-price '])[" & $listing_num & "]")
	Local $price_arr = StringRegExp($price, ">(.*)<", 1)
	$price = $price_arr[0]
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $price = ' & $price & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	Local $url = get_node("(//span[@class='property-price '])[" & $listing_num & "]/../..//h2/a/attribute::href")
	Local $url_arr = StringRegExp($url, """(.*)""", 1)
	$url = "https://www.realestate.com.au/" & $url_arr[0]
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $url = ' & $url & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	Local $address = get_node("(//span[@class='property-price '])[" & $listing_num & "]/../..//h2/a/span")
	;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $address = ' & $address & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	Local $address_arr = StringRegExp($address, ">(.*)<", 1)

	if @error = 0 Then

		$address = $address_arr[0]
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $address = ' & $address & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		$json_binary = InetRead("http://dev.virtualearth.net/REST/v1/Locations?q=" & $address & "&key=AoxOdc4UEoS2X8S0VhEvdaDrEGwq2dB886KowoLyZShfus0DshEpZTxIFHAWWaqU")
		Local $json = BinaryToString($json_binary)
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $json = ' & $json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		Local $decoded_json = Json_Decode($json)

		Local $latitude = Json_Get($decoded_json, '.resourceSets[0].resources[0].point.coordinates[0]')
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $latitude = ' & $latitude & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
		Local $longitude = Json_Get($decoded_json, '.resourceSets[0].resources[0].point.coordinates[1]')
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $longitude = ' & $longitude & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		$kml = $kml & "<Placemark><name>" & $address & " - " & $price & "</name><description>" & $url & "</description><Point><coordinates>" & $longitude & "," & $latitude & ",0</coordinates></Point></Placemark>" & @CRLF

	EndIf


Next

$kml = $kml & "</Document></kml>"
FileDelete(@ScriptDir & "\out.kml")
FileWrite(@ScriptDir & "\out.kml", $kml)


Exit








Global $attribute_locating_priority = "id|name|title|ng-controller|ng-click|ng-if|class|style|text()"
Global $attribute_locating_priority_arr = StringSplit($attribute_locating_priority, "|", 3)



; Startup SQLite

Local $aResult, $iRows, $iColumns, $iRval
_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
_SQLite_Open()
_SQLite_Exec(-1, "PRAGMA synchronous = OFF;")		; this should speed up DB transactions
_SQLite_Exec(-1, "CREATE TABLE node (level int,type,name,attribs,text);") ; CREATE a Table


HotKeySet("^!l", "dom_to_webtestobject_find")


While 1
    Sleep(100)
WEnd

_SQLite_Close()
_SQLite_Shutdown()



Func dom_to_webtestobject_find()

	Local $chrome_handle = WinActive("[REGEXPTITLE:.* - Google Chrome; REGEXPCLASS:Chrome.*]")

	if $chrome_handle <> 0 Then

		_Toast_Set(0, -1, -1, -1, -1, -1, "", 100, 100)
		_Toast_Show(0, $app_name, "Getting DOM from Chrome ...", -30, False, True)

;		WinActivate("Modules across all disciplines - Janison CLS - Google Chrome")

		Local $pos_arr = WinGetPos($chrome_handle)

		ClipPut("")
		;ControlSend($chrome_handle, "", "", "{APPSKEY}")
		MouseClick("right")
		Sleep(500)
		ControlSend($chrome_handle, "", "", "c{ENTER}{UP}{ENTER}")

		Local $chrome_xpath = ""

		Do

			$chrome_xpath = ClipGet()
		Until StringLen($chrome_xpath) > 0

		$chrome_xpath = StringReplace($chrome_xpath, """", "\""")

		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $chrome_xpath = ' & $chrome_xpath & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		;Local $chrome_xpath = "//*[@id=""btnPerformSearch""]/span[1]"

		Global $toast_label
		Local $dom = ""
		Local $dom_timer = _Timer_Init()

		Do

			ControlClick($chrome_handle, "", "", "left", 1, $pos_arr[2] - 50, $pos_arr[3] - 20)
			ClipPut("")
			ControlSend($chrome_handle, "", "", "copy(document.documentElement.innerHTML){ENTER}")
			Sleep(250)

			$dom = ClipGet()
		Until StringLen($dom) > 0 Or _Timer_Diff($dom_timer) > 5000

		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $dom = ' & $dom & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		if StringLen($dom) = 0 Then

			Exit
		EndIf

		FileDelete(@ScriptDir & "\fred.html")
		FileWrite(@ScriptDir & "\fred.html", $dom)
		_SQLite_Exec(-1, "DELETE FROM node;") ; CREATE a Table


		Global $num_levels = 0


		;$out0 = get_node_attribs($chrome_xpath)
		local $out_arr[0][5]

		_Toast_Show(0, $app_name, "                                                                                                         ", -30, False, True)
		$toast_label = GUICtrlCreateLabel("", 10, 20, 300, 20)

		get_siblings_attribs($chrome_xpath, 1, $out_arr)

		GUICtrlSetData($toast_label, "Building WebTestObject call ...")


		$iRval = _SQLite_GetTable2d(-1, "SELECT * FROM node;", $aResult, $iRows, $iColumns)
		_SQLite_Display2DResult($aResult)


		Local $selenium_xpath = ""

		for $level_num = 1 to $num_levels

			Local $correct_locator = ""

			; if there is only one node in the level (no sibling nodes), then the ancestor "name" is the correct locator

			if StringLen($correct_locator) = 0 Then

				$iRval = _SQLite_GetTable2d(-1, "SELECT count(*) FROM node WHERE level = " & $level_num & ";", $aResult, $iRows, $iColumns)

				if $aResult[1][0] = 1 Then

					$iRval = _SQLite_GetTable2d(-1, "SELECT name FROM node WHERE level = " & $level_num & ";", $aResult, $iRows, $iColumns)
					$correct_locator = "./" & $aResult[1][0]
				EndIf
			EndIf

			; if there are no siblings with the same name (ie. div, span, a, button, etc) as the ancestor node, then the ancestor "name" is the correct locator

			if StringLen($correct_locator) = 0 Then

				$iRval = _SQLite_GetTable2d(-1, "SELECT name FROM node WHERE level = " & $level_num & " AND type = 'A';", $aResult, $iRows, $iColumns)
				Local $name = $aResult[1][0]
				$iRval = _SQLite_GetTable2d(-1, "SELECT count(*) FROM node WHERE level = " & $level_num & " AND name = '" & $name & "';", $aResult, $iRows, $iColumns)

				if $aResult[1][0] = 1 Then

					$correct_locator = "./" & $name
				EndIf
			EndIf

			; search through a priority list of importance of attributes to locate

			if StringLen($correct_locator) = 0 Then



			EndIf



			; if the ancestor node has an id attribute, then the ancestor "id" is the correct locator

			if StringLen($correct_locator) = 0 Then

				$iRval = _SQLite_GetTable2d(-1, "SELECT name, attribs FROM node WHERE level = " & $level_num & " AND type = 'A';", $aResult, $iRows, $iColumns)

				Local $ancestor_id = StringRegExp($aResult[1][1], "(?U)id=""(.*)""", 1)

				if @error = 0 Then

					$correct_locator = "./" & $aResult[1][0] & "[@id='" & $ancestor_id[0] & "']"
				EndIf
			EndIf

			; if there are no siblings with the same name and class combination as the ancestor node, then the ancestor "class" is the correct locator

			if StringLen($correct_locator) = 0 Then

				; get the name and class of the ancestor node
				$iRval = _SQLite_GetTable2d(-1, "SELECT name, attribs FROM node WHERE level = " & $level_num & " AND type = 'A';", $aResult, $iRows, $iColumns)
				Local $ancestor_name = $aResult[1][0]
				Local $ancestor_attribs = $aResult[1][1]
				Local $ancestor_class = StringRegExp($ancestor_attribs, "(?U)class=""(.*)""", 1)

				if @error = 0 Then

					; default the locator to the ancestor class
					$correct_locator = "./" & $ancestor_name & "[@class='" & $ancestor_class[0] & "']"

					; get the name and class of the sibling nodes
					$iRval = _SQLite_GetTable2d(-1, "SELECT name, attribs FROM node WHERE level = " & $level_num & " AND type = 'S';", $aResult, $iRows, $iColumns)

					for $sibling_num = 1 to $aResult[0][0]

		;				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $sibling_num = ' & $sibling_num & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

						Local $sibling_name = $aResult[$sibling_num][0]
						Local $sibling_attribs = $aResult[$sibling_num][1]
						Local $sibling_class = StringRegExp($sibling_attribs, "(?U)class=""(.*)""", 1)

						if @error = 0 Then

							if StringCompare($ancestor_name, $sibling_name) = 0 And StringCompare($ancestor_class, $sibling_class) = 0 Then

								$correct_locator = ""
								ExitLoop
							EndIf
						EndIf
					Next
				EndIf
			EndIf


			Local $level_locator = "level " & $level_num & " = " & $correct_locator
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $level_locator = ' & $level_locator & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

			$selenium_xpath = 			"					By.XPath(""" & $correct_locator & """)," & @CRLF & $selenium_xpath
		Next

		Local $webtestobject_call = 		"			WebTestObject.FindUntilDisplayed(" & @CRLF & _
										"				null," & @CRLF & _
										"				new By[] {" & @CRLF & _
										"					By.TagName(""body"")," & @CRLF & _
										$selenium_xpath & _
										"				}" & @CRLF & _
										"			);" & @CRLF

		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $webtestobject_call = ' & $webtestobject_call & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		ClipPut($webtestobject_call)
		GUICtrlSetData($toast_label, "Clipboard ready (now paste in script).")
		Sleep(3000)
		_Toast_Hide()
	EndIf

EndFunc



Func get_siblings_attribs($node_xpath, $node_ancestor_num, ByRef $out_arr)

	Local $out

	for $ancestor_num = 1 to 100

		GUICtrlSetData($toast_label, "Getting siblings for DOM Ancestor " & $ancestor_num & " ...")

		Local $ancestor_node_name = get_node_name($node_xpath)
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $ancestor_node_name = ' & $ancestor_node_name & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		if StringLen($ancestor_node_name) = 0 or StringCompare($ancestor_node_name, "body") = 0 Then

			$num_levels = $ancestor_num - 1
			ExitLoop
		EndIf

		if StringCompare($ancestor_node_name, "script") <> 0 Then

			Local $ancestor_node_attribs = get_node_attribs($node_xpath)
			Local $ancestor_node_text = get_node_text($node_xpath)
			Local $ansestor_node = $ancestor_node_name & "|" & $ancestor_node_attribs & "|" & $ancestor_node_text
;			Local $out = $ancestor_num & "|A|" & $ansestor_node
;			_ArrayAdd($out_arr, $out)
			$ancestor_node_attribs = StringReplace($ancestor_node_attribs, "'", "''")
			$ancestor_node_text = StringReplace($ancestor_node_text, "'", "''")
			$query = "INSERT INTO node (level,type,name,attribs,text) VALUES ('" & $ancestor_num & "','A','" & $ancestor_node_name & "','" & $ancestor_node_attribs & "','" & $ancestor_node_text & "');"
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $query = ' & $query & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			_SQLite_Exec(-1, $query) ; INSERT Data


			$node_xpath = $node_xpath & "/.."

			Local $child_count = get_node_child_count($node_xpath)
	;		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $node_count = ' & $node_count & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

			for $i = 1 to $child_count

				Local $sibling_node_name = get_node_name($node_xpath & "/*[" & $i & "]")

				if StringCompare($sibling_node_name, "script") <> 0 Then

					Local $sibling_node_attribs = get_node_attribs($node_xpath & "/*[" & $i & "]")
					Local $sibling_node_text = get_node_text($node_xpath & "/*[" & $i & "]")
					Local $sibling_node = $sibling_node_name & "|" & $sibling_node_attribs & "|" & $sibling_node_text

					if StringCompare($ansestor_node, $sibling_node) <> 0 Then

						Local $out = $ancestor_num & "|S|" & $sibling_node
		;				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $out = ' & $out & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

;						_ArrayAdd($out_arr, $out, 0, "|", "~")
						$sibling_node_attribs = StringReplace($sibling_node_attribs, "'", "''")
						$sibling_node_attribs = StringStripCR($sibling_node_attribs)
						$sibling_node_text = StringReplace($sibling_node_text, "'", "''")
						$query = "INSERT INTO node (level,type,name,attribs,text) VALUES ('" & $ancestor_num & "','S','" & $sibling_node_name & "','" & $sibling_node_attribs & "','" & $sibling_node_text & "');"
						_SQLite_Exec(-1, $query) ; INSERT Data
					EndIf
				EndIf
			Next
		EndIf
	Next

EndFunc


Func get_node_name($xpath)

	Local $out = get_node("name(" & $xpath & ")")
	return $out
EndFunc

Func get_node_attribs($xpath)

	Local	$out = get_node($xpath & "/attribute::*")
	return $out
EndFunc


Func get_node_text($xpath)

	Local $out = get_node($xpath & "/text()")
	$out = StringStripWS($out, 3)
	return $out
EndFunc

Func get_node_child_count($xpath)

	Local $out = get_node("count(" & $xpath & "/*)")
	return $out

EndFunc


Func get_node($xpath)

	Local $sOutput = ""

	FileDelete(@ScriptDir & "\xpath.txt")
	Local $cmd = "xmllint.exe --html --xpath """ & $xpath & """ fred.html > xpath.txt"
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $cmd = ' & $cmd & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	Local $iPID = RunWait(@ComSpec & " /c " & $cmd, @ScriptDir, @SW_HIDE)

	if FileExists(@ScriptDir & "\xpath.txt") = True Then

		$sOutput = FileRead(@ScriptDir & "\xpath.txt")
;		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $sOutput = ' & $sOutput & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	EndIf

	return $sOutput

EndFunc


Func _URIEncode($sData)
    ; Prog@ndy
    Local $aData = StringSplit(BinaryToString(StringToBinary($sData,4),1),"")
    Local $nChar
    $sData=""
    For $i = 1 To $aData[0]
        ; ConsoleWrite($aData[$i] & @CRLF)
        $nChar = Asc($aData[$i])
        Switch $nChar
            Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
                $sData &= $aData[$i]
            Case 32
                $sData &= "+"
            Case Else
                $sData &= "%" & Hex($nChar,2)
        EndSwitch
    Next
    Return $sData
EndFunc

Func _URIDecode($sData)
    ; Prog@ndy
    Local $aData = StringSplit(StringReplace($sData,"+"," ",0,1),"%")
    $sData = ""
    For $i = 2 To $aData[0]
        $aData[1] &= Chr(Dec(StringLeft($aData[$i],2))) & StringTrimLeft($aData[$i],2)
    Next
    Return BinaryToString(StringToBinary($aData[1],1),4)
EndFunc

