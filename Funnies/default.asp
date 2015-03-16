<%@ Language=VBScript %>
<%
 	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Scrapes over 100 different online comics and
	' presents a comic page similar to the old newspaper
	' style some of us older programmers grew up reading.
	'
	Function GetImagefile(strURL, strMatch)
		Dim webdoc, str, pos, startpos

		Server.ScriptTimeout = Server.ScriptTimeout + 10

		Set webdoc = CreateObject("MSXML2.ServerXMLHTTP")
		webdoc.Open "GET", strURL, False
		webdoc.Send

		str = webDoc.responseText
		pos = InStr(1, str, strMatch, 1)
		If pos > 0 Then
			startpos = InStrRev(str, "<img src=""", pos, 1)
			str = Mid(str, startpos + 10)
			pos = InStr(str, """")
			GetImagefile = left(str, pos - 1)
		End If
	End Function

	Response.Buffer = True
%>
<html>
<body>
<%
Dim a(149), strImg, i

a(1) = "http://www.comics.com/comics/dilbert/index.html"
a(2) = "http://www.comics.com/comics/peanuts/index.html"
a(3) = "http://www.comics.com/comics/alleyoop/index.html"
a(4) = "http://www.comics.com/comics/arlonjanis/index.html"
a(5) = "http://www.comics.com/comics/ben/index.html"
a(6) = "http://www.comics.com/comics/betty/index.html"
a(7) = "http://www.comics.com/comics/bignate/index.html"
a(8) = "http://www.comics.com/comics/bornloser/index.html"
a(9) = "http://www.comics.com/comics/buckets/index.html"
a(10) = "http://www.comics.com/comics/bullsnbears/index.html"
a(11) = "http://www.comics.com/comics/chickweed/index.html"
a(12) = "http://www.comics.com/comics/committed/index.html"
a(13) = "http://www.comics.com/comics/drabble/index.html"
a(14) = "http://www.comics.com/comics/fatcats/index.html"
a(15) = "http://www.comics.com/comics/ferdnand/index.html"
a(16) = "http://www.comics.com/comics/forbetter/index.html"
a(17) = "http://www.comics.com/comics/franknernest/index.html"
a(18) = "http://www.comics.com/comics/frazz/index.html"
a(19) = "http://www.comics.com/comics/geech/index.html"
a(20) = "http://www.comics.com/comics/getfuzzy/index.html"
a(21) = "http://www.comics.com/comics/gingermeggs/index.html"
a(22) = "http://www.comics.com/comics/gofigure/index.html"
a(23) = "http://www.comics.com/comics/graffiti/index.html"
a(24) = "http://www.comics.com/comics/grandave/index.html"
a(25) = "http://www.comics.com/comics/grizzwells/index.html"
a(26) = "http://www.comics.com/comics/hedge/index.html"
a(27) = "http://www.comics.com/comics/herman/index.html"
a(28) = "http://www.comics.com/comics/janesworld/index.html"
a(29) = "http://www.comics.com/comics/jumpstart/index.html"
a(30) = "http://www.comics.com/comics/kitncarlyle/index.html"
a(31) = "http://www.comics.com/comics/lilabner/index.html"
a(32) = "http://www.comics.com/comics/luann/index.html"
a(33) = "http://www.comics.com/comics/lupo/index.html"
a(34) = "http://www.comics.com/comics/marmaduke/index.html"
a(35) = "http://www.comics.com/comics/meatloaf/index.html"
a(36) = "http://www.comics.com/comics/meg/index.html"
a(37) = "http://www.comics.com/comics/moderatelyconfused/index.html"
a(38) = "http://www.comics.com/comics/monty/index.html"
a(39) = "http://www.comics.com/comics/motley/index.html"
a(40) = "http://www.comics.com/comics/nancy/index.html"
a(41) = "http://www.comics.com/comics/offthemark/index.html"
a(42) = "http://www.comics.com/comics/pearls/index.html"
a(43) = "http://www.comics.com/comics/pibgorn/index.html"
a(44) = "http://www.comics.com/comics/potluck/index.html"
a(45) = "http://www.comics.com/comics/raisingduncan/index.html"
a(46) = "http://www.comics.com/comics/reality/index.html"
a(47) = "http://www.comics.com/comics/ripleys/index.html"
a(48) = "http://www.comics.com/comics/roseisrose/index.html"
a(49) = "http://www.comics.com/comics/rudypark/index.html"
a(50) = "http://www.comics.com/comics/sheldon/index.html"
a(51) = "http://www.comics.com/comics/shirleynson/index.html"
a(52) = "http://www.comics.com/comics/soup2nutz/index.html"
a(53) = "http://www.comics.com/comics/stockcartoons/index.html"
a(54) = "http://www.comics.com/comics/tarzan/index.html"
a(55) = "http://www.comics.com/comics/topofworld/index.html"
a(56) = "http://www.comics.com/comics/workingdaze/index.html"
a(57) = "http://www.comics.com/creators/agnes/index.html"
a(58) = "http://www.comics.com/creators/andycapp/index.html"
a(59) = "http://www.comics.com/creators/bachelorparty/index.html"
a(60) = "http://www.comics.com/creators/ballardst/index.html"
a(61) = "http://www.comics.com/creators/bc/index.html"
a(62) = "http://www.comics.com/creators/charlie/index.html"
a(63) = "http://www.comics.com/creators/flightdeck/index.html"
a(64) = "http://www.comics.com/creators/floandfriends/index.html"
a(65) = "http://www.comics.com/creators/heathcliff/index.html"
a(66) = "http://www.comics.com/creators/herbnjamaal/index.html"
a(67) = "http://www.comics.com/creators/liberty/index.html"
a(68) = "http://www.comics.com/creators/momma/index.html"
a(69) = "http://www.comics.com/creators/naturalselection/index.html"
a(70) = "http://www.comics.com/creators/onebighappy/index.html"
a(71) = "http://www.comics.com/creators/othercoast/index.html"
a(72) = "http://www.comics.com/creators/rubes/index.html"
a(73) = "http://www.comics.com/creators/speedbump/index.html"
a(74) = "http://www.comics.com/creators/strangebrew/index.html"
a(75) = "http://www.comics.com/creators/wizardofid/index.html"
a(76) = "http://www.comics.com/creators/workingitout/index.html"
a(77) = "http://www.comics.com/wash/bonanas/index.html"
a(78) = "http://www.comics.com/wash/cheapthrills/index.html"
a(79) = "http://www.comics.com/wash/genepool/index.html"
a(80) = "http://www.comics.com/wash/pcnpixel/index.html"
a(81) = "http://www.comics.com/wash/pickles/index.html"
a(82) = "http://www.comics.com/wash/redandrover/index.html"
a(83) = "http://www.comics.com/wash/thatslife/index.html"
a(84) = "http://www.ucomics.com/9to5/"
a(85) = "http://www.ucomics.com/adamathome/"
a(86) = "http://www.ucomics.com/animalcrackers/"
a(87) = "http://www.ucomics.com/annie/"
a(88) = "http://www.ucomics.com/baldo/"
a(89) = "http://www.ucomics.com/bigtop/"
a(90) = "http://www.ucomics.com/boondocks/"
a(91) = "http://www.ucomics.com/bornlucky/"
a(92) = "http://www.ucomics.com/bottomliners/"
a(93) = "http://www.ucomics.com/boundandgagged/"
a(94) = "http://www.ucomics.com/brendastarr/"
a(95) = "http://www.ucomics.com/broomhilda/"
a(96) = "http://www.ucomics.com/calvinandhobbes/"
a(97) = "http://www.ucomics.com/captainribman/"
a(98) = "http://www.ucomics.com/cathy/"
a(99) = "http://www.ucomics.com/catswithhands/"
a(100) = "http://www.ucomics.com/citizendog/"
a(101) = "http://www.ucomics.com/cleats/"
a(102) = "http://www.ucomics.com/closetohome/"
a(103) = "http://www.ucomics.com/compu-toon/"
a(104) = "http://www.ucomics.com/cornered/"
a(105) = "http://www.ucomics.com/dicktracy/"
a(106) = "http://www.ucomics.com/doodles/"
a(107) = "http://www.ucomics.com/doonesbury/"
a(108) = "http://www.ucomics.com/duplex/"
a(109) = "http://www.ucomics.com/forbetterorforworse/"
a(110) = "http://www.ucomics.com/foxtrot/"
a(111) = "http://www.ucomics.com/fredbasset/"
a(112) = "http://www.ucomics.com/garfield/"
a(113) = "http://www.ucomics.com/gasolinealley/"
a(114) = "http://www.ucomics.com/gilthorp/"
a(115) = "http://www.ucomics.com/heartofthecity/"
a(116) = "http://www.ucomics.com/helen/"
a(117) = "http://www.ucomics.com/housebroken/"
a(118) = "http://www.ucomics.com/inthebleachers/"
a(119) = "http://www.ucomics.com/james/"
a(120) = "http://www.ucomics.com/kudzu/"
a(121) = "http://www.ucomics.com/lacucaracha/"
a(122) = "http://www.ucomics.com/lola/"
a(123) = "http://www.ucomics.com/looseparts/"
a(124) = "http://www.ucomics.com/luckycow/"
a(125) = "http://www.ucomics.com/meehanstreak/"
a(126) = "http://www.ucomics.com/misterboffo/"
a(127) = "http://www.ucomics.com/mixedmedia/"
a(128) = "http://www.ucomics.com/mrpotatohead/"
a(129) = "http://www.ucomics.com/nonsequitur/"
a(130) = "http://www.ucomics.com/oddlyenough/"
a(131) = "http://www.ucomics.com/overboard/"
a(132) = "http://www.ucomics.com/pluggers/"
a(133) = "http://www.ucomics.com/poochcafe/"
a(134) = "http://www.ucomics.com/preteena/"
a(135) = "http://www.ucomics.com/reallifeadventures/"
a(136) = "http://www.ucomics.com/reynoldsunwrapped/"
a(137) = "http://www.ucomics.com/shoe/"
a(138) = "http://www.ucomics.com/shoecabbage/"
a(139) = "http://www.ucomics.com/stonesoup/"
a(140) = "http://www.ucomics.com/sylvia/"
a(141) = "http://www.ucomics.com/tankmcnamara/"
a(142) = "http://www.ucomics.com/thebigpicture/"
a(143) = "http://www.ucomics.com/thefifthwave/"
a(144) = "http://www.ucomics.com/thefuscobrothers/"
a(145) = "http://www.ucomics.com/themiddletons/"
a(146) = "http://www.ucomics.com/thequigmans/"
a(147) = "http://www.ucomics.com/tomthedancingbug/"
a(148) = "http://www.ucomics.com/willynethel/"
a(149) = "http://www.ucomics.com/ziggy/"

For i = 1 to 149
	Dim strMatch, strPrefix

	strPrefix = "http://www.comics.com"
	If i > 83 Then
		strMatch = "http://images.ucomics.com/comics"
		strPrefix = ""
	ElseIf i > 2 Then
		strMatch = "Today's Comic"
	ElseIf i > 1 Then
		strMatch = "Today's Strip"" BORDER"
	Else
		strMatch = "Today's Dilbert Comic"
	End If

	strImg = GetImagefile(a(i), strMatch)
	If strImg <> "" then
    	Response.Write "<img src=""" & strPrefix & strImg & """><br>"
	End If
  	Response.Write "<a href=""" & a(i) & """>" & a(i) & "</a><br>"
	Response.Flush
Next
%>
</body>
</html>
