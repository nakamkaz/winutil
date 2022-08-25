' Usage : cscript /nologo //B filename.vbs  outfile

Set updateSession = CreateObject("Microsoft.Update.Session")
updateSession.ClientApplicationID = "FINDNOTINST"
Set updateSearcher = updateSession.CreateupdateSearcher()
Set searchResult1 = _
updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

Set fso = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1, ForWriting = 2, ForAppending = 8 
Set citext = fso.OpenTextFile(WScript.Arguments.Item(0) & ".txt", ForWriting, true, true)

citext.WriteLine "Reported: " & Now

Dim pkgBoth

For i = 0 To searchResult1.Updates.Count-1
	Set update = searchResult1.Updates.Item(i)
	If update.AutoSelectOnWebSites Then 
		Dim categs1
		categs1=""
		Dim hasUtype1
		hasUtype1 ="*"

		For x = 0 To upateCategories.Count-1
			Set categ1 = update.Categories.Item(x)
			categs1=categs1 & "{" & categ1.Name & "} "
			If InStr(categ1.Name,"On Demand") > 0 Then
				hasUtype1 ="-"
			ElseIf InStr(categ1.Name,"Language") > 0 Then
				hasUtype1 ="L"
			ElseIf InStr(categ1.Name,"Critical Updates") > 0 Then
				hasUtype1 ="C"
			ElseIf InStr(categ1.Name,"Security Updates") > 0 Then
				hasUtype1 ="S"
			ElseIf InStr(categ1.Name,"Defender") > 0 Then
				hasUtype1 ="D"
			ElseIf InStr(categ1.Name,"Silverlight") > 0 Then
				hasUtype1 ="W"
			End If
		Next

		If hasUtype1 <> "-" Then
			citext.Write "ASoWS=" & update.AutoSelectOnWebSites & ", "
			citext.Write update.LastDeploymentChangeTime & ", "
			citext.Write hasUtype1 & ", "
			For Each sbid in  update.SecurityBulletinIDs
				If Instr(sbid,"MS") > 0 Then
					citext.Write sbid & " "
					pkgBoth = pkgBoth + "|"+sbid
				End If
			Next
			For Each strKB in  update.KBArticleIDs
				If NOT IsNull(strKB)  Then
					citext.Write strKB & " "
					pkgBoth = pkgBoth + "|"+strKB
				End If
			Next

			citext.Write ", "
			citext.Write categs1
			citext.Write ", " 
			citext.Write update.Title
			citext.WriteLine ""
		End If
	End If

Next

citext.WriteLine ""


For i = 0 To searchResult1.Updates.Count-1
	Set update = searchResult1.Updates.Item(i)
	If NOT update.AutoSelectOnWebSites Then 
		Dim categs2
		categs2=""
		Dim hasUtype2
		hasUtype2 ="*"

		For x = 0 To upateCategories.Count-1
			Set categ2=update.Categories.Item(x)
			categs2=categs2 & "{" & categ2.Name & "} "
			If InStr(categ2.Name,"On Demand") > 0 Then
				hasUtype2 ="-"
			ElseIf InStr(categ2.Name,"Language") > 0 Then
				hasUtype2 ="L"
			ElseIf InStr(categ2.Name,"Critical Updates") > 0 Then
				hasUtype2 ="C"
			ElseIf InStr(categ2.Name,"Security Updates") > 0 Then
				hasUtype2 ="S"
			ElseIf InStr(categ2.Name,"Defender") > 0 Then
				hasUtype2 ="D"
			ElseIf InStr(categ2.Name,"Silverlight") > 0 Then
				hasUtype2 ="W"
			End If
		Next

		If hasUtype2 <> "-" Then
			citext.Write "ASoWS=" & update.AutoSelectOnWebSites & ", "
			citext.Write update.LastDeploymentChangeTime & ", "
			citext.Write hasUtype2 & ", "
			For Each sbid in  update.SecurityBulletinIDs
				If Instr(sbid,"MS") > 0 Then
					citext.Write sbid & " "
					pkgBoth = pkgBoth + "|"+sbid
				End If
			Next
			For Each strKB in  update.KBArticleIDs
				If NOT IsNull(strKB)  Then
					citext.Write strKB & " "
					pkgBoth = pkgBoth + "|"+strKB
				End If
			Next

			citext.Write ", "
			citext.Write categs2
			citext.Write ", " 
			citext.Write update.Title
			citext.WriteLine ""
		End If
	End If
Next


citext.Write "Query-PACKGE: "
citext.Write pkgBoth
citext.WriteLine ""
citext.Close()
