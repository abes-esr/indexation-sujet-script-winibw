'(Declarations)


Dim IE_1
Dim ScriptEngine1
Dim ScriptEngine2

' Paramètre resultsCount qui modifie le nombre de résultats de chaque modèle IA sous-jacent.
' Ce paramètre ne joue pas mécaniquement sur le nombre de suggestions affichées, 
' car la liste des suggestions dépend d'une phase de sélection parmi les résultats des modèles, selon deux facteurs complémentaires :
' 1/ une autre IA sélectionne les résultats renvoyés par les modèles selon leur pertinence ;
' 2/ les résultats communs à au moins deux modèles sont sélectionnés.
Dim resultsCount
resultsCount = "6"

set IE1 = nothing
set IE2 = nothing

set doc1 = nothing
Set ScriptEngine1= nothing
Set ScriptEngine2= nothing



	
	
Function AddArrayElement(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
End Function

Function encodeURL1(str)
    'Dim ScriptEngine As Object
    'Dim encoded As String
    encoded = ScriptEngine1.Run("encodeURIComponent", str)

    encodeURL1 = encoded
End Function

Function getTagDollar(tag,dollar) 
	'200$ale monde
	'330$ala vie
	Dim returnValue
	'returnValue="Rien en "& tag & " $" & dollar
	returnValue=""
	Application.ActiveWindow.title.startOfBuffer (false)
	res = Application.ActiveWindow.title.findTag(tag, 0, true, true, false)
	sousZones=Split(res,"$")
	for each x in sousZones
		'MsgBox "sous zones :" & x
		if InStr(x,dollar)=1 then
			'MsgBox "ok :" & dollar & Mid(x,2) ok
			returnValue = Replace(returnValue, ".", "") & ". " & Mid(x,2) 
		end if
		
	next
	
	getTagDollar= returnValue
End Function

Sub Letitbe()
' Fonction vide pour contourner un bug de WinIBW : copier/coller cette fonction avec tout le reste du code ci-dessus
End Sub

Function z606withoutDollarsax() 
' Fonction qui supprime tous les $a et $x des 606 ayant un $3 par ailleurs
' Ne fonctionne par pour les lignes commençant par $3 (= les conceopts à employer comme subdivision seulement) : logique, ils sont censés être déplacés en fin d'une ligne 606
	For i = 0 To 20				' en supposant un maximum de 21 606...
		Dim retrait
		retrait = ""
		For ii = 0 To 5			' Pour prendre en compte l'affichage hiérarchisé, avec un retrait pour chaque niveau enfant
			'MsgBox i
			Dim returnValue
			returnValue=""
			'Application.ActiveWindow.title.startOfBuffer (false)
			res = Application.ActiveWindow.title.findTag(retrait&"606", i, true, true, false)
			'MsgBox res
			if InStr(res,"$3") > 0 then
				'MsgBox "$3 présent"
				if InStr(res,"$a") > 0 OR InStr(res,"$x") > 0 then
					'MsgBox "$a ou $x présent"
					'MsgBox res
					sousZones=Split(res,"$")
					for each sousZone in sousZones
						'MsgBox "sous zones :" & sousZone & "**"
						if InStr(Trim(sousZone),"6") = 1 then		' si le premier item du split commence par 6 (après suppression des esapces devant derrière)
							'MsgBox "Commence par 6"
							returnValue = returnValue + sousZone
						elseif InStr(sousZone,"a") = 1 OR InStr(sousZone,"x") = 1 then
							'MsgBox "$a ou $x"
							returnValue = returnValue + "" 
						else
							returnValue = returnValue + "$" + sousZone
						end if
					Next
				'MsgBox returnValue
				Application.ActiveWindow.title.insertText(returnValue)
				end if
			end if
			retrait = retrait + "   "        ' Retrait (trois espaces) ajouté à chaque niveau de la hiérarchie
		Next
	Next
	Application.ActiveWindow.title.endOfBuffer (false)
End Function

Sub indexationRameau_html_url()
' Renvoie les suggestions dans une page web, extérieure à WinIBW
	if Application.ActiveWindow.title is nothing then
		Application.ActiveWindow.Command "mod", False
	Else
		Dim agent
		agent = Replace(application.activeWindow.Variable("P3GUL")," ","")
		
		Dim PPN_aindexer
		PPN_aindexer = Replace(application.activeWindow.Variable("P3GPP")," ","")
	     				
		if ScriptEngine1 is nothing then
		    	Set ScriptEngine1 = CreateObject("scriptcontrol")
	    		ScriptEngine1.Language = "JScript"
	    	end if
	    		
		Dim resumee 
		resumee= application.activeWindow.title.selection
		resumee= Replace(resumee,"'","\x27")
		resumee= Replace(resumee,"&","\x26")
		'MsgBox "resumee:"&resumee
		Dim titre 
		titre = Replace(getTagDollar("200","a"),"'","\x27")
		titre = Replace(titre,"&","\x26")
		Application.ActiveWindow.title.startOfBuffer (false)
		Dim sstitre 
		sstitre = Replace(getTagDollar("200","e"),"'","\x27")
		sstitre = Replace(sstitre,"&","\x26")
		Application.ActiveWindow.title.startOfBuffer (false)
		Dim titrepartie 
		titrepartie = Replace(getTagDollar("200","i"),"'","\x27")
		titrepartie = Replace(titrepartie ,"&","\x26")
		Application.ActiveWindow.title.startOfBuffer (false)
		Dim titres 
		titres = titre & ". " & sstitre & " " + titrepartie 
		titres = Replace(titres,"@","")
		MsgBox  titres 

		' Construit et lance l'url dans le navifateur par defaut du poste
		url ="https://hub-rdftotal.idref.fr/subject_indexation/?docId="+PPN_aindexer+"&Title="+encodeURL1(titres)+"&Summary="+encodeURL1(resumee)+"&models=victor3_chain,victor1_concept,victor2&aggregationType=intersection2models,llm&subjectsMaxCount="+resultsCount+"&Agent="+agent+"&vocabulary=rameau&Format=html"
		
		Application.ShellExecute url, 5, "open", ""
	    	
		
	End If
End Sub

Sub indexationRameau_wsh()
' Renvoie les suggestions dans WinIBW, sans dépendance à Internet Explorer
	Dim ScriptEngine3
	Set ScriptEngine3= nothing

	'MsgBox resultsCount

	if Application.ActiveWindow.title is nothing then
		Application.ActiveWindow.Command "mod", False
	Else
		Dim agent
		agent = Replace(application.activeWindow.Variable("P3GUL")," ","")
		
		Dim PPN_aindexer
		PPN_aindexer = Replace(application.activeWindow.Variable("P3GPP")," ","")
	     	
	     	
	    		
		Dim resumee 
		resumee= application.activeWindow.title.selection
		resumee= Replace(resumee,"'","\x27")
		resumee= Replace(resumee,"&","\x26")
		'MsgBox "resumee:"&resumee
		Dim titre 
		titre = Replace(getTagDollar("200","a"),"'","\x27")
		titre = Replace(titre,"&","\x26")
		Application.ActiveWindow.title.startOfBuffer (false)
		Dim sstitre 
		sstitre = Replace(getTagDollar("200","e"),"'","\x27")
		sstitre = Replace(sstitre,"&","\x26")
		Application.ActiveWindow.title.startOfBuffer (false)
		Dim titrepartie 
		titrepartie = Replace(getTagDollar("200","i"),"'","\x27")
		titrepartie = Replace(titrepartie ,"&","\x26")
		Application.ActiveWindow.title.startOfBuffer (false)
		Dim titres 
		titres = titre & ". " & sstitre & " " + titrepartie 
		titres = Replace(titres,"@","")
		'MsgBox  titres 
    	
    		if ScriptEngine3 is nothing then
		    	Set ScriptEngine3 = CreateObject("MSScriptControl.ScriptControl")
	    		ScriptEngine3.Language = "JScript"
	    	end if

    	
		js= "var xhr = new ActiveXObject(""Microsoft.XMLHTTP"");  " 				
		js= js & " function prod1(){"
	    	js= js & "titre= encodeURIComponent('" & titres  & "');"
		js=js & "abstract= encodeURIComponent('" & resumee  &"');"
		js=js & "agent = encodeURIComponent('" & agent  & "');"
		js=js & "PPN_aindexer = encodeURIComponent('" & PPN_aindexer & "');"
		' new
		js=js & "subjectsMaxCount = encodeURIComponent('" & resultsCount & "');"
		js=js & "error_srv= '';"
		js=js & "url ='https://hub-rdftotal.idref.fr/subject_indexation/?docId='+PPN_aindexer+'&Title='+titre+'&Summary='+abstract+'&models=victor3_chain,victor1_concept,victor2&aggregationType=intersection2models,llm&subjectsMaxCount='+subjectsMaxCount+'&Agent='+agent+'&vocabulary=rameau&Format=text';"
		js=js & " try { "
		js=js & "xhr.open('GET', url,false); "
		js=js & "xhr.setRequestHeader('Content-Type','application/x-www-form-urlencoded; charset=ISO-8859-1'); "
		js=js & " xhr.send(); "
		js=js & " } catch (error) {  error_srv ='sur send:'+error.message;return error_srv;} "
				
		js= js & "return xhr.responseText}"


		With ScriptEngine3
		        .AddCode (js)
		        result = .Run("prod1")
		End With
		MsgBox result
		Application.ActiveWindow.title.endOfBuffer (false)
		on error resume next
		

		Application.ActiveWindow.title.insertText(vbCrLf)
		Application.ActiveWindow.title.insertText("... Suggestions : " & vbCrLf)
		Application.ActiveWindow.title.insertText(result)
		
   End If
   Set ScriptEngine3= nothing
End Sub

