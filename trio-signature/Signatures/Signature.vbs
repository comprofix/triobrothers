' *********************************************************************
' IBS InfoTech
'
' Outlook Signature Editor
'
' Details: Add to the login script for each user to modify a nominated
'   text/html/rtf file with values obtained from the AD user prifile.
'
' 22/06/2011
'
' Ver. 2.0
'
' Version History
' 2.0 - Included support for multiple files and pictures
' 1.0 - Original script
' *********************************************************************

'On Error Resume Next

' Constants
	Const ForReading = 1 ,ForWriting = 2, ForAppending = 8
    Const HKEY_CURRENT_USER = &H80000001
	Const cAppDataPath = "\AppData\Roaming\Microsoft\Signatures\"

' Create objects
	Set Fso = wScript.CreateObject("Scripting.FileSystemObject")
	Set Fso2 = wScript.CreateObject("Scripting.FileSystemObject")
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set WshProcessEnvironment = WshShell.Environment("Process")
	Set objSysInfo = CreateObject("ADSystemInfo")

' Variables
	Dim OfficePhone, OfficeFax, SourcePath, oSourceHTML, oSourceText, oSrcHTM, oSourceRText, oSourceRHTML, sSrcTXT, sUserProfile, sUser, objuser, oFile
	Dim sTitle
	Dim sDescription
	Dim sUID
	Dim sDisplayName
	Dim sFirstName
	Dim sLastName
	Dim sInitials
	Dim sEMail
	Dim sPhoneNumber
	Dim sFaxNumber
	Dim sMobileNumber
	Dim sHomePhone
	Dim sDepartment
	Dim sOfficeLocation
	Dim sWebAddress
	Dim sLastLogOff
	Dim sAddStreet
	Dim sAddCity
	Dim sAddPOBox
	Dim sAddPostCode
	Dim sAddState
	Dim sAddCountry

	' ############ Change Settings Here ############
	LogonServer = "\\trio.local\"
	SourcePath= "\\filesrv.trio.local\Signatures$\"
	SourceImagePath = "\\filesrv.trio.local\Signatures$\TrioSignature_Files"


	sSrcHTM="TrioSignature.htm"
	sSrcTXT="TrioSignature.txt"


	' ############ Change Settings Here ############

	' start running the script
	Call Start

'#####################################################################################################

	Private Sub Start

		' Set the user object
		sUserProfile = WshProcessEnvironment("USERPROFILE")
		sUser = objSysInfo.UserName
		Set objUser = GetObject("LDAP://" & sUser)

		' If the object returned null exit the script
		If isNull(objUser) Then WScript.Quit
		If IsNull(objUser.Displayname) Then WScript.Quit
		If Len(objUser.Displayname) = 0 Then WScript.Quit

		' Assign the user profile values to the variables
		sTitle = objuser.title
		sDescription = objuser.description
		sUID = objuser.cn
		sDisplayName = objuser.displayName
		sFirstName = objuser.givenName
		sLastName = objuser.sn
		sInitials = objuser.initials
		sEMail = objuser.mail
		sPhoneNumber = objuser.telephoneNumber
		sFaxNumber = objuser.facsimileTelephoneNumber
		sMobileNumber = objuser.mobile
		sHomePhone= objUser.homePhone
		sDepartment = objuser.department
		sOfficeLocation = objuser.physicalDeliveryOfficeName
		sWebAddress = objuser.wWWHomePage
		sAddStreet = objuser.streetAddress
		sAddCity = objuser.l
		sAddPOBox = objuser.postOfficeBox
		sAddPostcode = objuser.postalCode
		sAddState = objuser.st
		sAddCountry = objuser.c

		' open the 2 files and set the users' values
		Call OpenAndReplaceFiles

		' Get the paths for the files to be saved to
		Call BuildSignaturePath

		' save the signature files to the path
		Call SaveSignature (oSourceHTML, sUserProfile & cAppDataPath & sSrcHTM)
		Call SaveSignature (oSourceText, sUserProfile & cAppDataPath & sSrcTXT)

		' copy any image files to the path
		Call CopyAdditionalImages(SourceImagePath, sUserProfile & cAppDataPath)


		' Set the new signature files as the default signature for the users' default Outlook profile
		'SetDefaultSignature Left(sSrcHTM,Instr(sSrcHTM,".")-1),""
		'Call SetDefaultSignature(Left(sSrcHTM,Instr(sSrcHTM,".")-1), "")
	'	'Call SetDefaultSignature("TrioSignature", "")

		Set objWord = CreateObject("Word.Application")
		Set objEmailOptions = objWord.EmailOptions
		Set objSignatureObject = objEmailOptions.EmailSignature
		'objSignatureEntries.Add "TrioSignature", objSelection
		objSignatureObject.NewMessageSignature = "TrioSignature"
		objSignatureObject.ReplyMessageSignature = "TrioSignature"
		objWord.Quit
		
		' Exit the script
		WScript.Quit

	End Sub

'#####################################################################################################

	Sub OpenAndReplaceFiles

		' open the specified HTML file for read-only
		Set oFile = Fso.OpenTextFile(SourcePath & sSrcHTM,ForReading,True)
		If ofile.AtEndOfLine = True Then WScript.Echo "oFile = EOF"
		oSourceHTML= oFile.ReadAll
		oFile.Close
		Set oFile = Nothing


		' open the specified text file for read-only
		Set oFile = Fso.OpenTextFile(SourcePath & sSrcTXT,ForReading,True)
		oSourceText= oFile.ReadAll
		oFile.Close
		Set oFile = Nothing


		' Replace variables in the default text file
		oSourceText=replace(oSourceText,"#DISPLAYNAME#",sDisplayName)
		oSourceText=replace(oSourceText,"#TITLE#",sTitle)

		' Phone Numbers
		If len(trim(sPhoneNumber)) = 0 then	oSourceText=replace(oSourceText,"#PHONE#","<b>P:</b> 07 3440 5000 | <b>E:</b> <a href=mailto:" & sEmail & ">" & sEmail & "</a> | <b>W:</b> <a href=https://" & sWebAddress & ">" & sWebAddress & "</a>")'
		If len(trim(sPhoneNumber)) > 0 and len(trim(sMobileNumber)) < 0 then oSourceText=replace(oSourceText,"#PHONE#","P: 07 3440 5000 | D: " & sPhoneNumber & " | E: " & sEmail & " | W: " & sWebAddress)
		If len(trim(sPhoneNumber)) < 0 and len(trim(sMobileNumber)) > 0 then oSourceText=replace(oSourceText,"#PHONE#","P: 07 3440 5000 | M: " & sMobileNumber & " | E: " & sEmail & " | W: " & sWebAddress)
		If len(trim(sPhoneNumber)) > 0 and len(trim(sMobileNumber)) > 0 then oSourceText=replace(oSourceText,"#PHONE#","P: 07 3440 5000 | D: " & sPhoneNumber & " | M: " & sMobileNumber & " | E: " & sEmail & " | W: " & sWebAddress)



		' Replace variables in the Default HTML file
		oSourceHTML=replace(oSourceHTML,"#DISPLAYNAME#",sDisplayName)
		oSourceHTML=replace(oSourceHTML,"#TITLE#",sTitle)

		' Phone Numbers
		If len(trim(sPhoneNumber)) = 0 then	oSourceHTML=replace(oSourceHTML,"#PHONE#","<b>P:</b> 07 3440 5000 | <b>E:</b> <a href=mailto:" & sEmail & ">" & sEmail & "</a> | <b>W:</b> <a href=https://" & sWebAddress & ">" & sWebAddress & "</a>")'
		If len(trim(sPhoneNumber)) > 0 and len(trim(sMobileNumber)) < 0 then oSourceHTML=replace(oSourceHTML,"#PHONE#","<b>P:</b> 07 3440 5000 | <b>D:</b> " & sPhoneNumber & " | <b>E:</b> <a href=mailto:" & sEmail & ">" & sEmail & "</a> | <b>W:</b> <a href=https://" & sWebAddress & ">" & sWebAddress & "</a>")
		If len(trim(sPhoneNumber)) < 0 and len(trim(sMobileNumber)) > 0 then oSourceHTML=replace(oSourceHTML,"#PHONE#","<b>P:</b> 07 3440 5000 | <b>M:</b> " & sMobileNumber & " | <b>E:</b> <a href=mailto:" & sEmail & ">" & sEmail & "</a> | <b>W:</b> <a href=https://" & sWebAddress & ">" & sWebAddress & "</a>")
		If len(trim(sPhoneNumber)) > 0 and len(trim(sMobileNumber)) > 0 then oSourceHTML=replace(oSourceHTML,"#PHONE#","<b>P:</b> 07 3440 5000 | <b>D:</b> " & sPhoneNumber & " | <b>M:</b> " & sMobileNumber & " | <b>E:</b> <a href=mailto:" & sEmail & ">" & sEmail & "</a> | <b>W:</b> <a href=https://" & sWebAddress & ">" & sWebAddress & "</a>")



	End Sub


'#####################################################################################################

	Sub BuildSignaturePath

		' Create Base Folders if they do not exist

		If not fso.folderexists(sUserProfile & "\AppData") then
			fso.createFolder(sUserProfile & "\AppData")
		End If

		If not fso.folderexists(sUserProfile & "\AppData\Roaming") then
			fso.createFolder(sUserProfile & "\AppData\Roaming")
		End If

		If not fso.folderexists(sUserProfile & "\AppData\Roaming\Microsoft") then
			fso.createFolder(sUserProfile & "\AppData\Roaming\Microsoft")
		End If

	 	If not fso.folderexists(sUserProfile & "\AppData\Roaming\Microsoft\Signatures") then
			fso.createFolder(sUserProfile & "\AppData\Roaming\Microsoft\Signatures")
		End if

	End Sub

'#####################################################################################################

	Function SaveSignature(SignatureText,SignaturePath)

		' write the file to the given path

	  	Set ca = fso.CreateTextFile(SignaturePath, ForWriting, True)

	  	ca.write(SignatureText)
	 	ca.close
		set ca=nothing

	End Function

'#####################################################################################################

	Function CopyAdditionalImages(SourcePath, SignaturePath)

		' copy the images from the source path to the destination
		fso.copyFolder SourcePath, SignaturePath
		On Error resume Next

		fso.copyFolder SourcePath, SignaturePath

		'fso.copyFile SourcePath & "*.png", SignaturePath
		'fso.copyFile SourcePath & "*.gif", SignaturePath

	End Function

'#####################################################################################################

	Public Function StringToByteArray (Data, NeedNullTerminator)

	    Dim strAll

	    strAll = StringToHex4(Data)

	    If NeedNullTerminator Then
	        strAll = strAll & "0000"
	    End If

	    intLen = Len(strAll) \ 2
	    ReDim arr(intLen - 1)

	    For i = 1 To Len(strAll) \ 2
	        arr(i - 1) = CByte("&H" & Mid(strAll, (2 * i) - 1, 2))
	    Next

	    StringToByteArray = arr

	End Function

'#####################################################################################################

	Public Function StringToHex4(Data)

	    ' Input: normal text
	    ' Output: four-character string for each character,
	    '         e.g. "3204" for lower-case Russian B,
	    '        "6500" for ASCII e
	    ' Output: correct characters
	    ' needs to reverse order of bytes from 0432

	    Dim strAll

	    For i = 1 To Len(Data)
	        ' get the four-character hex for each character
	        strChar = Mid(Data, i, 1)
	        strTemp = Right("00" & Hex(AscW(strChar)), 4)
	        strAll = strAll & Right(strTemp, 2) & Left(strTemp, 2)
	    Next

	    StringToHex4 = strAll

	End Function
