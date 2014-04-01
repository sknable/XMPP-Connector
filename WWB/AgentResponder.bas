Attribute VB_Name = "Module1"
'//--------------------------------------------------------------------
'// Copyright (c) 2011 Syntellect, Inc.
'//
'// FILE: $Workfile:   AgentResponder.bas  $
'//
'// DESCRIPTION: This responder script provides various utilities including the ability to define
'// wrapup codes, perform phone number modification, and perform phone number validation. These
'// abilities are handled when the associated VBA Server receives a responder request that is in
'// the form of a dataset.
'//--------------------------------------------------------------------

'// A CallTypeEnum specifies the type of call being performed.
Enum CallTypeEnum
	CALLTYPE_UNKNOWN
	CALLTYPE_CALL
	CALLTYPE_CONSULTATION
	CALLTYPE_TRANSFER
	CALLTYPE_CONFERENCE
	CALLTYPE_SECOND_CALL
	CALLTYPE_CONSULTANT
End Enum

'// A DialTargetEnum specifies the type of object being dialed.
Enum DialTargetEnum
	DIALTARGET_UNKNOWN
	DIALTARGET_AGENT
	DIALTARGET_CONTACT
	DIALTARGET_QUEUE
End Enum

'// Encapsulates agent related data.
Type AgentDataType
	ConfigGroupID As String
	ConfigGroupName As String
	EmailAddress As String
	Extension As String
	LinkName As String
	UserID As String
	UserName As String
End Type

'// Encapsulates interaction information.
Type InteractionType
	DisplayInfo As String
	Disposition As String
	IID As String
	IsCallOwner As Boolean
	IType As String
	Name As String
	OriginatingIID As String
	ParentIID As String
	PopData As String
	Queue As String
	QueueID As String
	SingleAgent As String
	State As String
	SourceProcess As String
	StandardInteractionProperties As Object
	CustomInteractionProperties As Object
End Type

'// Encapsulates phone related data.
Type PhoneDataType
	CallType As CallTypeEnum
	ConfigGroupName As String
	DialTarget As DialTargetEnum
	LinkName As String
	UserExtension As String
End Type

'// The various contexts that this script handles.
Const CONTEXT_TOKEN$ = "CONTEXT"
Const GET_WRAPUP_CODES_TOKEN$ = "GET_WRAPUP_CODES"
Const VALIDATE_PHONE_NUMBER_TOKEN$ = "VALIDATE_PHONE_NUMBER"
Const MODIFY_PHONE_NUMBER_TOKEN$ = "MODIFY_PHONE_NUMBER"

'// Used for wrapup codes retrieval handling.
Const ID_TOKEN$ = "ID"
Const DISPLAY_NAME_TOKEN$ = "DISPLAY_NAME"
Const LOCALE_TOKEN$ = "LOCALE"
Const SELECTED_TOKEN$ = "SELECTED"
Const WRAPUP_CODE_TOKEN$ = "WRAPUP_CODE"
Const WRAPUP_CODES_TOKEN$ = "WRAPUP_CODES"

'// Used for when retrieving selection options for wrapup codes.
Const UID_TOKEN$ = "UID"
Const IID_TOKEN$ = "IID"
Const SINGLE_SELECT_TOKEN$ = "SINGLE_SELECT"
Const REQUIRE_SELECTION_TOKEN$ = "REQUIRE_SELECTION"

'// Used for phone validation handling.
Const PHONE_NUMBER_TOKEN$ = "PHONE_NUMBER"
Const VALID_PHONE_NUMBER_TOKEN$ = "VALID_PHONE_NUMBER"
Const INVALID_REASON_TOKEN$ = "INVALID_REASON"

'// Used for when building interaction properties from an xml string.
Const INTERACTION_PROPERTY_TOKEN$ = "INTERACTION_PROPERTY"
Const KEY_TOKEN$ = "KEY"
Const VALUE_TOKEN$ = "VALUE"

'// Used for when retrieving interaction related information from a dataset.
Const CUSTOM_INTERACTION_PROPERTIES_TOKEN$ = "CUSTOM_INTERACTION_PROPERTIES"
Const DISPLAY_INFO_TOKEN$ = "DISPLAY_INFO"
Const DISPOSITION_TOKEN$ = "DISPOSITION"
Const IS_CALL_OWNER_TOKEN$ = "IS_CALL_OWNER"
Const ITYPE_TOKEN$ = "ITYPE"
Const NAME_TOKEN$ = "NAME"
Const ORIGINATING_IID_TOKEN$ = "ORIGINATING_IID"
Const PARENT_IID_TOKEN$ = "PARENT_IID"
Const POP_DATA_TOKEN$ = "POP_DATA"
Const QUEUE_ID_TOKEN$ = "QUEUE_ID"
Const QUEUE_NAME_TOKEN$ = "QUEUE_NAME"
Const SINGLE_AGENT_UID_TOKEN$ = "SINGLE_AGENT_UID"
Const SOURCE_PROCESS_TOKEN$ = "SOURCE_PROCESS"
Const STANDARD_INTERACTION_PROPERTIES_TOKEN$ = "STANDARD_INTERACTION_PROPERTIES"
Const STATE_TOKEN$ = "STATE"

'// Used for when agent related information from a dataset.
Const CONFIG_GROUP_ID_TOKEN$ = "CONFIG_GROUP_ID"
Const EMAIL_ADDRESS_TOKEN$ = "EMAIL_ADDRESS"
Const EXTENSION_TOKEN$ = "EXTENSION"
Const USER_NAME_TOKEN$ = "USER_NAME"

'// String equivalents of call types.
Const CALLTYPE_UNKNOWN_TOKEN$ = "UNKNOWN"
Const CALLTYPE_CALL_TOKEN$ = "CALL"
Const CALLTYPE_CONSULTATION_TOKEN$ = "CONSULTATION"
Const CALLTYPE_TRANSFER_TOKEN$ = "TRANSFER"
Const CALLTYPE_CONFERENCE_TOKEN$ = "CONFERENCE"
Const CALLTYPE_SECOND_CALL_TOKEN$ = "SECOND_CALL"
Const CALLTYPE_CONSULTANT_TOKEN$ = "CONSULTANT"

'// String equivalents of dial targets.
Const DIALTARGET_UNKNOWN_TOKEN$ = "UNKNOWN"
Const DIALTARGET_AGENT_TOKEN$ = "AGENT"
Const DIALTARGET_CONTACT_TOKEN$ = "CONTACT"
Const DIALTARGET_QUEUE_TOKEN$ = "QUEUE"

'// Used for when retrieving phone related information from a dataset.
Const CALL_TYPE_TOKEN$ = "CALL_TYPE"
Const CONFIG_GROUP_NAME_TOKEN$ = "CONFIG_GROUP_NAME"
Const DIAL_TARGET_TOKEN$ = "DIAL_TARGET"
Const LINK_NAME_TOKEN$ = "LINK_NAME"
Const USER_EXTENSION_TOKEN$ = "USER_EXTENSION"

Dim gAllWrapupCodes As Object
Dim gAllWrapupCodesSettings As Object

'//--------------------------------------------------------------------
'// PROC: SetProcType
'//
'// DESCRIPTION:
'// This procedure is optional. It is called by the C++ program before
'// registering with the router. It determines the types of messages
'// which will be received. If it is not present, the default is 'PP',
'// which means that the server will only receive PP messages in which
'// it is the destination.
'//--------------------------------------------------------------------
Public Sub SetProcType()

    '// SetProcessType is an extended function (part of ConnectionType)
    SetProcessType ("PPA")

End Sub

'//--------------------------------------------------------------------
'// PROC: Initialize
'//
'// DESCRIPTION:
'// This procedure is required. It is called after registering with
'// the router, or after pressing the Restart button. Its primary
'// purpose is to initialize global variables which will be used in
'// the script.
'//--------------------------------------------------------------------
Public Sub Initialize()
	
End Sub

'//--------------------------------------------------------------------
'// PROC: HandleMessage
'//
'// DESCRIPTION:
'// This procedure is required.
'// It is called each time a message other than RESPONDER_REQUEST
'// or RESPONDER_RESPONSE is received.
'//--------------------------------------------------------------------
Public Sub HandleMessage(Msg As MessageType)

End Sub

'//--------------------------------------------------------------------
'// PROC: Terminate
'//
'// DESCRIPTION:
'// This procedure is required. It is called just before shutting down
'// the router ConnectionType and quitting the application. Put your
'// clean-up code here.
'//--------------------------------------------------------------------
Public Sub Terminate()

End Sub

'//--------------------------------------------------------------------
'// PROC: ReceivedRequestAsDataSet
'//
'// DESCRIPTION:
'// This procedure is called when the VBA Server receives a RESPONDER_REQUEST 
'// containing a dataset. Depending on the context defined in the passed
'// in dataset, it will handle one of three actions: retrieval of wrapup
'// codes, phone number modification, or phone number validation. Once
'// the desired action is handled, a result dataset is constructed and
'// populated with context specific information. The result dataset is
'// then propogated back to the sender of the RESPONDER_REQUEST in the
'// form of a RESPONDER_RESPONSE.
'//--------------------------------------------------------------------
Sub ReceivedRequestAsDataSet(id As String, src As APISource, dataset As APIDataSet)
    
    '// Determine the context of the passed in dataset. At the moment, there are 3 actions that
    '// this script will handle: retrieval of wrapup codes, phone number modification, and phone
    '// number validation.
    Dim context As String
    context = GetField(CONTEXT_TOKEN, dataset)
    Trace("Context: " & context)
    
    If context = GET_WRAPUP_CODES_TOKEN$ Then
    	HandleGetWrapupCodes(id, src, dataset)
    ElseIf context = MODIFY_PHONE_NUMBER_TOKEN$ Then
    	HandleModifyPhoneNumber(id, src, dataset)
    ElseIf context = VALIDATE_PHONE_NUMBER_TOKEN$ Then
    	HandleValidatePhoneNumber(id, src, dataset)
    End If
     
End Sub

'//--------------------------------------------------------------------
'// PROC: AddWrapupCode
'//
'// DESCRIPTION:
'// Calling this procedure adds a wrapup code to memory.
'//
'// PARAMETERS:
'//  name:
'//    The internal name of the wrapup code.
'//
'//  displayName:
'//    The display name of the wrapup code.
'//
'//  locale:
'//    The locale of the wrapup code (e.g., en_US).
'//
'//  selected:
'//    true if the wraup code should be initially selected when displayed
'//    to a user and false, otherwise.
'//--------------------------------------------------------------------
Sub AddWrapupCode(name As String, displayName As String, locale As String, selected As Boolean)

	If gAllWrapupCodes Is Nothing Then
		Set gAllWrapupCodes = CreateObject("Scripting.Dictionary")
	End If
	
	'// Initialize new wrapup code info and then add that to our collection.
	Dim wrapupCode As Object
	Set wrapupCode = CreateObject("Scripting.Dictionary")
	wrapupCode.Add(ID_TOKEN$, name)
	wrapupCode.Add(DISPLAY_NAME_TOKEN$, displayName)
	wrapupCode.Add(LOCALE_TOKEN$, locale)
	wrapupCode.Add(SELECTED_TOKEN$, selected)
	
	'// Note that if a wrapup code already exists for the passed in name, we'll
	'// overwrite it with the new wrapup code.
	If gAllWrapupCodes.Exists(name) = True Then
		Set gAllWrapupCodes.Item(name) = wrapupCode
	Else
		gAllWrapupCodes.Add(name, wrapupCode)
	End If
	
End Sub

'//--------------------------------------------------------------------
'// PROC: BuildInteractionProperties
'//
'// DESCRIPTION:
'// Reads the passed in xml string representation of interaction
'// properties and returns a Dictionary collection that maps interaction 
'// property keys to their values. The xml string should assume the
'// following format,
'//
'//     <?xml version="1.0"?>
'//     <INTERACTION_PROPERTIES>
'//         <INTERACTION_PROPERTY>
'//             <KEY>...</KEY>
'//             <VALUE>...</VALUE>
'//         </INTERACTION_PROPERTY>
'//     </INTERACTION_PROPERTIES>
'//
'//--------------------------------------------------------------------
Sub BuildInteractionProperties(xmlString As String, interactionProperties As Object)

	Trace(xmlString)
	If Not(interactionProperties Is Nothing) Then
	
		'// Load the xml string in order to generate an xml tree.
		Dim xmlDoc As Object
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
		xmlDoc.async = False
		xmlDoc.loadXML(xmlString)

		'// Retrieve a collection of nodes with the tag, INTERACTION_PROPERTY.
		Dim interactionPropertyNodes As Object
		Set interactionPropertyNodes = xmlDoc.getElementsByTagName(INTERACTION_PROPERTY_TOKEN$)

		'// Recurse through the INTERACTION_PROPERTY nodes in order to retrieve the key/value pair in each.
		Dim interactionPropertyNode As Object
		For Each interactionPropertyNode In interactionPropertyNodes

			'// Retrieve the key/value pair of the interaction property node.
			Dim key As String
			Dim value As String
			RetrieveKeyValuePair(interactionPropertyNode.childNodes, key, value)

			'// We want to make sure the key is not empty.
			If (Len(Trim(key)) > 0) Then
				
				'// If the key already exists, overwrite the old value with the new value.
				If interactionProperties.Exists(key) = True Then
					interactionProperties.Remove(key)
				End If
				interactionProperties.Add(key, value)
			End If

		Next interactionPropertyNode
		
		'// Release the xml DOM.
		Set xmlDoc = Nothing
	
	End If

End Sub

'//--------------------------------------------------------------------
'// PROC:  GetCallTypeEnum
'//
'// DESCRIPTION:
'// Returns the equivalent CallTypeEnum of the passed in string. Note
'// that if no match is found for the passed in string the CallTypeEnum,
'// CALLTYPE_UNKNOWN, is returned.
'//--------------------------------------------------------------------
Function GetCallTypeEnum(callTypeStr As String) As CallTypeEnum
	
	Dim callType As CallTypeEnum
	callType = CALLTYPE_UNKNOWN
	
	'// Match the passed in string to a CallTypeEnum.
	If callTypeStr = CALLTYPE_UNKNOWN_TOKEN$ Then
		callType = CALLTYPE_UNKNOWN
	ElseIf callTypeStr = CALLTYPE_CALL_TOKEN$ Then
		callType = CALLTYPE_CALL
	ElseIf callTypeStr = CALLTYPE_CONSULTATION_TOKEN$ Then
		callType = CALLTYPE_CONSULTATION
	ElseIf callTypeStr = CALLTYPE_TRANSFER_TOKEN$ Then
		callType = CALLTYPE_TRANSFER
	ElseIf callTypeStr = CALLTYPE_CONFERENCE_TOKEN$ Then
		callType = CALLTYPE_CONFERENCE
	ElseIf callTypeStr = CALLTYPE_SECOND_CALL_TOKEN$ Then
		callType = CALLTYPE_SECOND_CALL
	ElseIf callTypeStr = CALLTYPE_CONSULTANT_TOKEN$ Then
		callType = CALLTYPE_CONSULTANT
	End If
	
	GetCallTypeEnum = callType

End Function

'//--------------------------------------------------------------------
'// PROC:  GetCallTypeEnum
'//
'// DESCRIPTION:
'// Returns the string equivalent of the passed in CallTypeEnum.
'//--------------------------------------------------------------------
Function GetCallTypeEnumAsString(callType As CallTypeEnum) As String

	Dim callTypeStr As String
	
	'// Match the passed in CallTypeEnum to a string.
	If callType = CALLTYPE_UNKNOWN Then
		callTypeStr = CALLTYPE_UNKNOWN_TOKEN$
	ElseIf callType = CALLTYPE_CALL Then
		callTypeStr = CALLTYPE_CALL_TOKEN$
	ElseIf callType = CALLTYPE_CONSULTATION Then
		callTypeStr = CALLTYPE_CONSULTATION_TOKEN$
	ElseIf callType = CALLTYPE_TRANSFER Then
		callTypeStr = CALLTYPE_TRANSFER_TOKEN$
	ElseIf callType = CALLTYPE_CONFERENCE Then
		callTypeStr = CALLTYPE_CONFERENCE_TOKEN$
	ElseIf callType = CALLTYPE_SECOND_CALL Then
		callTypeStr = CALLTYPE_SECOND_CALL_TOKEN$
	ElseIf callType = CALLTYPE_CONSULTANT Then
		callTypeStr = CALLTYPE_CONSULTANT_TOKEN$
	End If
	
	GetCallTypeEnumAsString = callTypeStr
	
End Function

'//--------------------------------------------------------------------
'// PROC:  GetField
'//
'// DESCRIPTION:
'// It is assumed that the passed in dataset has 2 columns. The first
'// column of each row should contain a string token that describes the 
'// value of the field in the second column of that same
'// row. This method will search the first column of each row in the
'// passed in dataset until it finds the passed in token. If the token is
'// found, the value in the second column of that same row is returned.
'//--------------------------------------------------------------------
Function GetField(token As String, dataset As APIDataSet) As String

	'// Ensure that the passed in dataset has 2 column before proceeding...
	Dim field As String
	If dataset.GetColumns() = 2 Then
		
		'// Iterate through each row in the dataset until we find a row that
		'// contains the passed in token.
		Dim rowIndex As Integer
		Dim rows As Integer
		rows = dataset.GetRows()
		For rowIndex = 0 To rows - 1
			
			'// Has the token been found?
			If dataset.GetField(rowIndex, 0) = token Then
			
				'// Retrieve the value at column 2 of the current row.
				field = dataSet.GetField(rowIndex, 1)
				Exit For
			
			End If
			
		Next
	
	End If
	
	GetField = field

End Function

'//--------------------------------------------------------------------
'// PROC: GetRequireSelectionValue
'//
'// DESCRIPTION:
'// Returns the value of the require selection setting. A default value 
'// of false is returned if a value was never set.
'//--------------------------------------------------------------------
Function GetRequireSelectionValue() As Boolean
	
	Dim requireSelection As Boolean
	requireSelection = False
	
	If Not(gAllWrapupCodesSettings Is Nothing) Then
		If gAllWrapupCodesSettings.Exists(REQUIRE_SELECTION_TOKEN$) Then
			requireSelection = gAllWrapupCodesSettings.Item(REQUIRE_SELECTION_TOKEN$)
		End If
	End If
	
	GetRequireSelectionValue = requireSelection
	
End Function

'//--------------------------------------------------------------------
'// PROC: GetSingleSelectValue
'//
'// DESCRIPTION:
'// Returns the value of the single select setting. A default value of 
'// false is returned if a value was never set.
'//--------------------------------------------------------------------
Function GetSingleSelectValue() As Boolean
	
	Dim singleSelect As Boolean
	singleSelect = False
	
	If Not(gAllWrapupCodesSettings Is Nothing) Then
		If gAllWrapupCodesSettings.Exists(SINGLE_SELECT_TOKEN$) Then
			singleSelect = gAllWrapupCodesSettings.Item(SINGLE_SELECT_TOKEN$)
		End If
	End If
	
	GetSingleSelectValue = singleSelect
	
End Function

'//--------------------------------------------------------------------
'// PROC:  GetDialTargetEnum
'//
'// DESCRIPTION:
'// Returns the equivalent DialTargetEnum of the passed in string. Note
'// that if no match is found for the passed in string the DialTargetEnum,
'// DIALTARGET_UNKNOWN, is returned.
'//--------------------------------------------------------------------
Function GetDialTargetEnum(dialTargetStr As String) As DialTargetEnum
	
	Dim dialTarget As DialTargetEnum
	dialTarget = DIALTARGET_UNKNOWN
	
	'// Match the passed in string to a DialTargetEnum.
	If dialTargetStr = DIALTARGET_UNKNOWN_TOKEN$ Then
		dialTarget = DIALTARGET_UNKNOWN
	ElseIf dialTargetStr = DIALTARGET_AGENT_TOKEN$ Then
		dialTarget = DIALTARGET_AGENT
	ElseIf dialTargetStr = DIALTARGET_CONTACT_TOKEN$ Then
		dialTarget = DIALTARGET_CONTACT
	ElseIf dialTargetStr = DIALTARGET_QUEUE_TOKEN$ Then
		dialTarget = DIALTARGET_QUEUE
	End If
	
	GetDialTargetEnum = dialTarget

End Function

'//--------------------------------------------------------------------
'// PROC:  GetDialTargetEnumAsString
'//
'// DESCRIPTION:
'// Returns the string equivalent of the passed in DialTargetEnum.
'//--------------------------------------------------------------------
Function GetDialTargetEnumAsString(dialTarget As DialTargetEnum) As String

	Dim dialTargetStr As String
	
	'// Match the passed in DialTargetEnum to a string.
	If dialTarget = DIALTARGET_UNKNOWN Then
		dialTargetStr = DIALTARGET_UNKNOWN_TOKEN$
	ElseIf dialTarget = DIALTARGET_AGENT Then
		dialTargetStr = DIALTARGET_AGENT_TOKEN$
	ElseIf dialTarget = DIALTARGET_CONTACT Then
		dialTargetStr = DIALTARGET_CONTACT_TOKEN$
	ElseIf dialTarget = DIALTARGET_QUEUE Then
		dialTargetStr = DIALTARGET_QUEUE_TOKEN$
	End If
	
	GetDialTargetEnumAsString = dialTargetStr
	
End Function

'//--------------------------------------------------------------------
'// PROC: GetWrapupCodesAsXML
'//
'// DESCRIPTION:
'// Constructs an xml string representation of all the stored wrapup 
'// codes. The xml string assumes the following format,
'//
'//     <?xml version="1.0"?>
'//     <WRAPUP_CODES>
'//         <WRAPUP_CODE ID="...">
'//             <DISPLAY_NAME>...</DISPLAY_NAME>
'//             <LOCALE>...</LOCALE>
'//             <SELECTED>...</SELECTED>
'//         </WRAPUP_CODE>
'//     </WRAPUP_CODES>
'//
'// The xml string representation of 2 wrapup codes might look something 
'// like this,
'// 
'//     <?xml version="1.0"?>
'//     <WRAPUP_CODES>
'//         <WRAPUP_CODE ID="Code1">
'//             <DISPLAY_NAME>Code1</DISPLAY_NAME>
'//             <LOCALE>en_US</LOCALE>
'//             <SELECTED>FALSE</SELECTED>
'//         </WRAPUP_CODE>
'//         <WRAPUP_CODE ID="Code2">
'//             <DISPLAY_NAME>Code2</DISPLAY_NAME>
'//             <LOCALE>en_US</LOCALE>
'//             <SELECTED>TRUE</SELECTED>
'//         </WRAPUP_CODE>
'//     </WRAPUP_CODES>
'//
'//--------------------------------------------------------------------
Function GetWrapupCodesAsXML() As String

	Const QUOTE = """"
	Dim wrapupCodesXML As String
	wrapupCodesXML = "<?xml version=" & QUOTE & "1.0" & QUOTE & " ?>" & "<" & WRAPUP_CODES_TOKEN$ & ">"
	
	If Not(gAllWrapupCodes Is Nothing) Then
		
		'// Iterate through the container holding wrapup codes.
		Dim index As Integer
		Dim count As Integer
		count = gAllWrapupCodes.Count
		For index = 0 To count - 1

			'// Retrieve information of the next wrapup code.
			Dim wrapupCode As Variant
			Set wrapupCode = gAllWrapupCodes.Items()(index)

			'// Construct an xml string that holds the information of the wrapup code.
			Dim selected As String
			selected = IIf(wrapupCode.Item(SELECTED_TOKEN$) = True, "TRUE", "FALSE")
			Dim singleWrapupCodeXML As String
			singleWrapupCodeXML = "<"  & WRAPUP_CODE_TOKEN$  & " " & ID_TOKEN$ & "=" & QUOTE & wrapupCode.Item(ID_TOKEN$) & QUOTE & ">" & _
								  "<"  & DISPLAY_NAME_TOKEN$ & ">" & wrapupCode.Item(DISPLAY_NAME_TOKEN$) & "</" & DISPLAY_NAME_TOKEN$ & ">" & _
								  "<"  & LOCALE_TOKEN$       & ">" & wrapupCode.Item(LOCALE_TOKEN$)       & "</" & LOCALE_TOKEN$       & ">" & _
								  "<"  & SELECTED_TOKEN$     & ">" & selected                             & "</" & SELECTED_TOKEN$     & ">" & _
								  "</" & WRAPUP_CODE_TOKEN$  & ">"

			'// Then, append the string to our all inclusive xml string.
			wrapupCodesXML = wrapupCodesXML & singleWrapupCodeXML

		Next

	End If
	
	wrapupCodesXML = wrapupCodesXML & "</" & WRAPUP_CODES_TOKEN$ & ">"
	GetWrapupCodesAsXML = wrapupCodesXML

End Function

'//--------------------------------------------------------------------
'// PROC: HandleGetWrapupCodes
'//
'// DESCRIPTION:
'// Handles the logic that takes care of retrieving wrapup codes.
'//--------------------------------------------------------------------
Sub HandleGetWrapupCodes(id As String, src As APISource, dataset As APIDataSet)

    '// The result code will hold SUCCEEDED if wrapup codes are found
    '// for the specified user ID. Otherwise, the code will hold FAILED.
    Dim res As Object
    Set res = CreateObject("vbaapi.APIResult")
    res.Code = "FAILED"
    
    '// Will contain the wrapup codes information that we'll send back in a RESPONDER_RESPONSE.
    Dim resultDataSet As Object
    Set resultDataSet = CreateObject("vbaapi.APIDataSet")

	Dim uid As String
	uid = GetField(UID_TOKEN$, dataset)
	
	'// Call SetupWrapupUI so that the user of this script gets
	'// a chance to setup wrapup codes.
	Dim agentData As AgentDataType
	Dim ati As InteractionType
	InitWrapupUIData(agentData, ati, dataset)
	SetupWrapupUI(agentData, ati)

	'// Initialize the result dataset.
	resultDataSet.Initialize(2)
	resultDataSet.AddRow()
	resultDataSet.AddRow()
	resultDataSet.AddRow()
	resultDataSet.AddRow()
	resultDataSet.AddRow()

	'// Store the keys into the dataset. These will be inserted 
	'// into the first column of each row.
	resultDataSet.SetField(0, 0, CONTEXT_TOKEN$)
	resultDataSet.SetField(1, 0, UID_TOKEN$)
	resultDataSet.SetField(2, 0, IID_TOKEN$)
	resultDataSet.SetField(3, 0, WRAPUP_CODES_TOKEN$)
	resultDataSet.SetField(4, 0, SINGLE_SELECT_TOKEN$)

	'// Store the actual wrapup codes data. We'll store this into the 
	'// second column of each row. Data includes the UID, an associated
	'// IID, an xml string representation of the wrapup codes, and 
	'// the Single Select flag value.
	resultDataSet.SetField(0, 1, GET_WRAPUP_CODES_TOKEN$)
	resultDataSet.SetField(1, 1, uid)
	resultDataSet.SetField(2, 1, ati.IID)
	resultDataSet.SetField(3, 1, GetWrapupCodesAsXML())
	resultDataSet.SetField(4, 1, IIf(GetSingleSelectValue() = True, "TRUE", "FALSE"))
	
	'// Only add a dataset value for the Require Selection flag if SetRequireSelectionValue()
	'// was previously called.
	If Not(gAllWrapupCodesSettings Is Nothing) Then
		If gAllWrapupCodesSettings.Exists(REQUIRE_SELECTION_TOKEN$) Then
			resultDataSet.AddRow()
			resultDataSet.SetField(5, 0, REQUIRE_SELECTION_TOKEN$)
			resultDataSet.SetField(5, 1, IIf(GetRequireSelectionValue() = True, "TRUE", "FALSE"))	
		End If
	End If	

	'// Wrapup codes entries were found...
	res.Code = "SUCCEEDED"
    
    '// Construct the necessary information to send out a RESPONDER_RESPONSE.
    Dim dest As Object
    Set dest = CreateObject("vbaapi.APIDestination")
    dest.SetDestinationFromSource (src)
    
    '// Finally, send out a RESPONDER_RESPONSE back to the sender of the RESPONDER_REQUEST.
    SendResponseAsDataSet(id, dest, res, resultDataSet)

End Sub

'//--------------------------------------------------------------------
'// PROC: HandleModifyPhoneNumber
'//
'// DESCRIPTION:
'// Handles the logic that takes care of performing phone number
'// modification.
'//--------------------------------------------------------------------
Sub HandleModifyPhoneNumber(id As String, src As APISource, dataset As APIDataSet)

	'// Will contain the result of a phone number modification that we'll send back in a RESPONDER_RESPONSE.
	Dim resultDataSet As Object
	Set resultDataSet = CreateObject("vbaapi.APIDataSet")
	
	'// Retrieve the phone number to modify.
	Dim phoneNumber As String
	phoneNumber = GetField(PHONE_NUMBER_TOKEN$, dataset)
	
	'// Call ModifyPhoneNumber so that phone number modification occurs.
	Dim modifiedPhoneNumber As String
	Dim phoneData As PhoneDataType
	InitPhoneData(phoneData, dataset)
	modifiedPhoneNumber = ModifyPhoneNumber(phoneNumber, phoneData)
	
	'// Initialize the result dataset.
	resultDataSet.Initialize(2)
	resultDataSet.AddRow()
	resultDataSet.AddRow()
	
	'// Store the keys into the dataset. These will be inserted
	'// into the first column of each row.
	resultDataSet.SetField(0, 0, CONTEXT_TOKEN$)
	resultDataSet.SetField(1, 0, PHONE_NUMBER_TOKEN$)
	
	'// Store the modified phone number into the result dataset.
	resultDataSet.SetField(0, 1, MODIFY_PHONE_NUMBER_TOKEN$)
	resultDataSet.SetField(1, 1, modifiedPhoneNumber)
	
	'// Construct the necessary information to send out a RESPONDER_RESPONSE.
	Dim res As Object
	Dim dest As Object
	Set res = CreateObject("vbaapi.APIResult")
	Set dest = CreateObject("vbaapi.APIDestination")
	res.Code = "SUCCEEDED"
	dest.SetDestinationFromSource(src)
	
	'// Finally, send out a RESPONDER_RESPONSE back to the sender of the RESPONDER_REQUEST.
	SendResponseAsDataSet(id, dest, res, resultDataSet)

End Sub

'//--------------------------------------------------------------------
'// PROC: HandleValidatePhoneNumber
'//
'// DESCRIPTION:
'// Handles the logic that takes care of performing phone number 
'// validation.
'//--------------------------------------------------------------------
Sub HandleValidatePhoneNumber(id As String, src As APISource, dataset As APIDataSet)

	'// Will contain the result of a phone number validation that we'll send back in a RESPONDER_RESPONSE.
	Dim resultDataSet As Object
	Set resultDataSet = CreateObject("vbaapi.APIDataSet")
	
	'// Retrieve the phone number to validate.
	Dim phoneNumber As String
	phoneNumber = GetField(PHONE_NUMBER_TOKEN$, dataset)
	
	'// Call ValidatePhoneNumber so that phone number validation occurs.
	Dim validPhoneNumber As Boolean
	Dim invalidReason As String
	Dim phoneData As PhoneDataType
	InitPhoneData(phoneData, dataset)
	validPhoneNumber = ValidatePhoneNumber(phoneNumber, invalidReason, phoneData)
	
	'// Initialize the result dataset.
	resultDataSet.Initialize(2)
	resultDataSet.AddRow()
	resultDataSet.AddRow()
	resultDataSet.AddRow()
	
	'// Store the keys into the dataset. These will be inserted
	'// into the first column of each row.
	resultDataSet.SetField(0, 0, CONTEXT_TOKEN$)
	resultDataSet.SetField(1, 0, PHONE_NUMBER_TOKEN$)
	resultDataSet.SetField(2, 0, VALID_PHONE_NUMBER_TOKEN$)
	
	'// Store result information about the phone number validation. We'll
	'// store this into the second column of each row. Information will include
	'// the phone number, whether or not the phone number was interpreted 
	'// as being valid, and an invalid reason.
	resultDataSet.SetField(0, 1, VALIDATE_PHONE_NUMBER_TOKEN$)
	resultDataSet.SetField(1, 1, phoneNumber)
	resultDataSet.SetField(2, 1, IIf(validPhoneNumber = True, "TRUE", "FALSE"))
	
	'// Only add an invalid reason to the result dataset if the call to ValidatePhoneNumber
	'// returned false.
	If validPhoneNumber = False Then
		resultDataSet.AddRow()
		resultDataSet.SetField(3, 0, INVALID_REASON_TOKEN$)
		resultDataSet.SetField(3, 1, invalidReason)
	End If	

	'// Construct the necessary information to send out a RESPONDER_RESPONSE.
	Dim res As Object
	Dim dest As Object
	Set res = CreateObject("vbaapi.APIResult")
	Set dest = CreateObject("vbaapi.APIDestination")
	res.Code = "SUCCEEDED"
	dest.SetDestinationFromSource(src)
	
	'// Finally, send out a RESPONDER_RESPONSE back to the sender of the RESPONDER_REQUEST.
	SendResponseAsDataSet(id, dest, res, resultDataSet)

End Sub

'//--------------------------------------------------------------------
'// PROC: InitPhoneData
'//
'// DESCRIPTION:
'// Performs initialization of the passed in PhoneDataType object.
'//--------------------------------------------------------------------
Sub InitPhoneData(phoneData As PhoneDataType, dataset As APIDataSet)

	If dataset.GetColumns() = 2 Then
	
		'// Iterate through each row of the dataset, assuming that the first column of each
		'// row contains a token value while the second column of each rows contains
		'// actual data related to that token. Note that the dataset must have exactly 2 columns.
		Dim token As String
		Dim field As String
		Dim rowIndex As Integer
		Dim rows As Integer
		rows = dataset.GetRows()
		For rowIndex = 0 To rows - 1
		
			'// Determine the token at the current row and then store the actual data 
			'// in that row.
			token = dataset.GetField(rowIndex, 0)
			field = dataset.GetField(rowIndex, 1)
			
			If token = CONFIG_GROUP_NAME_TOKEN$ Then
				phoneData.ConfigGroupName = field
			ElseIf token = LINK_NAME_TOKEN$ Then
				phoneData.LinkName = field
			ElseIf token = USER_EXTENSION_TOKEN$ Then
				phoneData.UserExtension = field
			ElseIf token = CALL_TYPE_TOKEN$ Then
				phoneData.CallType = GetCallTypeEnum(field)
			ElseIf token = DIAL_TARGET_TOKEN$ Then
				phoneData.DialTarget = GetDialTargetEnum(field)
			End If
			
		Next
		
	End If
	
	Trace("Call Type = " & GetCallTypeEnumAsString(phoneData.CallType))
	Trace("Configuration Group Name = " & phoneData.ConfigGroupName)
	Trace("Dial Target = " & GetDialTargetEnumAsString(phoneData.DialTarget))
	Trace("Link Name = " & phoneData.LinkName)
	Trace("User Extension = " & phoneData.UserExtension)
	
End Sub

'//--------------------------------------------------------------------
'// PROC: InitWrapupUIData
'//
'// DESCRIPTION:
'// Performs initialization of the passed in AgentDataType and 
'// InteractionType objects.
'//--------------------------------------------------------------------
Sub InitWrapupUIData(agentData As AgentDataType, ati As InteractionType, dataset As APIDataSet)

	If dataset.GetColumns() = 2 Then
	
		'// Iterate through each row of the dataset, assuming that the first column of each
		'// row contains a token value while the second column of each rows contains
		'// actual data related to that token. Note that the dataset must have exactly 2 columns.
		Dim token As String
		Dim field As String
		Dim rowIndex As Integer
		Dim rows As Integer
		rows = dataset.GetRows()
		For rowIndex = 0 To rows - 1
		
			'// Determine the token at the current row and then store the actual data 
			'// in that row. Note that custom and standard interaction properties will 
			'// be received as xml strings. We'll need to parse the xml strings and 
			'// store the results into 2 Dictionary containers -- one for standard 
			'// interaction properties and the other for custom interaction properties. 
			'// The Dictionary containers will map interaction property keys to values.
			token = dataset.GetField(rowIndex, 0)
			field = dataset.GetField(rowIndex, 1)
			
			If token = STANDARD_INTERACTION_PROPERTIES_TOKEN$ Then
				Set ati.StandardInteractionProperties = CreateObject("Scripting.Dictionary")
				BuildInteractionProperties(field, ati.StandardInteractionProperties)
			ElseIf token = CUSTOM_INTERACTION_PROPERTIES_TOKEN$ Then
				Set ati.CustomInteractionProperties = CreateObject("Scripting.Dictionary")
				BuildInteractionProperties(field, ati.CustomInteractionProperties)
			ElseIf token = DISPLAY_INFO_TOKEN$ Then
				ati.DisplayInfo = field
			ElseIf token = DISPOSITION_TOKEN$ Then
				ati.Disposition = field
			ElseIf token = IID_TOKEN$ Then
				ati.IID = field
			ElseIf token = IS_CALL_OWNER_TOKEN$ Then
				ati.IsCallOwner = (UCase(field) = "TRUE")
			ElseIf token = ITYPE_TOKEN$ Then
				ati.IType = field
			ElseIf token = NAME_TOKEN$ Then
				ati.Name = field
			ElseIf token = ORIGINATING_IID_TOKEN$ Then
				ati.OriginatingIID = field
			ElseIf token = PARENT_IID_TOKEN$ Then
				ati.ParentIID = field
			ElseIf token = POP_DATA_TOKEN$ Then
				ati.PopData = field
			ElseIf token = QUEUE_ID_TOKEN$ Then
				ati.QueueID = field
			ElseIf token = QUEUE_NAME_TOKEN$ Then
				ati.Queue = field
			ElseIf token = SINGLE_AGENT_UID_TOKEN$ Then
				ati.SingleAgent = field
			ElseIf token = SOURCE_PROCESS_TOKEN$ Then
				ati.SourceProcess = field
			ElseIf token = STATE_TOKEN$ Then
				ati.State = field
			ElseIf token = CONFIG_GROUP_ID_TOKEN$ Then
				agentData.ConfigGroupID = field
			ElseIf token = CONFIG_GROUP_NAME_TOKEN$ Then
				agentData.ConfigGroupName = field
			ElseIf token = EMAIL_ADDRESS_TOKEN$ Then
				agentData.EmailAddress = field
			ElseIf token = EXTENSION_TOKEN$ Then
				agentData.Extension = field
			ElseIf token = LINK_NAME_TOKEN$ Then
				agentData.LinkName = field
			ElseIf token = UID_TOKEN$ Then
				agentData.UserID = field
			ElseIf token = USER_NAME_TOKEN$ Then
				agentData.UserName = field
			End If
			
		Next
		
	End If

	Trace("ConfigGroupID = " & agentData.ConfigGroupID)
	Trace("ConfigGroupName = " & agentData.ConfigGroupName)
	Trace("DisplayInfo = " & ati.DisplayInfo)
	Trace("Disposition = " & ati.Disposition)
	Trace("EmailAddress = " & agentData.EmailAddress)
	Trace("Extension = " & agentData.Extension)
	Trace("IID = " & ati.IID)
	Trace("IsCallOwner = " & CStr(ati.IsCallOwner))
	Trace("IType = " & ati.IType)
	Trace("LinkName = " & agentData.LinkName)
	Trace("Name = " & ati.Name)
	Trace("OriginatingIID = " & ati.OriginatingIID)
	Trace("ParentIID = " & ati.ParentIID)
	Trace("PopData = " & ati.PopData)
	Trace("Queue = " & ati.Queue)
	Trace("QueueID = " & ati.QueueID)
	Trace("SingleAgent = " & ati.SingleAgent)
	Trace("State = " & ati.State)
	Trace("SourceProcess = " & ati.SourceProcess)
	Trace("UID = " & agentData.UserID)
	Trace("UserName = " & agentData.UserName)
	Trace("StandardInteractionProperties = {" & InteractionPropertiesToString(ati.StandardInteractionProperties) & "}")
	Trace("CustomInteractionProperties = {" & InteractionPropertiesToString(ati.CustomInteractionProperties) & "}")
	
End Sub

'//--------------------------------------------------------------------
'// PROC: InteractionPropertiesToString
'//
'// DESCRIPTION:
'// Returns a nicely formatted string containing the contents of the
'// passed in interaction properties Dictionary container.
'//--------------------------------------------------------------------
Function InteractionPropertiesToString(interactionProperties As Object) As String

	Dim result As String
	If Not(interactionProperties Is Nothing) Then
	
		Dim index As Integer
		Dim maxIndex As Integer
		Dim keyArray As Variant
		keyArray = interactionProperties.Keys()
		maxIndex = UBound(interactionProperties.Keys())
		For index = 0 To maxIndex

			Dim key As String
			Dim value As String
			key = keyArray(index)
			value = interactionProperties.Item(key)
			result = result & "[" & key & "] = " & value & IIf(index < maxIndex, ", ", "")
			
		Next index
		
	End If
	InteractionPropertiesToString = result
	
End Function

'//--------------------------------------------------------------------
'// PROC: ModifyPhoneNumber
'//
'// DESCRIPTION:
'// This function is called before a phone number is dialed from an agent.
'// The number that is to be dialed is passed into the function as 
'// phoneNumber. The agent dials the string returned from this function. 
'// This function can be used to modify a phone number, such as adding 
'// special dialing codes, before dialing the phone number. The passed in
'// PhoneDataType object contains phone-related information that can be
'// used to place conditions around phone modifications.
'//--------------------------------------------------------------------
Function ModifyPhoneNumber(phoneNumber As String, phoneData As PhoneDataType) As String
	ModifyPhoneNumber = phoneNumber
End Function

'//--------------------------------------------------------------------
'// PROC: RetrieveKeyValuePair
'//
'// DESCRIPTION:
'// Recursively traverses an xml tree until a KEY and VALUE are found.
'//--------------------------------------------------------------------
Sub RetrieveKeyValuePair(childNodes As Object, ByRef key As String, ByRef value As String)

	If Not(childNodes Is Nothing) Then

		Dim childNode As Object
		For Each childNode In childNodes

			If UCase(childNode.parentNode.nodeName) = KEY_TOKEN$ Then
				key = childNode.nodeValue
			ElseIf UCase(childNode.parentNode.nodeName) = VALUE_TOKEN$ Then
				value = childNode.nodeValue
			Else
				RetrieveKeyValuePair(childNode.childNodes, key, value)
			End If

		Next childNode
		
	End If
		
End Sub

'//--------------------------------------------------------------------
'// PROC: SetRequireSelectionValue
'//
'// DESCRIPTION:
'// If false is passed in, then users are allowed to cancel out a wrapup 
'// code selection without making a choice. A value of true indicates 
'// that a wrapup codes selection is required.
'//--------------------------------------------------------------------
Sub SetRequireSelectionValue(requireSelection As Boolean)

	If gAllWrapupCodesSettings Is Nothing Then
		Set gAllWrapupCodesSettings = CreateObject("Scripting.Dictionary")
	End If
	
	'// Note that if the require selection setting has already been set, we'll
	'// just overwrite it with the passed in value.
	If gAllWrapupCodesSettings.Exists(REQUIRE_SELECTION_TOKEN$) = True Then
		Set gAllWrapupCodesSettings.Item(REQUIRE_SELECTION_TOKEN$) = requireSelection
	Else
		gAllWrapupCodesSettings.Add(REQUIRE_SELECTION_TOKEN$, requireSelection)
	End If
	
End Sub

'//--------------------------------------------------------------------
'// PROC: SetSingleSelectValue
'//
'// DESCRIPTION:
'// If true is passed in, wrapup codes are deemed to support selection of 
'// a single wrapup code. On the otherhand, if false is passed in, then 
'// wrapup codes are deemed to support selection of multiple wrapup codes.
'//--------------------------------------------------------------------
Sub SetSingleSelectValue(singleSelect As Boolean)
	
	If gAllWrapupCodesSettings Is Nothing Then
		Set gAllWrapupCodesSettings = CreateObject("Scripting.Dictionary")
	End If
	
	'// Note that if the single select setting has already been set, we'll 
	'// just overwrite it with the passed in value.
	If gAllWrapupCodesSettings.Exists(SINGLE_SELECT_TOKEN$) = True Then
		Set gAllWrapupCodesSettings.Item(SINGLE_SELECT_TOKEN$) = singleSelect
	Else
		gAllWrapupCodesSettings.Add(SINGLE_SELECT_TOKEN$, singleSelect)	
	End If
	
End Sub

'//--------------------------------------------------------------------
'// PROC: SetupWrapupUI
'//
'// DESCRIPTION:
'// This procedure is used to set up wrapup codes. Wrapup codes are
'// inserted by making calls to AddWrapupCode. The passed in AgentDataType
'// and InteractionType objects can be used to set conditionals around 
'// wrapup codes.
'//--------------------------------------------------------------------
Sub SetupWrapupUI(agentData As AgentDataType, ati As InteractionType)
	AddWrapupCode("DONE","Done","en_US",true)
	AddWrapupCode("Follow_Up","Follow Up","en_US",true)
End Sub

'//--------------------------------------------------------------------
'// PROC: ValidatePhoneNumber
'//
'// DESCRIPTION:
'// This function is called when an agent places an outbound call. It
'// takes care of determining whether or not the passed in phoneNumber
'// is valid. This function should return True if the phone number is
'// valid and False, if it is invalid. If the phone number is invalid,
'// invalidReason should be set to a string that the user would like to
'// display. The passed in PhoneDataType object contains phone-related 
'// information that can be used to place conditions around phone
'// validations.
'//--------------------------------------------------------------------
Function ValidatePhoneNumber(phoneNumber As String, ByRef invalidReason As String, phoneData As PhoneDataType) As Boolean

    Dim returnAnswer As Boolean
    returnAnswer = True
    
    '// Validation logic goes here

    ValidatePhoneNumber = returnAnswer
    
End Function
