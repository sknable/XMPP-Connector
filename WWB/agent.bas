Attribute VB_Name = "Module1"
'--------------------------------------------------------------------
'
' DESCRIPTION: This is an example SaxBasic script file for the Apropos
' Agent Application. A lot of customization work is done through this
' script. Please refer to the Apropos Programmer's Guide for full
' documentation of the SaxBasic Macro Language, as well as the
' SaxBasic Agent Extensions.
'
' Note that there are several "helper" procedures defined in this
' script that you may find generally useful:
'
'       Sub popAndStuff (partialTitle As String, keyStrokes As String)
'       Sub stuffKeys (keyStrokes As String)
'       Function findWindowTitle (partialTitle As String) As String
'       Sub popApp(partialTitle As String)
'       Sub sendDDE(t As String)
'       Function appTitle As String
'
'--------------------------------------------------------------------

Option Explicit

' Global Variables

'Object used for IE browser that displays iVault
Dim IEiVault As Object 

Dim responderRequestId As Integer

' Define what demos we can run (also see appTitle)
Const aproposDemo = 1
Const remedy = 2
Const scopus = 3
Const expertAdvisor = 4

' Set this constant to determine which demo to run
Const runDemo = aproposDemo

' Define what PBXs we support for Enhanced and Switch Link agents
Const UNKNOWN_SWITCH = 1
Const MITEL_SX2000L = 2
Const NORTEL_MERIDIAN1 = 3
Const SIEMENS_HICOM_300E_9006 = 4

' Set this constant to determine which PBX is being used
Const pbxType = UNKNOWN_SWITCH

'Windows API function declarations:
Declare Function GetTopWindow& Lib "user32" (ByVal hWnd&)
Declare Function GetActiveWindow& Lib "user32" ()
Declare Function GetWindow& Lib "user32" (ByVal hWnd&, ByVal wCmd&)
Declare Function GetWindowText& Lib "user32" (ByVal hWnd&, ByVal lpString$, ByVal nMaxCount&)
Declare Function GetWindowTextLength& Lib "user32" (ByVal hWnd&)
Declare Function MessageBox& Lib "user32" (ByVal hWnd&, ByVal lpText$, ByVal lpCaption$, ByVal style&)
Declare Function ShellExecute& Lib "shell32"  (ByVal hWnd&, ByVal lpOperation$, ByVal lpFile$, ByVal lpParameters$, ByVal lpDirectory$, ByVal nShowCmd&)
Declare Function IsIconic& Lib "user32" (ByVal hWnd&)
Declare Function ShowWindow& Lib "user32" (ByVal hWnd&, ByVal nCmdShow&)


' Load the constants

'#Uses "const.bas"

'---------------------------------------------------------------------
' FUNCTION: appTitle
'
' DESCRIPTION:
'   Return the desktop application title based on value in runDemo.
'---------------------------------------------------------------------
Function appTitle() As String
    Select Case runDemo
        Case aproposDemo
           appTitle = "Apropos"
        Case remedy
           appTitle = "Action Request System"
        Case scopus
           appTitle = "Scopus Foundation 32"
        Case expertAdvisor
           appTitle = "Expert Advisor"
        Case Else
           appTitle = "Notepad"
    End Select
End Function

'--------------------------------------------------------------------
' SUB: Initialize
'
' DESCRIPTION:
'   This procedure will be executed when Agent starts and every time you
'   stop running Sax script and resume running again from the Sax Basic
'   Debug IDE. You should set dial strings, phone user interface buttons,
'   macro tasks, unavailability reason codes here. You can use Sax Basic
'   IDE to debug those Settings.
'-------------------------------------------------------------------
Sub Initialize()

    If AgentMode = AGENTMODE_VOICE_CARD Then

        ' Set up the proper dial strings for the PBX
        Call SetupDialStrings(pbxType)

    End If

       
    ' Set phone user interface for Agent with native link or voice card
    Call SetupPhoneUI

    ' Set up pre-programmed macro tasks
    Call SetupMacroUI
        
    ' Set unavailability reason codes
    Call SetupUnavailabilityReasonUI
         
    ' Set the URLs for the browser part of the web chat
    Call SetupFavoriteURLs
End Sub

'--------------------------------------------------------------------
' SUB: SetupDialStrings
'
' DESCRIPTION:
'   This procedure will set the proper dial strings for the specified
'   switch type
'
'--------------------------------------------------------------------
Sub SetupDialStrings(switchType As Integer)

    If (switchType = MITEL_SX2000L) Then

        ' Settings for a Mitel SX2000 Light PBX
        DialStringForHold = "!,*7"
        DialStringForHoldRetrieve = "!*14"
        DialStringForTransfer = "!,"
        DialStringForConference = "!,"
        DialStringForCancelTransfer = "!,*22!,*14"
        DialStringForCompleteConference = "!,*3"
        DialStringForCancelConference = "!,*22!,*14"

    ElseIf (switchType = SIEMENS_HICOM_300E_9006) Then

        ' Settings for a Siemens Hicom 300E 9006 PBX
        DialStringForHold = "!,*9"
        DialStringForHoldRetrieve = "!*9"
        DialStringForTransfer = "!,"
        DialStringForConference = "!,"
        DialStringForCancelTransfer = "!,!"
        DialStringForCompleteConference = "!"
        DialStringForCancelConference = "!,!"

    ElseIf (switchType = NORTEL_MERIDIAN1) Then

        ' Settings for a Nortel Meridian 1 Option 11 PBX
        DialStringForHold = "!,#4"
        DialStringForHoldRetrieve = "!#4"
        DialStringForTransfer = "!,"
        DialStringForConference = "!,"
        DialStringForCancelTransfer = "!,!"
        DialStringForCompleteConference = "!"
        DialStringForCancelConference = "!,!"

    Else

        ' Unknown PBX type-- use general settings
        DialStringForHold = "!"
        DialStringForHoldRetrieve = "!"
        DialStringForTransfer = "!,"
        DialStringForConference = "!,"
        DialStringForCancelTransfer = "!,!"
        DialStringForCompleteConference = "!"
        DialStringForCancelConference = "!,!"

    End If

End Sub


'---------------------------------------------------------------------
' SUB: AgentStarted
'
' DESCRIPTION:
'   This procedure is only executed once after Agent starts. You can put
'   ACD login, etc. initialization work here
'---------------------------------------------------------------------
Sub AgentStarted()

    responderRequestId = 1

    ' this needs to be done in AgentStarted because IsSupported depends on
    ' phone being initialized.
    ' Note: PHONEFEATURE_SWAP is defined in const.bas as of version 7.0.0
    ' If upgrading, this file must be copied from configfiles/default_files to configfiles/global_scripts
    ' on the primary server to prevent an Agent startup failure.

    If AgentMode = AGENTMODE_SWITCH_LINK And IsSupported(PHONEFEATURE_SWAP) Then
    ' Swap button-- only valid for switchlink and switches that support it

        AddPhoneCommand("Swap", PHONEBUTTON_SWAP, "", "", ENABLINGRULE_ENABLED_BY_STATE)
    End If
End Sub

'---------------------------------------------------------------------
' SUB: PrepareExit
'
' DESCRIPTION:
'   This procedure is called just before exiting the agent.  The phone
'   connection will still be active, so you can send messages to the
'   switch.
'---------------------------------------------------------------------
Sub PrepareExit()

End Sub

'---------------------------------------------------------------------
' SUB: AgentExit
'
' DESCRIPTION:
'   This procedure will be executed when Agent exits. You can put
'   your cleanup code here, e.g. delete the objects you created in the
'   script.
'
'   Note that the phone connection has been terminated before this
'   function gets called!  If you need to perform phone-related shutdown
'   tasks, they must be done in PrepareExit()
'---------------------------------------------------------------------
Sub AgentExit()
    Set IEiVault = Nothing
End Sub

'--------------------------------------------------------------------
' SUB: SetupPhoneUI
'
' DESCRIPTION:
'   This procedure is to used to set up the buttons in the phone
'   interface.
'
'   Note: If you need a special dial string for a standard button (for
'   instance if you must send a different string for cancel transfer
'   if you get a busy signal), you can do it using something like the
'   following:
'
'   ' Send a normal Cancel Transfer
'   AddPhoneCommand("Cancel Transfer", PHONEBUTTON_CANCEL_TRANSFER, "", "", ENABLINGRULE_ENABLED_BY_STATE)
'
'   ' Cancel Transfer after a busy signal (dial code "*123")
'   AddPhoneCommand("Cancel Transfer (Busy)", PHONEBUTTON_CANCEL_TRANSFER, "*123", "", ENABLINGRULE_ENABLED_BY_STATE)
'--------------------------------------------------------------------
Sub SetupPhoneUI()

    ' Add the Standard Phone buttons
    AddPhoneCommand("Answer", PHONEBUTTON_ANSWER, "", "", ENABLINGRULE_ENABLED_BY_STATE)
    AddPhoneCommand("Call", PHONEBUTTON_CALL, "?", "", ENABLINGRULE_ENABLED_BY_STATE)
    AddPhoneCommand("Cancel", PHONEBUTTON_CANCEL_CALL, "", "", ENABLINGRULE_ENABLED_BY_STATE)
    AddPhoneCommand("Hold", PHONEBUTTON_HOLD, "", "", ENABLINGRULE_ENABLED_BY_STATE)
    AddPhoneCommand("Hold Retrieve", PHONEBUTTON_HOLD_RETRIEVE, "", "", ENABLINGRULE_ENABLED_BY_STATE)
    AddPhoneCommand("Start Transfer", PHONEBUTTON_START_TRANSFER, "Q", "", ENABLINGRULE_ENABLED_BY_STATE)
    AddPhoneCommand("Complete Transfer", PHONEBUTTON_COMPLETE_TRANSFER, "", "", ENABLINGRULE_ENABLED_BY_STATE)
    If (pbxType = MITEL_SX2000L And AgentMode = AGENTMODE_VOICE_CARD) Then
        AddPhoneCommand("Cancel Transfer After Answer", PHONEBUTTON_CANCEL_TRANSFER, "", "", ENABLINGRULE_ENABLED_BY_STATE)
        AddPhoneCommand("Cancel Transfer No Answer", PHONEBUTTON_CANCEL_TRANSFER, "!,*14", "", ENABLINGRULE_ENABLED_BY_STATE)
    Else
        AddPhoneCommand("Cancel Transfer", PHONEBUTTON_CANCEL_TRANSFER, "", "", ENABLINGRULE_ENABLED_BY_STATE)
    End If
    AddPhoneCommand("Start Conference", PHONEBUTTON_START_CONFERENCE, "Q", "", ENABLINGRULE_ENABLED_BY_STATE)
    AddPhoneCommand("Complete Conference", PHONEBUTTON_COMPLETE_CONFERENCE, "", "", ENABLINGRULE_ENABLED_BY_STATE)
    AddPhoneCommand("Cancel Conference", PHONEBUTTON_CANCEL_CONFERENCE, "", "", ENABLINGRULE_ENABLED_BY_STATE)

    If ((pbxType = NORTEL_MERIDIAN1 And AgentMode = AGENTMODE_SWITCH_LINK) Or (pbxType = SIEMENS_HICOM_300E_9006 And AgentMode = AGENTMODE_SWITCH_LINK)) Then
                ' For these switches, the dial command doesn't work for a digital phone
    Else
        AddPhoneCommand("Key Pad", PHONEBUTTON_DIAL, "K", "", ENABLINGRULE_ALWAYS_ENABLED)
    End If

    AddPhoneCommand("Hang Up", PHONEBUTTON_DISCONNECT, "", "", ENABLINGRULE_ALWAYS_ENABLED)

    AddPhoneCommand("Redial", PHONEBUTTON_REDIAL, "", "", ENABLINGRULE_ENABLED_BY_STATE)

    If AgentMode = AGENTMODE_VOICE_CARD Then
        ' Mute Button-- only valid for voicecard enhanced agents!
        AddPhoneCommand("Mute", PHONEBUTTON_CUSTOM, "", "Mute", ENABLINGRULE_ALWAYS_ENABLED)
   End If

End Sub

'--------------------------------------------------------------------
' SUB: Mute
'
' DESCRIPTION:
'   This procedure implements a simple phone mute interface.  It
'   will only work for voicecard enhanced agents!
'--------------------------------------------------------------------
Sub Mute(dialstring As String, ati As ATInteraction)

    ' Mute the headset mike
    SetMute(VOICECARD_CHANNEL_AGENT_MIC,1)

        ' Display a message box letting the agent know we are muted
        ' When he presses OK, we will continue.  Note that we use
        ' the Win32 MessageBox instead of MsgBox, so that we can
        ' specify the "always on top" option (&h40000)
        MessageBox( 0, "Press OK to restore normal speaking volume", "Phone Muted", &H40040 )

    ' Reset the original volume
    SetMute(VOICECARD_CHANNEL_AGENT_MIC,0)

End Sub

'--------------------------------------------------------------------
' SUB: SetupUnavailabilityReasonUI
'
' DESCRIPTION:
'   This procedure is to set up unavailability reason menu list
'   in the agent App. Note: If the Unavailable Reasons setting in
'   Configuration Manager is not blank, these values will be ignored.
'--------------------------------------------------------------------
Sub SetupUnavailabilityReasonUI()

End Sub

'--------------------------------------------------------------------
' SUB: SetupMacroUI
'
' DESCRIPTION:
'   This procedure is to set up Macro task menu list in the agent App.
'--------------------------------------------------------------------
Sub SetupMacroUI()

    AddTaskMenuItem("Send Jabber", "Send Jabber")

End Sub

'--------------------------------------------------------------------
' SUB: TaskMenuItemSelected
'
' DESCRIPTION:
'   This procedure handles task menu items not associated with the 
'   current interaction.
'--------------------------------------------------------------------
Sub TaskMenuItemSelected(name As String)

End Sub

'--------------------------------------------------------------------
' SUB: TaskMenuItemSelectedForInteraction
'
' DESCRIPTION:
'   This procedure handles task menu items associated with the 
'   current interaction.
'--------------------------------------------------------------------
Sub TaskMenuItemSelectedForInteraction(name As String, ati As ATInteraction)

End Sub

'--------------------------------------------------------------------
' SUB: ContactInfo
'
' DESCRIPTION:
'   This item on the task menu displays information about the
'   interaction and lets the user modify it.
'--------------------------------------------------------------------

Sub ContactInfo(ati As ATInteraction)

    Dim newATI As Object

    'create a new ATInteraction object and make it an
    'exact copy of ati
    Set newATI = CreateObject("AgentX.ATInteraction6")
    newATI.Copy (ati)

    Dim atiname As String
    Dim atiqueue As String
    Dim atiphone As String
    Dim atiOrg As String
    Dim atiContactID As String
    Dim atiIssueID As String
    Dim atiDisplayInfo As String
    Dim atiEmail As String
        
    atiname = ati.Name
    atiqueue = ati.Queue
    atiDisplayInfo = ati.DisplayInfo
    atiphone = ati.GetKeyValue("PHONE")
    atiOrg = ati.GetKeyValue("ORGANIZATION")
    atiIssueID = ati.GetKeyValue("ISSUEID")
    atiEmail = ati.GetKeyValue("EMAILADDRESS")


    ' Create a simple dialog to display the information
    Begin Dialog UserDialog 300,180, "Update Customer Information"
        
        Text 10, 10, 100, 15, "&Name"
        TextBox 120, 10, 150, 15, .myname$
        
        ' There is no read-only option, and TextBox is the only
        ' text field which will horizontally scroll. This is important
        ' because the queue name can be pretty long.
        Text 10, 30, 100, 15, "&Queue"
        TextBox 120, 30, 150, 15, .myqueue$

        Text 10, 50, 100, 15, "&Display Info"
        TextBox 120, 50, 150, 15, .mydisplayinfo$

        Text 10, 70, 100, 15, "&Phone Number"
        TextBox 120, 70, 150, 15, .myPhonenumber$
        
        Text 10, 90, 100, 15, "&Organization"
        TextBox 120, 90, 150, 15, .myOrganization$
                
        Text 10, 110, 100, 15, "&Issue ID"
        TextBox 120, 110, 150, 15, .myIssueID$

    Text 10, 130, 100, 15, "&Email"
        TextBox 120, 130, 150, 15, .myEmail$
        
        OKButton 80, 152, 60, 20
        CancelButton 160, 152, 60, 20
   
    End Dialog

    Dim dlg As UserDialog

    ' we intentionally don't let the user change the queue.
    dlg.myname$ = atiname
    dlg.myqueue$ = atiqueue
    dlg.myPhonenumber$ = atiphone
    dlg.mydisplayinfo$ = atiDisplayInfo
    dlg.myOrganization$ = atiOrg
    dlg.myIssueID$ = atiIssueID
    dlg.myEmail$ = atiEmail

    ' show the dialog
    Dim nRes As Integer
    nRes = Dialog(dlg)

    If nRes = -1 Then
        ' OK button is clicked
        newATI.Name = dlg.myname
        newATI.DisplayInfo = dlg.mydisplayinfo
        newATI.SetKeyValue("PHONE", dlg.myPhonenumber)
        newATI.SetKeyValue("ORGANIZATION", dlg.myOrganization)
        newATI.SetKeyValue("ISSUEID", dlg.myIssueID)
        newATI.SetKeyValue("EMAILADDRESS", dlg.myEmail)

        UpdateCustomerData (newATI)
    End If
        
    'delete the new ati object
    Set newATI = Nothing

End Sub


'--------------------------------------------------------------------
' SUB: EditConfigFiles
'
' DESCRIPTION:
'    This procedure opens up the global .ini file and
'    the .settings file for the user.  After modifying
'    these files, the Agent application
'    needs to be restarted for any changes to take effect.
'--------------------------------------------------------------------

Sub EditConfigFiles()
    Dim X As Long
    Dim uid As Integer
    Dim filename As String
    Dim cmd As String
    filename = SystemDirectory + "settings.ini"
    cmd = "notepad " + filename
    X = Shell(cmd, vbNormalFocus)
    AppActivate X

    uid = AgentUID
    filename = SystemDirectory + "Users\user" + CStr(uid) + ".settings"
    cmd = "notepad " + filename
    X = Shell(cmd, vbNormalFocus)
    AppActivate X
End Sub

'--------------------------------------------------------------------
' SUB: WriteMemo
'
' DESCRIPTION:
'    This procedure is the implementation of the task "Write a Memo"
'--------------------------------------------------------------------
Sub WriteMemo(ati As ATInteraction)

    popAndStuff "WordPad", "%FN{enter}%N%FO" & "template.doc{enter}{enter}{enter}%ID{down}{down}{down}{down}{down}{down}{down}{enter}{enter}{enter}" & ati.Name & "{enter}"

End Sub

'--------------------------------------------------------------------
' SUB: SendFax
'
' DESCRIPTION:
'    This procedure is the implementation of the task "Send a Fax"
'--------------------------------------------------------------------
Sub SendFax(ati As ATInteraction)

    ' Define interface to your fax software

End Sub


'--------------------------------------------------------------------
' SUB: WriteLetter
'
' DESCRIPTION:
'   This procedure is the implementation of task "Write a Letter"
'--------------------------------------------------------------------
Sub WriteLetter(ati As ATInteraction)

    popAndStuff "WordPad", "%FN{enter}%N%FO" & "template.doc{enter}{enter}{enter}%ID{down}{down}{down}{down}{down}{down}{down}{enter}{enter}{enter}" & ati.Name & "{enter}"

End Sub


'--------------------------------------------------------------------
' SUB: SetupWrapupUI
'
' DESCRIPTION:
'   This procedure is to set up wrap-up interface in the agent App.
'   Note that the available wrap-up codes are different based on
'   the type of interaction and its final disposition
'--------------------------------------------------------------------
Sub SetupWrapupUI(ati As ATInteraction)

    Dim atidisposition As String
    Dim atiitype As Integer
    atidisposition = ati.Disposition
    atiitype = ati.IType

    AddWrapupCodeEx(ati.iid,"Issue resolved", "Resolved", False)

    If Len(atidisposition) = 0 Or atidisposition = "COMPLETED" Or atidisposition = "UNKNOWN" Then
        AddWrapupCodeEx(ati.iid,"Schedule follow up tomorrow", "Call Tomorrow", True)
    Else
        AddWrapupCodeEx(ati.iid,"Schedule follow up tomorrow", "Call Tomorrow", False)
    End If

    AddWrapupCodeEx(ati.iid, "Follow up question", "Follow up question", False)


    If atidisposition = "TRANSFERRED" Or _
       atidisposition = "TRANSFERRED_EXTERNAL" Then
        AddWrapupCodeEx(ati.iid,"No follow-up required", "No Followup", False)
        AddWrapupCodeEx(ati.iid,"Transferred", "Transferred", True)
    ElseIf atidisposition = "REQUEUED" Then
        AddWrapupCodeEx(ati.iid,"No follow-up required", "No Followup", False)
        AddWrapupCodeEx(ati.iid,"Re-queued", "Requeued", True)
    ElseIf atidisposition = "DELETED" Then ' deleted e-mail
        AddWrapupCodeEx(ati.iid,"No follow-up required", "No Followup", True)
    End If

    If atiitype = ITYPE_OUTBOUND_DIRECT_CALL Then
        AddWrapupCodeEx(ati.iid,"Busy", "Busy", False)
        AddWrapupCodeEx(ati.iid,"No answer", "No Answer", False)
    End If

End Sub

'--------------------------------------------------------------------
' SUB: Wrapup
'
' DESCRIPTION:
'   This procedure is called after Agent finishes an interaction and
'   get wrap-up codes. The wrap-up codes are passed from Agent and
'   in variable wrapupCode.
'--------------------------------------------------------------------
Sub Wrapup(wrapupCode As String, ati As ATInteraction)

    If InStr(wrapupCode, "Follow up question") <> 0 Then

        Dim objDest As APIDestination
        Dim objData As APIDataSet

        Dim id As String

        Set objDest = CreateObject("VBAAPI.APIDestination")
        Set objData = CreateObject("VBAAPI.APIDataSet")


        objDest.SetDestination(1, "XMPP")
        objDest.ResponderGroup = ""
        objData.Initialize(4)
        objData.AddRow()
        objData.SetField(0, 0, "ContactJabber")
        objData.SetField(0, 1, ati.GetKeyValue("JID"))
        objData.SetField(0, 2, AgentName)
        objData.SetField(0, 3, ati.Queue())


        responderRequestId = responderRequestId + 1
        id = CStr(responderRequestId)

        SendRequestAsDataSet(id, objDest, objData)

    End If

End Sub



'--------------------------------------------------------------------
' SUB: SetupFavoriteURLsI
'
' DESCRIPTION:
'   This procedure is to set up the URLs that first
'   appear in the drop down combo box for a web chat
'--------------------------------------------------------------------
Sub SetupFavoriteURLs()
    AddFavoriteURL("www.llbeam.com")
    AddFavoriteURL("www.amazon.com")
    AddFavoriteURL("www.fidelity.com")
End Sub

'---------------------------------------------------------------------
' FUNCTION: InitialURL
'
' DESCRIPTION: Returns the initial URL which should be
' displayed in the Agent's browser when taking a web chat
'---------------------------------------------------------------------

Function InitialURL(ati As ATInteraction) As String
    ' To change the initial page to a custom page in the Agent's
    ' web browser, set the InitialURL as in the example below
    'InitialURL = "http://www.syntellect.com"
End Function

'--------------------------------------------------------------------
' SUB: Preview
'
' DESCRIPTION:
'   This procedure is called when Agent previews an interaction. It
'   pops up the desk application.
'--------------------------------------------------------------------
Sub Preview(ati As ATInteraction)

    ' Do the standard Take processing-- note that you will get
    ' ANOTHER Take if the user takes the preview, so this
    ' would not be appropriate if you cannot double-pop your
    ' desk application

        Call PopScreen(ati)

End Sub

'--------------------------------------------------------------------
' SUB: TransferTaken
'
' DESCRIPTION:
'   This procedure is called when Agent takes a transferred interaction.
'   It pops up the desk application.
'--------------------------------------------------------------------
Sub TransferTaken(ati As ATInteraction)

    ' Do the standard Take processing
        Call PopScreen(ati)

End Sub

'--------------------------------------------------------------------
' SUB: Taken
'
' DESCRIPTION:
'   This procedure is called when Agent takes an interaction. It
'   pops up the desk application.
'--------------------------------------------------------------------
Sub Taken(ati As ATInteraction)

End Sub

'---------------------------------------------------------------------
' SUB: WebChatDelivered
'
' DESCRIPTION:
'   called when the INTERACTION_TAKEN message comes through
' for a web chat.
' The subroutine below demonstrates typical web chat handling
'---------------------------------------------------------------------
Sub WebChatDelivered(ati As ATInteraction)
    If ati.IType = ITYPE_WEB_CHAT Then

        Dim iid As String
        Dim YourURL As String

        iid = ati.IID

        'YourURL = "http:demo1/wisp/agents/" + AgentName
        YourURL = "http://www.google.com/"
        
        SendBrowserURL(iid, YourURL)

    End If
End Sub

'---------------------------------------------------------------------
' SUB: PopScreen
'
' DESCRIPTION:
'   Called when a screen pop is needed.
'---------------------------------------------------------------------
Sub PopScreen(ati As ATInteraction)

    Dim keyStrokes As String
    Dim firstName As String
    Dim lastName As String
    Dim middle As Integer

    ' Figure out the agent first and last name from ati.name
    middle = InStr(ati.Name, " ")
    If middle = 0 Then
        firstName = ""
        lastName = ati.Name
    Else
        firstName = Left$(ati.Name, middle - 1)
        lastName = Right$(ati.Name, Len(ati.Name) - middle)
    End If

    ' Build the proper screen pop string for the application
    If (runDemo = scopus) Then
        keyStrokes = "%fc~^nSuccessful Prods{Tab}{Tab}" + firstName + "{Tab}" + lastName + "{Tab}" + "555-555-1212%wh"
    ElseIf (runDemo = remedy) Then
        popAndStuff appTitle, "^{F4}~^s~"
        Wait 3
        keyStrokes = "{Tab}" + lastName + "{Tab}" + firstName + "{Tab}174532{Tab}800-555-1212"
    ElseIf (runDemo = expertAdvisor) Then
        popAndStuff "Call Registration", "{esc}"
        keyStrokes = "%fn{Tab}{Tab}" + lastName + "{Tab}"
    End If
        
    ' Pop the desktop app with the screen pop string
    popAndStuff appTitle, keyStrokes
        
End Sub

'---------------------------------------------------------------------
' SUB: EmailDelivered
'
' DESCRIPTION:
'   Called when an email is delivered to the agent who is taking an
'   email request.  We no longer pop up outlook because the Agent
'   uses an integrated mail client.
'---------------------------------------------------------------------
Sub EmailDelivered(ati As ATInteraction)
End Sub

'---------------------------------------------------------------------
' SUB: sendDDE
'
' DESCRIPTION:
'   A helper method to call DDE. Here the example is ExpertAdvisor server.
'---------------------------------------------------------------------
Sub sendDDE(t As String)

    Dim c
    Dim first As Integer
    Dim lastTwo As String
    Dim errflag As String

    lastTwo = Right(t, 2)

    first = InStr(lastTwo, "^")

    On Error GoTo leaveDDE
    If first = 1 Then
        c = ddeinitiate("eadvisor", "incomingcall")
                        errflag = "Init:"
                        ddepoke c, "incomingcall", t
                        errflag = "poke:"
    Else
        c = ddeinitiate("eadvisor", "incomingproblem")
                        errflag = "Init:"
                        ddepoke c, "incomingproblem", t
                        errflag = "poke:"
    End If

    ddeterminate c
    errflag = "Term:"
    GoTo exitDDE

leaveDDE:
    MsgBox errflag & ":" & t
    Exit Sub

exitDDE:
End Sub

 
'---------------------------------------------------------------------
' SUB: findWindowTitle
'
' DESCRIPTION:
'   find a window that contains the "partialTitle"
'---------------------------------------------------------------------
Function findWindowTitle(partialTitle As String) As String

    Dim currWnd As Long
    Dim length As Long
    Dim listItem As String
    Dim X As Long
    Dim FirstWnd As Long
   
    FirstWnd = AgentWindow
    currWnd = GetWindow(FirstWnd, GW_HWNDFIRST)

    'Loop while the hWnd returned by GetWindow is valid.
    While currWnd <> 0

        'Get the length of task name identified by CurrWnd in the list.
        length = GetWindowTextLength(currWnd)

        'Get task name of the task in the master list.
        listItem$ = Space$(length + 1)
        length = GetWindowText(currWnd, listItem$, length + 1)

        'If there is a task name in the list, add the item to the list.
        If length > 0 Then
            If InStr(LCase$(listItem), LCase$(partialTitle)) <> 0 Then
                findWindowTitle = listItem$
                Exit Function
            End If
        End If

        'Get the next task list item in the master list.
        currWnd = GetWindow(currWnd, GW_HWNDNEXT)

        'Process Windows events.
        DoEvents

    Wend

    findWindowTitle = ""

End Function

'---------------------------------------------------------------------
' SUB: findWindowHandle
'
' DESCRIPTION:
' find a window that contains the "partialTitle"
'---------------------------------------------------------------------
Function findWindowHandle(partialTitle As String) As Long

    Dim currWnd As Long
    Dim length As Long
    Dim listItem As String
    Dim X As Long
    Dim FirstWnd As Long
   
    FirstWnd = AgentWindow
    currWnd = GetWindow(FirstWnd, GW_HWNDFIRST)

    'Loop while the hWnd returned by GetWindow is valid.
    While currWnd <> 0

        'Get the length of task name identified by CurrWnd in the list.
        length = GetWindowTextLength(currWnd)

        'Get task name of the task in the master list.
        listItem$ = Space$(length + 1)
        length = GetWindowText(currWnd, listItem$, length + 1)

        'If there is a task name in the list, add the item to the list.
        If length > 0 Then
            If InStr(LCase$(listItem), LCase$(partialTitle)) <> 0 Then
                findWindowHandle = currWnd
                Exit Function
            End If
        End If

        'Get the next task list item in the master list.
        currWnd = GetWindow(currWnd, GW_HWNDNEXT)

        'Process Windows events.
        DoEvents

    Wend

    findWindowHandle = 0

End Function

'---------------------------------------------------------------------
' SUB: popApp
'
' DESCRIPTION:
'   Pop up a desk application.
'---------------------------------------------------------------------
Sub popApp(partialTitle As String)

   'Resume if "Illegal function call" occurs on AppActivate statement.
    On Error Resume Next
    PopMinimizedWindow (partialTitle)

    AppActivate findWindowTitle(partialTitle)

End Sub

'---------------------------------------------------------------------
' SUB: stuffKeys
'
' DESCRIPTION:
'   Send an application key strokes.
'---------------------------------------------------------------------
Sub stuffKeys(keyStrokes As String)

    'Resume if "Illegal function call" occurs on sendKeys statement.
    On Error Resume Next
    SendKeys keyStrokes

End Sub

'---------------------------------------------------------------------
' SUB: popAndStuff
'
' DESCRIPTION:
'   Pop an application and send the application key strokes.
'---------------------------------------------------------------------
Sub popAndStuff(partialTitle As String, keyStrokes As String)

    'Resume if "Illegal function call" occurs on AppActivate statement.
    On Error GoTo noSendKeys
    PopMinimizedWindow (partialTitle)

    AppActivate findWindowTitle(partialTitle)

    'Resume if "Illegal function call" occurs on sendKeys statement.
    On Error Resume Next
    SendKeys keyStrokes

noSendKeys:
    Exit Sub

End Sub

'---------------------------------------------------------------------
' SUB:PopMinimizedWindow
'
' DESCRIPTION:
'   Pop an application if it is minimized on the tool bar
'---------------------------------------------------------------------
Sub PopMinimizedWindow(partialTitle As String)
        Dim currWnd As Long

        'If the Application is minimized restore it
        currWnd = findWindowHandle(partialTitle)
        If IsIconic(currWnd) Then
            ShowWindow(currWnd, SW_RESTORE)
        End If
   
End Sub

'--------------------------------------------------------------------
' SUB: CreateBrowserForiVault
'
' DESCRIPTION:
'  This procedure is called to create a browser for iVault
'--------------------------------------------------------------------
' ShowWindow() Commands


Sub CreateBrowserForiVault(bMakeTop As Boolean)
    Dim visible As Boolean
    Dim currWnd As Long
    Dim agentWnd As Long
   
    On Error GoTo LoadIE
    ' attempt to get a property in the IE object. if error, then we have no IE object
    visible = IEiVault.Visible
    ' IE object already exists, skip creating
    GoTo ContinueIE
LoadIE:
    ' Create an instance of the IE application
    Set IEiVault = CreateObject("InternetExplorer.Application")

ContinueIE:
    IEiVault.MenuBar = False
    IEiVault.StatusBar = False
    IEiVault.AddressBar = False
    
    ' Get the window handle to the IE browser instance
    currWnd = IEiVault.Application.HWND
    ' If it is minimized, restore the window
    If IsIconic(currWnd) Then
        ShowWindow(currWnd, SW_RESTORE)
    Else
        ' If we do not want IE to be the topmost window,
        '   show it not activated otherwise show it normally
        If bMakeTop = False Then
            ShowWindow(currWnd, SW_SHOWNOACTIVATE)
        Else
            'ShowWindow(currWnd, SW_NORMAL)
            ShowWindow(currWnd, SW_SHOWMAXIMIZED)
            
        End If
    End If
End Sub


'--------------------------------------------------------------------
' SUB:IVaultQuery
'
' Called from SetupMacroUI()
'
' DESCRIPTION:
'    This procedure will pass a query to the IVault tool
'    If a web browser is open it will be used, if the user
'    does not have a browser open the specified browser
'    will be open, and the user will be required to log into
'    IVault
'
' HISTORY:
'
'    12/18/2000 - David Kelly - added this function from Samples
'
'--------------------------------------------------------------------

Sub IVaultQuery(ati As ATInteraction)

    Dim machineName As String
    Dim historylocation As String
    Dim customer As String
    Dim query As String

    ' figure out what machine IVault is on, then use that to figure out its web location

    machineName = InteractionHistoryMachineName
    If Len(machineName) = 0 Then
        'The following line must be updated with the correct machine name
        machineName = "localhost"
    End If


    historylocation = "http://" + machineName + "/IVault/servlet/com.apropos.ihist.servlet.InteractionMainFS"

    'To query IVault, we must contruct string which looks like this:
    'historylocation?name1=value1&name2=value2&name3=value3
    'Here is one example of a query... it searches for all customer names that start with "g" or "G".
    'Remember to Call HTMLEncode for each Value that needs to be searched for.
    customer = ati.Name
    query = historylocation + "?customerNameWildcard=startswith&fldvalueORGANIZATION_LIKE=" + HTMLEncode(customer)

    ' Load the default browser and display the specified page
    debug.print "The iVault URL: " & query

        CreateBrowserForiVault True

    ' Now, navigate to the URL for iVault with parameters
    IEiVault.Navigate(query)


End Sub


'---------------------------------------------------------------------
' Function: HTMLEncode
'
' DESCRIPTION:
'  Changes characters to comply with HTML encoding rules
'---------------------------------------------------------------------
Function HTMLEncode(sValue As String)
    
    Dim tempdata As String
    tempdata = sValue

    If tempdata <> "" Then
    tempdata = Replace(tempdata,"%"," percent")
    tempdata = Replace(tempdata,"@","%40")
    tempdata = Replace(tempdata,"!","%21")
    tempdata = Replace(tempdata,"#","%23")
    tempdata = Replace(tempdata,"$","%24")
    tempdata = Replace(tempdata,"^","%5E")
    tempdata = Replace(tempdata,"&","%26")
    tempdata = Replace(tempdata,"(","%28")
    tempdata = Replace(tempdata,")","%29")
    tempdata = Replace(tempdata,"=","%3D")
    tempdata = Replace(tempdata,"+","%2B")
    tempdata = Replace(tempdata,",","%2C")
    tempdata = Replace(tempdata,"<","%3C")
    tempdata = Replace(tempdata,">","%3E")
    tempdata = Replace(tempdata,"/","%2F")
    tempdata = Replace(tempdata,"?","%3F")
    tempdata = Replace(tempdata," ","%20")
    End If
    HTMLEncode = tempdata


End Function

