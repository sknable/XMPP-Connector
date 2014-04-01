package require java

proc MapChatQueueName { }  {
    global AppData
    variable sQueue ""
    set sReturnValue "OK"

    set sReasonTag [$AppData getValue "REASON_TAG"]

    set result [queuealias -timeout 30000 -itype WEB_CHAT -alias $sReasonTag -result sQueue]
    if {$result != "SUCCEEDED"} {
        puts "Error getting queue alias for reason tag $sReasonTag"
        return $sReturnValue
    }

    set listResult ""
    set result [sysinfo -resultlist listResult AGENT_INFO $sQueue]
    if {$result == "SUCCEEDED"} {
        set sQueueResult [lindex [lindex $listResult 0] 1]
        set nLoggedInAgents [lindex [lindex $listResult 0] 3]
        if {$sQueueResult == "SUCCEEDED"} {
            if {$nLoggedInAgents == "0"} {
                set sQueue "Status - Web Chat"
            }
        }
    }
    $AppData setValue "QUEUE" $sQueue
    return $sReturnValue
}

proc MapRequestQueueName { }  {
    global AppData
    variable sQueue ""
    set sReturnValue "OK"

    set sReasonTag [$AppData getValue "REASON_TAG"]

    set result [queuealias -timeout 30000 -itype WEB_CALL_BACK -alias $sReasonTag -result sQueue]
    if {$result == "SUCCEEDED"} {
        $AppData setValue "QUEUE" $sQueue
    } else {
        puts "Error getting queue alias for reason tag $sReasonTag"
    }

    return $sReturnValue
}

proc CheckSAA {} {

	global AppData
	set saa [$AppData getValue "SAA"]
	
	if { $saa == "TRUE" } {
	
		$AppData setValue "AGENTID" 70
	
	}
	

}


