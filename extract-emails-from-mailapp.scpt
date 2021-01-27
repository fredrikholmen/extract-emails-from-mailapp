set emailList to {}
set recipientList to {}
set senderList to {}

# Iterate through all mails in the selected mailbox
tell application "Mail" to repeat with aMessage in (get (get selection)'s item 1's mailbox's messages)
	
	set senderMail to extract address from sender of aMessage
	if senderMail is not in emailList then
		set end of emailList to senderMail
		set senderName to my processName(extract name from sender of aMessage)
		set end of senderList to senderName & tab & senderMail
	end if
	
	set aList to {}
	repeat with aRecipient in recipients of aMessage
		set email to address of aRecipient
		if email is not in emailList then
			set end of emailList to email
			set recipientName to my processName(name of aRecipient)
			set end of aList to recipientName & tab & email
		end if
	end repeat
	if aList is not {} then
		set end of recipientList to aList
	end if
end repeat

# Setup and create files on the desktop with the results
set {TID, text item delimiters} to {text item delimiters, return}
set resultText to recipientList as text
set senderText to senderList as text
set text item delimiters to TID

set textFile to ((path to desktop as text) & "recipients.txt")
try
	set fileRef to open for access file textFile with write permission
	write resultText to fileRef
	close access fileRef
on error
	try
		close access file textFile
	end try
end try

set textFile to ((path to desktop as text) & "senders.txt")
try
	set fileRef to open for access file textFile with write permission
	write senderText to fileRef
	close access fileRef
on error
	try
		close access file textFile
	end try
end try

# Fix tabulation of first and last name
on processName(theString)
	if theString is missing value then return "n/a"
	set {TID, text item delimiters} to {text item delimiters, space}
	set countTextItems to count text items of theString
	tell text items of theString
		if countTextItems > 2 then
			set theResult to ((items 1 thru -2) as text) & tab & last item
		else if countTextItems = 2 then
			set theResult to item 1 & tab & item 2
		else
			set theResult to theString
		end if
	end tell
	set text item delimiters to TID
	return theResult
end processName
