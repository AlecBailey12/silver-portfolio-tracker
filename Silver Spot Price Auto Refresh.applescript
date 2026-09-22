on getSilverPrice(dummy)
	
	set apiURL to "https://api.gold-api.com/price/XAG"
	set jsonText to do shell script "/usr/bin/curl -s " & quoted form of apiURL
	
	set priceText to do shell script "/usr/bin/python3 -c " & quoted form of Â
		"import json,sys; print(json.loads(sys.stdin.read())['price'])" & Â
		" <<'EOF'" & linefeed & jsonText & linefeed & "EOF"
	
	return priceText
	
end getSilverPrice


on startWatcher(dummy)
	
	set watcherPath to (POSIX path of (path to library folder from user domain)) & "Application Scripts/com.microsoft.Excel/Excel Silver Sheet Watcher.app"
	
	do shell script "/usr/bin/open -g " & quoted form of watcherPath
	
	return "OK"
	
end startWatcher


on stopWatcher(dummy)
	
	try
		tell application "Excel Silver Sheet Watcher" to quit
	end try
	
	return "OK"
	
end stopWatcher