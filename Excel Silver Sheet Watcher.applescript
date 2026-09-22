property wasSilverActive : false
property lastRefresh : 0

on run
	set wasSilverActive to false
	set lastRefresh to 0
end run

on idle
	
	set silverActive to false
	set frontApp to ""
	set activeBookName to ""
	
	-- Determine which application is currently frontmost
	tell application "System Events"
		set frontApp to name of first application process whose frontmost is true
	end tell
	
	-- Only inspect Excel if Excel is actually frontmost
	if frontApp is "Microsoft Excel" then
		
		tell application "Microsoft Excel"
			try
				set activeBookName to name of active workbook
				
				if activeBookName is "Silver Stack Sheet.xlsm" then
					set silverActive to true
				end if
				
			end try
		end tell
		
	end if
	
	if silverActive then
		
		set currentTime to current date
		
		-- Silver workbook has just become frontmost
		if wasSilverActive is false then
			
			my refreshSilver()
			set lastRefresh to current date
			
			-- Silver workbook has remained frontmost for 60 seconds
		else if (currentTime - lastRefresh) ³ 60 then
			
			my refreshSilver()
			set lastRefresh to current date
			
		end if
		
	else
		
		-- Workbook is no longer active, so discard the timer
		set lastRefresh to 0
		
	end if
	
	set wasSilverActive to silverActive
	
	-- Check frontmost application every second
	return 1
	
end idle

on refreshSilver()
	
	tell application "Microsoft Excel"
		run VB macro "WatcherRefresh"
	end tell
	
end refreshSilver