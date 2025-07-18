# Define task details
$TaskName = "RunPythonEmailExtractor"
$TaskDescription = "Extract email tables and convert to CSV"
$PythonExe = "C:\Python311\python.exe"
$ScriptPath = "C:\Scripts\ExtractEmailTables.py"
$TriggerTime = (Get-Date).AddMinutes(1)  # Schedule 1 minute from now

# Create action to run Python script
$Action = New-ScheduledTaskAction -Execute $PythonExe -Argument $ScriptPath

# Create daily trigger at a specific time (e.g., 10:00 AM)
$Trigger = New-ScheduledTaskTrigger -Daily -At 10:00AM

# Register the scheduled task
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Description $TaskDescription -User "$env:USERNAME" -RunLevel Highest
