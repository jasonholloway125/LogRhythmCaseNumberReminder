#Case Number Notification script v2.2
#By Jason Holloway
#Used for creating Application event logs that list AIE cases that have not been updated with their corresponding ticket numbers
#Intended to be run every friday as a reminder.

#starts logging transcript - creates log dir
function Start-Logging{

    $log_file_name = "case_num_notif_$(Get-Date -Format 'yyyy-MM-dd-hh-mm-ss-fffffff').log";
    $log_path = Join-Path $PSScriptRoot 'logs';

    if(-not (Test-Path -Path $log_path)){
        New-Item -Path $log_path -ItemType Directory;
    }

    $log_file_path = Join-Path $log_path $log_file_name;

    Start-Transcript -Path $log_file_path -NoClobber;
}

#Used for logging - typically goes to cwd\case_num_notif[datetime].log
function Log-Message{
    param (
        [string]$msg
    )

    $tmz = Get-TimeZone | Select -ExpandProperty DisplayName;
    $dateTime = Get-Date -Format 'yyyy-MM-dd hh:mm:ss.fffffff';

    Write-Host "[$tmz - $dateTime] : ""$msg""";
}

#Deletes old log files
function Clear-ExpiredLogs{
    param (
        [object]$cfg
    )

    $log_path = Join-Path $PSScriptRoot 'logs';

    if(-not (Test-Path -Path $log_path)){
        return;
    }

    $limit = (Get-Date).AddDays(-$cfg.log_expiration);

    Get-ChildItem -Path $log_path -Force | Where-Object { $_.Extension.Equals('.log') -and $_.CreationTime -lt $limit } | Remove-Item -Force
}

#Check if today's date is scheduled Date
function Is-ScheduledDate{
	param (
        [object]$cfg
    )
	
    $dates = $cfg.dates;
    $today = (Get-Date).ToString("yyyy/MM/dd");

	return $dates.contains($today);
}

#API request for cases of specified number of days
function Send-Request{
    param (
        [object]$cfg
    )
    
    $date = Get-Date;
    $dateStr = $date.AddDays(-$cfg.search_days).ToString("yyyy-MM-ddT00:00:00Z");


    $headers = @{
        'Authorization' = 'Bearer ' + $cfg.token
        'Content-Type' = 'application/json'
        'count' = $cfg.max_entries + 1
        'createdAfter' = $dateStr
        'direction' = 'desc'
    }

    return Invoke-RestMethod -Uri ($cfg.uri + $cfg.uri_path) -Headers $headers -Method Get
}




#Checks if Application event sources exists - creates one if not
function Validate-EventSource{
    param (
        [object]$cfg
    )

    function Validate{
        param (
            [object]$cfg
        )	

        if(-not [System.Diagnostics.EventLog]::Exists($cfg.event_log)){
		    return $false;
	    }

	    if(-not [System.Diagnostics.EventLog]::SourceExists($cfg.event_source)){
		    [System.Diagnostics.EventLog]::CreateEventSource($cfg.event_source, $cfg.event_log);
	    }
		
	    return $true;
    }

    if($cfg.event_hostname -eq "localhost"){
        return Validate -cfg $cfg;
    }
    else{

    }
	return Invoke-Command -ComputerName $cfg.event_hostname -ScriptBlock ${function:Validate} -ArgumentList $cfg;
}



#Retrieves cases that do not have ticket numbers
function Get-Cases{
    param (
        [object]$cfg,
        [object[]]$resp
    )

	$cases = @();

    foreach($case in $resp){
        if(($case.name -match $cfg.include_regex) -and $case.priority -lt $cfg.min_priority -and -not ($case.name -match $cfg.exclude_regex)){
            $cases += $case;
        }
    }
	
	return $cases;
}

#Writes event logs containing case numbers
function Write-Log{
    param (
        [object]$cfg,
        [Int32]$event_id,
        [string]$message
    )

    if($cfg.event_hostname -eq "localhost"){
        Write-EventLog -LogName $cfg.event_log -Source $cfg.event_source -EventId $event_id -EntryType Information -Message $message;
    }
    else{
        Invoke-Command -ComputerName $cfg.event_hostname -ScriptBlock {
            Param(
                [string]$event_log,
                [string]$event_source,
                [Int32]$event_id,
                [string]$message
            )
            Write-EventLog -LogName $event_log -Source $event_source -EventId $event_id -EntryType Information -Message $message;
        } -ArgumentList $cfg.event_log,$cfg.event_source,$event_id,$message
    }
}

#clear cache
Remove-Variable * -ErrorAction SilentlyContinue;

#retrieve logging file - starts transcript
Start-Logging;

#basically main - loads config file -> clears expired logs -> checks date -> sends/receives api request/response -> writes event logs  
try {
    
    
    Log-Message "Starting Case Number Notification script...";

    $config_file_name = 'config.json';
    $config_file_path = Join-Path $PSScriptRoot $config_file_name;


    $config = Get-Content -Raw $config_file_path | ConvertFrom-Json;
    Log-Message "Successfully loaded config ($config_file_path).";


    Log-Message "Removing expired logs...";
    Clear-ExpiredLogs -cfg $config;


    if (-not (Is-ScheduledDate -cfg $config)){
        $message = "Today is not a scheduled date.";
        Log-Message $message;
        return;
    }


    if(-not (Validate-EventSource -cfg $config)){
        throw "Application event log not found.";
    }
    
    Log-Message "Sending API request...";

    $response = Send-Request -cfg $config;

    Log-Message "Received API response. $($response.Count) results.";


    $cases = Get-Cases -cfg $config -resp $response;

    if(-not($cases)){
        $message = "All ticket numbers sufficed from the last $($config.search_days) days."

        Log-Message $message;

        Write-Log -cfg $config -event_id $config.evid_noresults -message $message;   
    }
    else{

        $message = @{
            "case_number(s)" = ($cases | Select -ExpandProperty number)
        }

        Log-Message ($message | Out-String);

        Write-Log -cfg $config -event_id $config.evid_results -message ($message | ConvertTo-Json);    
    }

}
#logs exceptions
catch {
    $message = "An exception occurred. Please review log at '$PSScriptRoot': '$_'.";

    Log-Message $message;

    if($config){
        Write-Log -cfg $config -event_id $config.evid_error -message $message;
    }
}
#stops transcript
finally {
    Log-Message "Stopping...";
    Stop-Transcript;
}

