#Case Number Notification Installation script v2.1
#By Jason Holloway
#Used for creating Application event logs that list AIE cases that have not been updated with their corresponding ticket numbers
#Intended to be run every friday as a reminder.

#Used for logging - typically goes to cwd
function Log-Message{
    param (
        [string]$msg
    )

    $tmz = Get-TimeZone | Select -ExpandProperty DisplayName;
    $dateTime = Get-Date -Format 'yyyy-MM-dd hh:mm:ss.fffffff';

    Write-Host "[$tmz - $dateTime] : ""$msg""";
}

#Creates the schtask that runs script to find cases without ticket numbers
function Create-SchTask{
    param (
        [object]$cfg
    )

    $xml_path = Join-Path $PSScriptRoot $cfg.schtask_xml_name;

    $xml = (Get-Content $xml_path | Out-String)

    Register-ScheduledTask -Xml $xml -TaskName $cfg.task_name -Force;
}

#clear cache
Remove-Variable * -ErrorAction SilentlyContinue

#retrieve logging file - starts transcript
$log_file_name = "case_num_notif_install.log";
$log_file_path = Join-Path $PSScriptRoot $log_file_name;

Start-Transcript -Path $log_file_path -Append;


Log-Message "Starting Case Number Notification Installation script...";

$config_file_name = 'install_config.json';
$config_file_path = Join-Path $PSScriptRoot $config_file_name;

$config = Get-Content -Raw $config_file_path | ConvertFrom-Json;

Log-Message "Successfully loaded config ($config_file_path).";

#basically main - checks/creates installation dir - copies files - create schtask
try {
    if(-not (Test-Path -Path $config.seamless_install_path)){
        throw "Seamless Intelligence directory not found."
    }
 
    $install_path = Join-Path $config.seamless_install_path $config.case_num_notif_install_dir;

    if(-not (Test-Path -Path $install_path)){
        Log-Message "Creating installation directory ($install_path)."
        New-Item -Path $install_path -ItemType Directory;
    }

    Log-Message "Copying files to installation directory.";

    Copy-Item (Join-Path $PSScriptRoot $config.script_file_name) -Destination $install_path -Force;
    Copy-Item (Join-Path $PSScriptRoot $config.config_file_name) -Destination $install_path -Force;


    Create-SchTask -cfg $config;

    Log-Message "Created scheduled task ($($config.task_name))";

    Log-Message "Installation completed without exceptions.";
}
catch {
    Log-Message "An exception occurred: '$_'";
}
finally {
    Log-Message "Stopping...";
    Stop-Transcript;

}
