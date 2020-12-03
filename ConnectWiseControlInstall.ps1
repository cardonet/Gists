# Installing ConnectWise Control
Param(
    [Parameter(Mandatory=$true)]
    [string]$HostName,
    [Parameter(Mandatory=$true)]
    [string]$CompanyName,
    [Parameter(Mandatory=$true)]
    [string]$SiteName
 
 )
 $method = 'set'
 
 [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
 Add-Type -AssemblyName System.Web
 Function Get-CWControlServerInfo
 {
     param([Uri]$Uri)
     $Builder = New-Object "System.UriBuilder" $Uri
     $Builder.Path = "/Script.ashx"
     $Info = @{}
     $UriString = $Builder.ToString()
     $Response = Invoke-WebRequest -UseBasicParsing -Uri $UriString
     if($Response.Content -match '\"k"\:\"(.+?)\"')
     {
         $Info.PublicKey = $matches[1]
     }
 
     if($Response.Content -match '\"instanceUrlScheme"\:\"sc-(.+?)\"')
     {
         $Info.InstanceId = $matches[1]
     }
 
     if($Response.Content -match '\"p"\:(\d+)')
     {
         $Info.Port = $matches[1]
     }
     [pscustomobject]$Info
 }
 
 
 $ControlUriBuilder = New-Object "System.UriBuilder" $HostName
 if($Port)
 {
     $ControlUriBuilder.Port = $Port
 }
 else
 {
     $ControlUriBuilder.Port = 443
 }
 $ControlUriBuilder.Scheme = "https"
 $ControlUri = $ControlUriBuilder.ToString()
 $ControlInstanceInfo = Get-CWControlServerInfo -Uri $ControlUri
 
 $Parameters = [System.Web.HttpUtility]::ParseQueryString([String]::Empty)
 $Parameters['e'] = "Access"
 $Parameters['y'] = "Guest"
 $Parameters['h'] = $HostName
 $Parameters['p'] = $ControlInstanceInfo.Port
 $Parameters.Add('c', $CompanyName)
 $Parameters.Add('c', $SiteName)
 $Parameters['k'] = $ControlInstanceInfo.PublicKey
 $Params = $Parameters.ToString();
 
 if(!$ControlInstanceInfo.PublicKey)
 {
     Write-Error "Unable to retrieve publickey from $HostName`:$Port"
     return
 }
 

 switch ($method) {
     "get" {
         $ControlServiceName = "ScreenConnect Client ($($ControlInstanceInfo.InstanceId))"
         Get-Service -Name $ControlServiceName | Format-List *
     }
     "set" {
         $InstallerLogFile = New-TemporaryFile
         $ControlUriBuilder.Path = "/Bin/ConnectWiseControl.ClientSetup.msi"
         $ControlUriBuilder.Query = $Params
 
         $InstallerUri = $ControlUriBuilder.ToString()
         $InstallerFile = [IO.Path]::ChangeExtension([IO.Path]::GetTempFileName(),".msi")
         Get-Package "ScreenConnect Client ($($ControlInstanceInfo.InstanceId))" -ErrorAction SilentlyContinue | Uninstall-Package
         (New-Object System.Net.WebClient).DownloadFile($InstallerUri, $InstallerFile)
         $Arguments = @"
 /c msiexec /i "$InstallerFile" /qn /norestart /l*v "$InstallerLogFile" REBOOT=REALLYSUPPRESS SERVICE_CLIENT_LAUNCH_PARAMETERS="$Params"
"@
         Write-Host "InstallerLogFile: $InstallerLogFile"
         $Process = Start-Process -Wait cmd -ArgumentList $Arguments -Passthru
         if($Process.ExitCode -ne 0)
         {
             Get-Content $InstallerLogFile -ErrorAction SilentlyContinue | select -Last 100
         }
         Write-Host "Exit Code: $($Process.ExitCode)";
         $ControlService = Get-Service -Name "ScreenConnect Client ($($ControlInstanceInfo.InstanceId))"
         if($ControlService.Status -ne "Running")
         {
             $ControlService | Start-Service -Passthru
         }
     }
     "test" {
         $ControlService = Get-Service -Name "ScreenConnect Client ($($ControlInstanceInfo.InstanceId))" -ErrorAction SilentlyContinue  
         return ($null -ne $ControlService -and $ControlService.Status -eq "Running")
     }
 }
