param([switch]$Elevated)

#Line 4-16 is checking for admin, you potentially may need elevated permissions to install and configure odbc.
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) {
        # tried to elevate, did not work, aborting
    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    exit
}

#Asking user for credential inputs. Creds will be used later to configure ODBC. 
#Note: User or Proxy account that is being entered must have access to Database otherwise ODBC connection will not work.
$pgServer = Read-Host "Input Postgres Server Name "
$pgDatabase = Read-Host "Input Postgres Database Name "
$creds= Get-credential  #<-- User Input for credentials
$pgusername= $creds.UserName  

$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.Password)
$pgpassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Downloading specific version of postgres odbc. If you need a different version get the version and change here.
$downloadVersion = 'psqlodbc_16_00_0000-x64.zip' #<---- Change version as needed.
$downloadDirectory ='c:\temp\'+$downloadVersion
$uri = 'https://ftp.postgresql.org/pub/odbc/versions/msi/' + $downloadVersion


$odbcName = 'PGODBC'
$driverName = 'PostgreSQL Unicode(x64)'


#Check if Driver Exists 

$drivers = Get-Odbcdriver -Platform '64-bit' | Where-Object Name -Like "Postgres*" 

$dExists = 0 
foreach ($d in $drivers)
{
    IF ($d.Name -eq $driverName)
    {
        $dExists=1
    }
}



#Install If it does not exists. 
IF ($dExists=0)
{

    #Download postgres file latest
    IF (Test-Path $downloadDirectory)
    {
        Remove-Item $downloadDirectory
    }

    Invoke-WebRequest -Uri $uri -Outfile $downloadDirectory


    $expandDirectory ='c:\temp\pgodbc\'

    IF (Test-Path $expandDirectory)
    {
        Remove-Item $expandDirectory
    }

    New-Item -Path 'c:\temp\' -Name "pgodbc" -ItemType "directory"

    #unarchive
    Expand-Archive -Path $downloadDirectory -DestinationPath $expandDirectory


    IF (Test-Path $expandDirectory)
    {
        $msiFile = (Get-ChildItem -Path $expandDirectory | Where-Object { $_.Extension -eq '.msi' }).Name 

        
        $installFile = $expandDirectory + $msiFile

        
        Start-Process msiexec.exe -Wait -ArgumentLIst '/I $installFile /quiet ACCEPT_EULA=TRUE'

    }

}else{
    Write-Host "Driver already installed" -BackgroundColor Red
    Write-Host ""
}



#Once ODBC have been downlaoded, it need to be configured. Check If ODBC Configuration Exists
$listODBCDsn = Get-OdbcDsn
$configure=1


foreach($l in $listODBCDsn)
{

    $attribute = $l.Attribute
    #$attribute 
    $lServer = $attribute.Servername
    $lData= $attribute.Database

    IF (($lServer -eq $pgServer) -and ($lData -eq $pgDatabase))
    {
        Write-Host ''
        Write-Host 'Connector already exists. Exiting configuratoin. ' -BackgroundColor RED
        Write-Host ''

        $configure=0
        break
    
    }

}

IF ($configure -eq 1)
{
    Write-Host "Confuring Postgres ODBC Connector" -BackgroundColor Green

    Add-OdbcDsn -Name $odbcName -DriverName $driverName -DsnType "System" -SetPropertyValue @("Server=$pgServer", "Trusted_Connection=Yes", "Database=$pgDatabase","UserName=$pgusername","Password=$pgpassword")
}   


PAUSE
