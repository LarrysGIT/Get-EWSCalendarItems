
PARAM(
    [string]$MailAddress,
    [datetime]$StartTime,
    [datetime]$EndTime,
    [string]$EWSUrl,
    [string]$ExchangeVersion,
    [System.Net.NetworkCredential]$Credential,
    [int]$ResultSize
)

<# Key parts, below is an example removed unnecessary lines
Import-Module .\Microsoft.Exchange.WebServices.dll
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$FolderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, 'larry@consoto.com')
$StartTime = [datetime]'2016-02-22'
$EndTime = [datetime]'2016-03-33'
$CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView('2016-02-22', $EndTime)
$Service.AutodiscoverUrl($MailAddress)
$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $FolderID)
$Service.FindAppointments($Calendar.Id, $CalendarView)
#>

Set-Location (Get-Item ($MyInvocation.MyCommand.Definition)).DirectoryName

if(Import-Module .\Microsoft.Exchange.WebServices.dll -ErrorAction:SilentlyContinue)
{
    Write-Host 'The dll file Microsoft.Exchange.WebServices.dll is not found'
    Write-Host 'Please download EWS API from https://www.microsoft.com/en-us/download/details.aspx?id=35371'
    exit 1
}

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion
if(!$?)
{
    Write-Host 'You need to supply a correct exchange version'
    Write-Host 'You can find exchange version string from https://msdn.microsoft.com/en-us/library/office/microsoft.exchange.webservices.data.exchangeversion(v=exchg.80).aspx'
    exit 1
}

$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
if($Credential)
{
    $Service.Credentials = $Credential
}

$FolderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $MailAddress)

if($ResultSize)
{
    $CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartTime, $EndTime, $ResultSize)
}
else
{
    $CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartTime, $EndTime)
}
if(!$?)
{
    Write-Host 'Error when creating calendar view object'
    Write-Host 'Makes sure you supplied valid "StartTime" and "EndTime", if "ResultSize" provided, please greater than 0'
    exit 1
}

if($EWSUrl)
{
    $Service.Url = $EWSUrl
}
else
{
    $Service.AutodiscoverUrl($MailAddress)
    if(!$?)
    {
        Write-Host 'Does your AudoDiscover functionality normal?'
        Write-Host 'You may need to supply EWSUrl manually'
        Write-Host 'EWSUrl example: https://mail.consoto.com/ews/exchange.asmx'
        exit 1
    }
}

$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $FolderID)

return $Service.FindAppointments($Calendar.Id, $CalendarView)
