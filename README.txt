# Get-EWSCalendarItems

Credit to https://emg.johnshopkins.edu/?p=837

Example:
	Get-EWSCalendarItems -MailAddress larry@consoto.com -StartTime (Get-Date).AddDays(-1) -EndTime (Get-Date) -ExchangeVersion Exchange2013
	Get-EWSCalendarItems -MailAddress larry@consoto.com -StartTime  2016-02-01 -EndTime 2016-03-31 -ExchangeVersion Exchange2013

Download:
	EWS API - https://www.microsoft.com/en-us/download/details.aspx?id=35371

 - Larry.Song@outlook.com