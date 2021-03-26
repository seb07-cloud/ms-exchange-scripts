Get-Mailbox -ResultSize Unlimited |
   Get-MailboxStatistics |
   Select DisplayName,StorageLimitStatus, `
   @{name="TotalItemSize (MB)"; expression={[math]::Round( `
   ($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}, `
   ItemCount |
   Sort "TotalItemSize (MB)" -Descending