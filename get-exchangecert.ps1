$cert = Get-ExchangeCertificate | Where-Object {$_.Services -like "*smtp*" }
$date = $cert.NotAfter
 if($date.Subtract((Get-Date)).Days -le 30) {
   "Critical - Certificate will Expire in " + $date.ToString("dd\/MM\/yyyy")
  }   else {
   "OK - Certificate will Expire in " + $date.ToString("dd\/MM\/yyyy")
  }