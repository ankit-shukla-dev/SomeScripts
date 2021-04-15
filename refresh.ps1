$endpoint = "https://my.ieltsessentials.com/IELTS?countryId=35&testCentreId=18705&testVenueId=4007&testCentreLocationId=2680&testSessionDate=2021-05-12T09:00:00&lang=en&isSelt=false&restrictToSpecifiedDate=true&testmoduleid=2&ref=v3.testsession.idpOwned&_ga=2.14585335.1946166389.1618449012-410471934.1618449012"

# Create a new SpVoice objects
$voice = New-Object -ComObject Sapi.spvoice

# Set the speed - positive numbers are faster, negative numbers, slower
$voice.rate = 0

$maxtry = 3600
$iter = 0
$sleepDuration = 5

do{
 write-host "try number - $iter"
 sleep $sleepDuration
 try{
	$response = Invoke-WebRequest $endpoint
	$voice.speak("hurray, it's done")
	break
 }
 catch{
  $_.Exception.Response.StatusCode.Value__
 }
 
 $iter += 1
}while($iter -le $maxtry)


