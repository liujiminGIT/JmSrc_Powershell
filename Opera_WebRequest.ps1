

$JSON = @{
    "username" = "劉積民";
    "text"=$comment
   }
$response = Invoke-WebRequest -Uri "URI" -Method Post -Body $JSON


<#########################################################################>