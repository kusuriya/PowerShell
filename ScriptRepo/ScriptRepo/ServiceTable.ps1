
param ( [string]$txt1 = $null )

$var1=[System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain() 
$dom=$var1.name.split(".")[0]
if ($txt1 -eq $null){
$txt1 = "c:\scripts\_input\" +  $dom + "_serverlist.txt"
}
$header = "WSUS Service Status."
$footer = "fin"
$title = "Report for services"
$head = '
<head>
<style>
BODY{font-family:Verdana; background-color:lightblue;}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH{font-size:1.3em; border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:#FFCCCC}
TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:yellow}
</style>
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script>
<script type="text/javascript">
$(function(){
  var linhas = $("table tr");
  $(linhas).each(function(){
   var Valor = $(this).find("td:last").html();
   if(Valor === "Stopped"){
    $(this).find("td").css("background-color","Red");
   }else if(Valor === "Running"){
    $(this).find("td").css("background-color","Green");
   }
  });
});
</script>
'
$head = $head + "<title>$title</title>"
$head = $head + '</head><body>'
$computerlist = get-content $txt1
$Gather = foreach ($computer in $computerlist) {
$computer
Get-Service -computername $computer -name "wuauserv"
}
$Gather |select MachineName,Name,DisplayName,Status|Sort -Descending Status|ConvertTo-Html -Head $head -PreContent $header -PostContent $footer|Out-File test.html