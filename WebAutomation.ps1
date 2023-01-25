
$r=Invoke-WebRequest https://www.ultimatix.net -SessionVariable fb
$r.links
$form = $r.Forms[0]

$form.Fields[“username”]=”842774”
#$form.Fields[“password”]=”Password”

$r=Invoke-WebRequest -Uri (“https://www.ultimatix.net” + $form.Action) -WebSession $fb -Method POST -Body $form.Fields
