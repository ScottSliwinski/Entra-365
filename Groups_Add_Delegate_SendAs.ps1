Connect-ExchangeOnline

$GroupEmail = "CommServShiftSupervisors@awmi.net"
$DelegateUPN = "ezekielnabor@awmi.net"

Add-RecipientPermission `
    -Identity $GroupEmail `
    -Trustee $DelegateUPN `
    -AccessRights SendAs `
    -Confirm:$false
	