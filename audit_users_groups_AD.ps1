[reflection.assembly]::loadwithpartialname("Microsoft.Office.Interop.Excel")
$date = Get-Date -format "dd-MM-yyyy" #dd-mm-yyyy

#---------------------------------------------------------------------------------------
# Excel Instance
#---------------------------------------------------------------------------------------
$xl = New-Object -ComObject "Excel.Application"
$xl.Visible = $True
$wb=$xl.Workbooks.Add()
$ws=$wb.ActiveSheet
$ws.PageSetup.CenterHorizontally = $True
$ws.PageSetup.LeftMargin = 15
$ws.PageSetup.RightMargin = 15
$ws.PageSetup.TopMargin = 15
$ws.PageSetup.Bottommargin = 15
$ws.PageSetup.Zoom = $false
$ws.PageSetup.FitToPagesWide = 1
$i=1

#---------------------------------------------------------------------------------------
# Sheet title
#---------------------------------------------------------------------------------------
$titre = "YOUR TITLE (" + $date + ")"
$xl.cells.item($i,1) = $titre
$xl.cells.item($i,1).font.bold = $True
$xl.cells.item($i,1).font.size = 24
$xl.cells.item($i,1).HorizontalAlignment = -4108
$select = $ws.range("a1","c1")
$select.MergeCells = $True
$i++
$i++

#---------------------------------------------------------------------------------------
# Description Title
#---------------------------------------------------------------------------------------
$titre = "YOUR TITLE"
$xl.cells.item($i,1) = $titre
$xl.cells.item($i,1).font.bold = $True
$xl.cells.item($i,1).font.size = 16
$i++


#---------------------------------------------------------------------------------------
# Columns
#---------------------------------------------------------------------------------------
$startLine = $i
$xl.columns.item(1).columnWidth = 30
$xl.cells.item($i,1).HorizontalAlignment = -4108
$xl.cells.item($i,1) = "User"
$xl.cells.item($i,1).font.bold = $True
$xl.cells.item($i,2).HorizontalAlignment = -4108
$xl.columns.item(2).columnWidth = 35 
$xl.cells.item($i,2) = "Groups"
$xl.cells.item($i,2).font.bold = $True
$xl.cells.item($i,3).HorizontalAlignment = -4108
$xl.columns.item(3).columnWidth = 35 
$xl.cells.item($i,3) = "Site"
$xl.cells.item($i,3).font.bold = $True
$xl.cells.item($i,4).HorizontalAlignment = -4108
$xl.columns.item(4).columnWidth = 15
$xl.cells.item($i,4) = "Associated to"
$xl.cells.item($i,4).font.bold = $True
$xl.cells.item($i,5).HorizontalAlignment = -4108
$xl.columns.item(5).columnWidth = 35 
$xl.cells.item($i,5) = "Member of following group"
$xl.cells.item($i,5).font.bold = $True
$i++

$location = "OU=XXX,dc=XXX,dc=XXX,dc=XXX"
[array]$OUlist = Get-ADOrganizationalUnit -Filter {name -like "*S-*"} -SearchBase $location | Select name, DistinguishedName
foreach ($ou in $OUlist) {
    $name = $ou.name
    $site = $ou.DistinguishedName.Split(',')[1].split('=')[1]
    $searchBase = "OU=$name,OU=$site,OU=C-DPN,dc=atlas,dc=edf,dc=fr"

$request = Get-ADUser -Filter {name -like "XXX"} -SearchBase $searchBase.ToString() -Properties name, userWorkstations, DistinguishedName, MemberOf, SamAccountName | select `    @{e={$_.Name};l="User"},`    @{e={$_.DistinguishedName.split(',')[2].split('=')[1]}; l="Site"},`    @{e={$_.userWorkstations -replace ",","`n"};l="Associated to"},`    @{e={$_.MemberOf -split(',') -join "`n"};l="Membre of"},`    SamAccountName

  foreach ($users in $request) {  
        $userGroupTab = @()
        $userGroupRequest = Get-ADUser $users.SamAccountName -Properties Memberof | %{$_.memberof} | %{get-adgroup $_ | select -ExpandProperty name}
            foreach($userGroup in $userGroupRequest) {
                $userGroupTab+=$userGroup
            }  
                
    if ($users.'Associated to'.Length -gt 10) {
        $origin = @($users.'Associated to' -replace ',',' ')
        $valeur = $origin -split '\s+'
            foreach ($value in $valeur) {
                        foreach ($computer in $users.'Associated to'){
                            try {
                                $computerGroupRequest = Get-ADComputer $value.ToString() -Properties Memberof | %{$_.memberof} | %{get-adgroup $_ | select -ExpandProperty name} -ErrorAction Stop
                            }
                            catch {
                                $computerGroupRequest = "Computer not found"
                            }
                
                        }
                            $xl.cells.item($i,1).VerticalAlignment = -4108
                            $xl.cells.item($i,1).HorizontalAlignment = -4108
                            $xl.cells.item($i,1) = $users.User
                            $xl.cells.item($i,2).VerticalAlignment = -4108
                            $xl.cells.item($i,2).HorizontalAlignment = -4108
                            $xl.cells.item($i,2) = $userGroupTab -split(',') -join "`n"
                            $xl.cells.Item($i,3).VerticalAlignment = -4108
                            $xl.cells.item($i,3).HorizontalAlignment = -4108
                            $xl.cells.item($i,3) = $users.Site
                            $xl.cells.item($i,4).HorizontalAlignment = -4108
                            $xl.cells.Item($i,4).VerticalAlignment = -4108
                            $xl.cells.item($i,4) = $value
                            $xl.cells.item($i,5).HorizontalAlignment = -4108
                            $xl.cells.Item($i,5).VerticalAlignment = -4108
                            $xl.cells.item($i,5) = $computerGroupRequest -split(',') -join "`n"
                            $i++
            }
            
  }
    else {
        $computerGroupTab = @()
            foreach ($computer in $users.'Associated to'){
                try {
                    $computerGroupRequest = Get-ADComputer $users.'Associated to'.ToString() -Properties Memberof | %{$_.memberof} | %{get-adgroup $_ | select -ExpandProperty name} -ErrorAction Stop
                }
                catch {
                    $computerGroupRequest = "Computer not found"
                }
            }
                            $xl.cells.item($i,1).VerticalAlignment = -4108
                            $xl.cells.item($i,1).HorizontalAlignment = -4108
                            $xl.cells.item($i,1) = $users.User
                            $xl.cells.item($i,2).VerticalAlignment = -4108
                            $xl.cells.item($i,2).HorizontalAlignment = -4108
                            $xl.cells.item($i,2) = $userGroupTab -split(',') -join "`n"
                            $xl.cells.Item($i,3).VerticalAlignment = -4108
                            $xl.cells.item($i,3).HorizontalAlignment = -4108
                            $xl.cells.item($i,3) = $users.Site
                            $xl.cells.item($i,4).HorizontalAlignment = -4108
                            $xl.cells.Item($i,4).VerticalAlignment = -4108
                            $xl.cells.item($i,4) = $users.'Associated to'
                            $xl.cells.item($i,5).HorizontalAlignment = -4108
                            $xl.cells.Item($i,5).VerticalAlignment = -4108
                            $xl.cells.item($i,5) = $computerGroupRequest -split(',') -join "`n"
                            $i++                
    } 

  }
}

$endLine = $i -1
$listObject = $ws.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $ws.range("a$startLine","e$endLine"), $null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes,$null)
$listObject.TableStyle = "TableStyleMedium6"
$i++
$i++