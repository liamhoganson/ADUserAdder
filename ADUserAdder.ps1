# Import AD Module
Import-Module ActiveDirectory

# new_users .csv filepath
$csvfile = "path\new_users.csv"

# Import file into variable
# Make sure the file path is valid
if ([System.IO.File]::Exists($csvfile)) {
    Write-Host "Importing CSV..."
    $CSV = Import-CSV -LiteralPath $csvfile
} else {
       Write-Host "File Path specified was not valid..."
       Exit
}

# Gets domain credentials
$cred = Get-Credential

# DN Variables
$DEPT1 = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
$DEPT2 = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
$DEPT3 = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
$DEPT4 = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
$DEPT5 = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
$DEPT6 = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
$DEPT7 = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")
$DEPT8 = ("OU=Users,OU=Dept,OU=Comp,DC=DC,DC=DC")


# Grabs Default password from User
$getpassword = Read-Host -Prompt "Enter default New Hire Password"
$SecurePassword = ConvertTo-SecureString $getpassword -AsPlainText -Force

# Iterate through each line in CSV File
ForEach($user in $CSV){
    # Setup Variables 
   $fname = $user.'First'
   $lname = $user.'Last'
   $fullname = "$($user.'First') $($user.'Last')"
   $useremail = $user.'EMAIL'
   $logonname = "$($fname).$($lname)".ToLower()
   $title = $user.'Title'
   $SBU = $user.'DEPT'

   # Function that Creates AD Users
    function Create-ADUser{
        [CmdletBinding()]
        param(
            [parameter(Mandatory)]
            [string] $DN,
            [parameter(Mandatory)]
            [string] $template
        )
        $CheckADUser = Get-ADUser -Filter {UserPrincipalName -eq $useremail}
        if ($CheckADUser -eq $null){
            New-ADUser -Credential $cred `
            -Name $fullname `
            -GivenName $fname `
            -SurName $lname `
            -SamAccountName $logonname `
            -UserPrincipalName $useremail `
            -DisplayName $fullname `
            -AccountPassword $SecurePassword `
            -Description $title `
            -Office $SBU `
            -Enabled $True `
            -path $DN
            Set-ADUser -Identity $logonname -Title $title -EmailAddress $useremail -Department $SBU -Company "Company" 
            $source = Get-ADPrincipalGroupMembership -Identity $template 
            Add-ADPrincipalGroupMembership -Credential $cred -Identity $logonname -MemberOf $source
            Write-Host "Successfully Added User: " $fullname
        }
        else{
            Write-Host "$fullname Already Exists! Skipping..."
        }
        }
       
    # Logic for diffrent DEPTs
    if ($user.'DN' -eq $DEPT1){
        Create-ADUser -DN $DEPT1 -template "dept1templateuser"
    } 
    elseif ($user.'DN' -eq $DEPT2){
        Create-ADUser -DN $DEPT2 -template "dept2templateuser"
    }
    elseif ($user.'DN' -eq $DEPT3){
        Create-ADUser -DN $DEPT3 -template "dept3templateuser"
    }
    elseif ($user.'DN' -eq $DEPT4){
        Create-ADUser -DN $DEPT4 -template "dept4templateuser"
    }
    elseif ($user.'DN' -eq $DEPT5){
        Create-ADUser -DN $DEPT5 -template "dept5templateuser"
    }
    elseif ($user.'DN' -eq $DEPT6){
        Create-ADUser -DN $DEPT6 -template "dept6templateuser"
    }
    elseif ($user.'DN' -eq $DEPT7){
        Create-ADUser -DN $DEPT7 -template "dept7templateuser"
    }
    elseif ($user.'DN' -eq $DEPT8){
        Create-ADUser -DN $DEPT8 -template "dept8templateuser"
    }
    elseif ($user.'DN' -eq 'pass'){
        Write-Host "Pass"
    }
}
