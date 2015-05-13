#Global variables
#Set working directory here
$workdir = "c:\scripts\"
#Quote character " for use as a CSV separator since some fields use ,.: and ;
$sc = [char]34

#Read in a list of computers from a text file. One entry per line.  Uncomment if this method is desired.
$computers = get-content -path $workdir"machinelist.txt"


#Read in computer objects from Active Directory.  Uncomment if this method is desired. Requires that you have imported the Powershell Active Directory Module
#Replace the OU structure to match the OU structure of your domain.
#$computers = get-adcomputer -filter * -searchbase "OU=Computers ,DC=busboy, DC=org" |select name





if (-not(test-path -Path "$workdir$((Get-Date).ToString('yyyyMMddHH00'))"))
{
	New-Item -ItemType Directory -Path "$workdir$((Get-Date).ToString('yyyyMMddHH00'))"
}
 

$computers = get-content -path $workdir"machinelist.txt"

$array = @()

foreach($computer in $computers){

if (Test-Connection -ComputerName $computer -Quiet -count 1)
{
#Read in 64bit Applications

    $UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"

    $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computer)

    $regkey=$reg.OpenSubKey($UninstallKey)

    $subkeys=$regkey.GetSubKeyNames()

    foreach($key in $subkeys){

        $thisKey=$UninstallKey+"\\"+$key

        $thisSubKey=$reg.OpenSubKey($thisKey)

        $obj = New-Object PSObject

        $obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computer

        $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))

        $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))

        $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))

        $obj | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value $($thisSubKey.GetValue("InstallDate"))

        $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))

#You can add other key values here if desired

#Add results to the array

        $array += $obj


    }

#Read in 32bit Applications

    $UninstallKey="SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"

    $regkey=$reg.OpenSubKey($UninstallKey)

    $subkeys=$regkey.GetSubKeyNames()

    foreach($key in $subkeys){

        $thisKey=$UninstallKey+"\\"+$key

        $thisSubKey=$reg.OpenSubKey($thisKey)

        $obj = New-Object PSObject

        $obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computer

        $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))

        $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))

        $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))

        $obj | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value $($thisSubKey.GetValue("InstallDate"))

        $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))

#You can add other key values here if desired

#Add results to the array

        $array += $obj


    }
	
	#Output the results of the array. Default is to the screen, but thanks to .Net, we can push the data any where!

$array | Where-Object { $_.DisplayName } | select ComputerName, DisplayName, DisplayVersion, Publisher, InstallDate | export-csv -Path "$workdir$((Get-Date).ToString('yyyyMMddHH00'))\$computer.csv"-notypeinformation -delimiter "$sc"

}

Else
{write-host $computer not found}
}


