clear

Try{
    Invoke-Sqlcmd}
Catch{
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Install-Module Sqlserver -Force}

Class BadFile{
    [String]$computerName
    [String]$userName
    [String]$ipAddress
    [String]$filePath
    [String]$fileName
    [String]$fileSize

    [void]SetComputerName([String]$suppliedComputerName){
        $this.computerName = $suppliedComputerName
    }
    [String]GetComputerName(){
        return $this.computerName
    }
    [void]SetUserName([String]$suppliedUserName){
        $this.userName = $suppliedUserName
    }
    [String]GetUserName(){
        return $this.userName
    }
    [void]SetIpAddress([String]$suppliedIpAddress){
        $this.ipAddress = $suppliedIpAddress
    }
    [String]GetIpAddress(){
        return $this.ipAddress
    }
    [void]SetFilePath([String]$suppliedFilePath){
        $this.filePath = $suppliedFilePath
    }
    [String]GetFilePath(){
        return $this.filePath
    }
    [void]SetFileName([String]$suppliedFileName){
        $this.fileName = $suppliedFileName
    }
    [String]GetFileName(){
        return $this.fileName
    }
    [void]SetFileSize([String]$suppliedFileSize){
        $this.fileSize = $suppliedFileSize
    }
    [String]GetFileSize(){
        return $this.fileSize
    }
}

#Microsoft deprecated the ArrayList and replaced it with List. Anything that was an ArrayList (has ArrayList in its name) is now a List.

#The host name.
$theDesiredComputer = ($env:COMPUTERNAME | Out-String).Trim()
#A list of nicely parsed IP addresses that the host could have.
$possibleAddresses = New-Object System.Collections.Generic.List[String]
#Will hold the same data as $filesScannedArrayList but will not be parsed.
$filesScannedArray = Get-ChildItem -Name -Path "C:\Users\" -Recurse
#Holds a list of nicely parse file locations on the computer.
$filesScannedArrayList = New-Object System.Collections.Generic.List[String]
#Holds a list of files that are a foribidden file type.
$badFiles = New-Object System.Collections.Generic.List[String]
#Holds a list of extensions that the script will search for.
$searchTerms = New-Object System.Collections.Generic.List[String]
#Will hold the same data as $oldFlagIDsArrayList but will not be parsed.
$oldFlagIDs = ((Invoke-Sqlcmd -Query "SELECT [FlagID] FROM FlaggedItems WHERE ComputerName = '$theDesiredComputer'" -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*') | Format-Wide | Out-String).Trim()
#Hold a list of FlagIDs that correspond to entries in a table that have NOT been verified to STILL EXIST on the computer.
$oldFlagIDsArrayList = New-Object System.Collections.Generic.List[String]

#Checks each item in the ArrayList for whether or not they contain a certain string.
function FindTheBadChildren{
    PARAM([System.String] $searchTerm, [System.Collections.ArrayList] $children)
    ForEach($child in $children){

        IF($child -Like $searchTerm){
            $badFiles.Add($child) | Out-Null
        }
    }
}
#Parses through a string that contains multiple IP Addresses.
#Separates each address and adds them to an ArrayList.
function ParseThroughAddresses{
    PARAM([String]$providedString)

    $providedString = $providedString.Trim()

    IF(($providedString.IndexOf(' ')) -gt -1){

        $firstSpace = ($providedString.IndexOf(' '))
        $addressBeingAdded = $providedString.Substring(0,$firstSpace)
        $theRemaining = ($providedString.Substring($firstSpace+1)).Trim()

        $possibleAddresses.Add($addressBeingAdded) | Out-Null
        ParseThroughAddresses -providedString $theRemaining
    }ELSE{
        $possibleAddresses.Add($providedString) | Out-Null
    }
}

function ParseThroughFlagIDs{
    PARAM([String]$providedString)

    $providedString = $providedString.Trim()

    IF(($providedString.IndexOf(' ')) -gt 1){

        $firstSpace = ($providedString.IndexOf(' '))
        $idBeingAdded = $providedString.Substring(0,$firstSpace)
        $theRemaining = ($providedString.Substring($firstSpace+1)).Trim()

        $oldFlagIDsArrayList.Add($idBeingAdded) | Out-Null
        ParseThroughFlagIDs -providedString $theRemaining
    }ELSE{
        $oldFlagIDsArrayList.Add($providedString) | Out-Null
    }
}

#To enter a value in sql that contains a quote, we must add another quote. So we must check and add quotes where needed.
function CheckForQuoteInPath{
    PARAM([String]$providedString)

    $providedString = $providedString.Trim()

	IF($providedString.Contains("'")){
		$indexOfApos = $providedString.IndexOf("'")

        $firstHalf = $providedString.Substring(0,$indexOfApos)
        $secondHalf = $providedString.Substring($indexOfApos)
		
        #<HashTable>.<Key> will return the corresponding value. (The value is the BadFile object, which is why we can call the setter method.)
		($flaggedFiles.$keyName).SetFilePath("$firstHalf'$secondHalf")
	}ELSE{
        #<HashTable>.<Key> will return the corresponding value. (The value is the BadFile object, which is why we can call the setter method.)
        ($flaggedFiles.$keyName).SetFilePath($providedString)
    }
}

#To enter a value in sql that contains a quote, we must add another quote. So we must check and add quotes where needed.
function CheckForQuoteInName{
    PARAM([String]$providedString)

    $providedString = $providedString.Trim()

	IF($providedString.Contains("'")){
		$indexOfApos = $providedString.IndexOf("'")

        $firstHalf = $providedString.Substring(0,$indexOfApos)
        $secondHalf = $providedString.Substring($indexOfApos)

		#<HashTable>.<Key> will return the corresponding value. (The value is the BadFile object, which is why we can call the setter method.)
		($flaggedFiles.$keyName).SetFileName("$firstHalf'$secondHalf")
	}ELSE{
        #<HashTable>.<Key> will return the corresponding value. (The value is the BadFile object, which is why we can call the setter method.)
        ($flaggedFiles.$keyName).SetFileName($providedString)
    }
}

#After converting to a string and running the substring method, the file size might not be formatted correctly.
function CheckFileSize{
    PARAM([String]$providedSize, [String]$providedUnit)

    $tempSize = ''

    TRY{
        $tempSize = ((($providedSize/"1$providedUnit")|Out-String).SubString(0,4)).Trim()
    }CATCH{
        $tempSize = ((($providedSize/"1$providedUnit")|Out-String).SubString(0,3)).Trim()
    }

    IF(($tempSize.Substring(($tempSize.Length-1))).Contains('.')){
        $tempSize = $tempSize.Substring(0, ($tempSize.Length-1))
    }

    #<HashTable>.<Key> will return the corresponding value. (The value is the BadFile object, which is why we can call the setter method.)
    ($flaggedFiles.$keyName).SetFileSize("$tempSize $providedUnit")
}

ParseThroughFlagIDs -providedString $oldFlagIDs

$maxExt = (Invoke-Sqlcmd -Query "SELECT MAX([ExtID]) FROM Extensions" -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*' | Format-Wide | Out-String).Trim()
#SQL identity columns have their first row start at ONE NOT ZERO.
$currentExt = 1
While($currentExt -le $maxExt){
    $theExt = (Invoke-Sqlcmd -Query "SELECT [ExtName] FROM Extensions WHERE [ExtID] = '$currentExt'" -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*' | Format-Wide | Out-String).Trim()
    #If an extension is removed, a newly added extension WILL NOT take it's extension ID. So, if we retrieve a blank extension that extension has been removed.
    IF($theExt -ne ''){
        $searchTerms.Add($theExt) | Out-Null
    }
    $currentExt++
}

#Places each item in the array to an ArrayList.
#ArrayLists are more maliable, therefore preferable. 
ForEach($file in $filesScannedArray){
    $filesScannedArrayList.Add($file) | Out-Null
}

#Calls the FindTheBadChildren function for each item in the $searchTerms ArrayList.
#Check the comment above the FindTheBadChildren declaration for more information.
ForEach($searchExt in $searchTerms){
    FindTheBadChildren -searchTerm $searchExt -children $filesScannedArrayList
}

#Declaration for a hash table.
$flaggedFiles = @{}
#Used in creating the key names of the $flaggedFiles hash table.
$loopNumber = 0
#Parses through each file to determine different attributes. (computerName, userName, ipAddress, filePath, fileName, fileSize.)
ForEach($theFile in $badFiles){
	#Preliminary steps for creating the key of the hash table entry
    $itemPrefix = 'FlaggedItem-'
    $keyName = $itemPrefix+$loopNumber
	#Creates a hash table entry. (Each iteration gets an entry.)
    $flaggedFiles.Add($keyName, [BadFile]::new())
    
	#Since the script is designed to handle one computer at a time, the computer name will be the same for every flagged/bad file.
	#<HashTable>.<Key> will return the corresponding value. (The value is the BadFile object, which is why we can call the setter method.)
    ($flaggedFiles.$keyName).SetComputerName($theDesiredComputer)

	#Format of $theFile --> <userName>\Dir\Dir\Dir\etc. . .
	#We want everything before the slash to find the username
    $firstSlash = $theFile.IndexOf('\')
	#<HashTable>.<Key> will return the corresponding value. (The value is the BadFile object, which is why we can call the setter method.)
	($flaggedFiles.$keyName).SetUserName(($theFile.Substring(0, $firstSlash)))
	
	#A fully qualified domain name is required for the Reolve-DnsName cmdlet.
    $fQDN = $theDesiredComputer+'.armormetal.com'
	#The Format-Wide and Out-String cmdlets format the output of piped cmdlets to be a single string containing 88 spaces between usefull characters.
    ParseThroughAddresses -providedString ((Resolve-DnsName -Name $fQDN | Format-Wide -Property IPAddress | Out-String))
    #Checks which address should be used as the computer's IP Address. (Remember: A computer can be connected to multiple networks at once.)
	ForEach($currentAddr in $possibleAddresses){
		#Addresses apart of a wired network connection take precedence over wireless addresses.
        IF(($currentAddr -LIKE '10.1.1*') -OR ($currentAddr -LIKE '10.1.5.*')){
			#<HashTable>.<Key> will return the corresponding value. (The value is the BadFile object, which is why we can call the setter method.)
			($flaggedFiles.$keyName).SetIpAddress($currentAddr)
			#A break is needed to prevent a less important address that is later in the ArrayList to be assigned as the computer's IP Address.
            BREAK
        #10.1.200.0, 10.1.210.0, 10.1.215.0, 10.1.220.0 are the network addresses for the wireless networks.
        }ELSEIF($currentAddr -LIKE '10.1.2*'){
			#<HashTable>.<Key> will return the corresponding value. (The value is the BadFile object, which is why we can call the setter method.)
            ($flaggedFiles.$keyName).SetIpAddress($currentAddr)
        }
    }

    #A file path consists of every before the final backslash. (Including the backslash itself.)
    $lastSlash = $theFile.LastIndexOf('\')
    $absoluteFilePath = 'C:\Users\' + $theFile.Substring(0, $lastSlash+1)
    
    #A file's path and name can contain a quote.
	CheckForQuoteInPath -providedString $absoluteFilePath
    CheckForQuoteInName -providedString $theFile.Substring($lastSlash+1)
    
    #I declare the variable with a empty string to prevent PowerShell from printing the value of the variable
	#Holds the size of a file as a string.
    $displayedSize = ''
	#Must convert the parsed file and computer name to a syntax understandable by the Get-Item cmdlet.
    $completeFileName = $absoluteFilePath+$theFile.Substring($lastSlash+1)
	#The Length function returns a file's size when called on a file.
    $sizeInBytes = ((Get-Item -Path $completeFileName).Length)
	#Series of nested IF Statements to determine if the base unit (BYTE) can be converted into larger units.
    IF($sizeInBytes -ge 1000){
        IF(($sizeInBytes/1kb) -ge 1000){
            IF(($sizeInBytes/1mb) -ge 1000){
                IF(($sizeInBytes/1gb) -ge 1000){
                    CheckFileSize -providedSize $sizeInBytes -providedUnit 'TB'
                }ELSE{
                    CheckFileSize -providedSize $sizeInBytes -providedUnit 'GB'
                }
            }ELSE{
                CheckFileSize -providedSize $sizeInBytes -providedUnit 'MB'
            }
        }ELSE{
            CheckFileSize -providedSize $sizeInBytes -providedUnit 'KB'
        }
    }ELSE{
        #<HashTable>.<Key> will return the corresponding value. (The value is the BadFile object, which is why we can call the setter method.)
        ($flaggedFiles.$keyName).SetFileSize("$sizeInBytes B")
    }

    #Created to make the sql query easier to read
    $sqlComputerName = ($flaggedFiles.$keyName).GetComputerName()
    $sqlUserName = ($flaggedFiles.$keyName).GetUserName()
    $sqlIpAddress = ($flaggedFiles.$keyName).GetIpAddress()
    $sqlFilePath = ($flaggedFiles.$keyName).GetFilePath()
    $sqlFileName = ($flaggedFiles.$keyName).GetFileName()
    $sqlFileSize = ($flaggedFiles.$keyName).GetFileSize()
    
    #Do not want files from the Public user
    IF($sqlUserName -ne 'Public'){

        #If the ComputerName retrieved by the query is ANYTHING BUT EMPTY then this flagged item already has an entry in the table. (A blank ComputerName means the entry does not exist in the table. --> Run the insert statement.)
        #Any column would have worked for the SELECT statement, just make sure to adjust the IF test to reflect the changes.
        $foundComputerName = ((Invoke-Sqlcmd -Query "SELECT [ComputerName] FROM FlaggedItems WHERE ComputerName = '$sqlComputerName' AND UserName = '$sqlUserName' AND FilePath = '$sqlFilePath' AND FileName = '$sqlFileName'" -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*') | Format-Wide | Out-String).Trim()
        IF(($sqlComputerName -eq $foundComputerName)){
            
            #I understand that when automated, no one will be able to read the console output. This was added in case a technician/admin wanted to manually run the script while troubleshooting.
            Write-Host 'Found a Pre-Existing Entry that Matches this Flagged Item. --> Not Adding to Database.'
            Write-Host "Computer Name: $sqlComputerName"
            Write-Host "Username: $sqlUserName"
            Write-Host "IP Address: $sqlIpAddress"
            Write-Host "File Path: $sqlFilePath"
            Write-Host "File Name: $sqlFileName"
            Write-Host "File Size: $sqlFileSize"
            Write-Host '--------------------------------------------------------------------------------------'

            
            $oldFlagIDsArrayList.Remove(((Invoke-Sqlcmd -Query "SELECT [FlagID] FROM FlaggedItems WHERE ComputerName = '$sqlComputerName' AND UserName = '$sqlUserName' AND FilePath = '$sqlFilePath' AND FileName = '$sqlFileName'" -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*') | Format-Wide | Out-String).Trim()) | Out-Null       
        }ELSE{
            $theQuery = "INSERT INTO FlaggedItems VALUES('$sqlComputerName', '$sqlUserName', '$sqlIpAddress', '$sqlFilePath', '$sqlFileName', '$sqlFileSize')";
            Invoke-Sqlcmd -Query $theQuery -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*'
           
            #I understand that when automated, no one will be able to read the console output. This was added in case a technician/admin wanted to manually run the script while troubleshooting.
            Write-Host 'New Entry Added to the Database.'
            Write-Host "Computer Name: $sqlComputerName"
            Write-Host "Username: $sqlUserName"
            Write-Host "IP Address: $sqlIpAddress"
            Write-Host "File Path: $sqlFilePath"
            Write-Host "File Name: $sqlFileName"
            Write-Host "File Size: $sqlFileSize"
            Write-Host '--------------------------------------------------------------------------------------'

            #No FlagIDs have to be removed because these new entries are added after the collection of the FlagIDs in the table is created. (The $oldFlagIDsArrayList wouldn't contain the FlagID that corresponds to this new entry.)
        }
    }
    $loopNumber++
}


ForEach($flagID IN $oldFlagIDsArrayList){
    #These string must be declared before the actual deletion because otherwise the variables would hold empty strings. 
    $removedComputerName = ((Invoke-Sqlcmd -Query "SELECT [ComputerName] FROM FlaggedItems WHERE FlagID = '$flagID'" -ServerInstance 'ARM-SQL-4.armormetal.com' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*') | Format-Wide | Out-String).Trim()
    $removedUserName = ((Invoke-Sqlcmd -Query "SELECT [UserName] FROM FlaggedItems WHERE FlagID = '$flagID'" -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*') | Format-Wide | Out-String).Trim()
    $removedIpAddress = ((Invoke-Sqlcmd -Query "SELECT [IPAddress] FROM FlaggedItems WHERE FlagID = '$flagID'" -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*') | Format-Wide | Out-String).Trim()
    $removedFilePath = ((Invoke-Sqlcmd -Query "SELECT [FilePath] FROM FlaggedItems WHERE FlagID = '$flagID'" -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*') | Format-Wide | Out-String).Trim()
    $removedFileName = ((Invoke-Sqlcmd -Query "SELECT [FileName] FROM FlaggedItems WHERE FlagID = '$flagID'" -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*') | Format-Wide | Out-String).Trim()
    $removedFileSize = ((Invoke-Sqlcmd -Query "SELECT [FileSize] FROM FlaggedItems WHERE FlagID = '$flagID'" -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*') | Format-Wide | Out-String).Trim()

    (Invoke-Sqlcmd -Query "DELETE FROM FlaggedItems WHERE FlagID = '$flagID'"  -ServerInstance '*REDACTED*' -Database '*REDACTED*' -Username '*REDACTED*' -Password '*REDACTED*') | Out-Null

    #I understand that when automated, no one will be able to read the console output. This was added in case a technician/admin wanted to manually run the script while troubleshooting.
    Write-Host 'An Entry was removed from the Database.'
    Write-Host "Computer Name: $removedComputerName"
    Write-Host "Username: $removedUserName"
    Write-Host "IP Address: $removedIpAddress"
    Write-Host "File Path: $removedFilePath"
    Write-Host "File Name: $removedFileName"
    Write-Host "File Size: $removedFileSize"
    Write-Host '--------------------------------------------------------------------------------------'
 }
