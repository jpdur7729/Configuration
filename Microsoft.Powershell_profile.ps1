# ------------------------------------------------------------------------------
#                     Author    : F2 - JPD
#                     Time-stamp: "2021-01-24 11:36:48 jpdur"
# ------------------------------------------------------------------------------
# Actual Location is c:\Users\jpdur\Documents\WindowsPowerShell

$global:CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()

# read a table in an ORG file and create an xlsx spreadsheet accordingly 
# WIP -- This is just the 1st key steps than can be used 
function TestOrgTable2XLSX() {
    # 1) Caution it seems that there is a 1st extea column added in the process which is not 
    # required as per the example below ==> TBC 
    # 2) The Dest should be provided even if is it a new buffer 
    # how does that work with the interactive FFind File ????
    # to be tested 
    $Data = Import-CSV List.csv -Delimiter '|'
    Import-CSV List.csv -Delimiter '|' | Export-Excel -show List.xlsx -WorksheetName Data
}


# read an XL file, extract a table and presents it into an org compatible table format
function XLTable2String($Dest)
{
    # # Extract the Table into a csv file for Debug only
    # Import-Excel $Dest | Export-Csv Test2.csv -Delimiter '|' -NoTypeInformation 

    # Extract the Table into a String with | delimiters between the fields
    $TextTable = Import-Excel $Dest | ConvertTo-Csv -Delimiter '|' -NoTypeInformation 

    # It comes as follows "Lisa"|"Mum"|"62"
    $TextTable = $TextTable.replace('"|"','|')

    # Replace the " at the beginning and end of file 
    # !!!! .replace does not process regexp but -replace does !!!!!
    $TextTable = $TextTable -replace('^"(.*?)"$','|$1|')

    return $TextTable 
}

function ImageinClipboard2File($Dest)
{

    # Any image avaiable in the clipboard
    $img=get-clipboard -format image
    if ($img -ne $null) {

	# Create the file and the directory if required - No Output
	New-Item -Path $Dest -ItemType File -Force | out-null

	# Actually save thr image into the file 
	$img.save($Dest)
	
	return "Image found in clipboard and created in $Dest"
    }
    else {
	return "No image found in clipboard"
    }
}

function prompt
{
    $wintitle = $CurrentUser.Name + " " + $Host.Name + " " + $Host.Name
    $host.ui.rawui.WindowTitle = $wintitle
    Write-Host ("PS " + $(get-location) +">") -nonewline -foregroundcolor Magenta
    return " "
}

# Check if running an elevated prompt - Added 3/11/2020 JPD
function check-prompt-elevation {

	Write-Host "Checking for elevated permissions..."

	if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
	  [Security.Principal.WindowsBuiltInRole] "Administrator")) {
		  Write-Warning "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
		  Break
	  }
	else {
		Write-Host "Code is running as administrator go on executing the script..." -ForegroundColor Green
	}
}

# Create an alias to be able to call it cpe
Set-Alias -Name cpe -Value check-prompt-elevation
Set-Alias -Name cep -Value check-prompt-elevation

# Import some modules for sql instead of SqlServerCmdletSnapin100
# !!!! Danger -- a couple of non anticipated side effect (cf. sqlmodels)
# Import-Module "sqlps"

###[System.Reflection.Assembly]::LoadWithPartialName(“Microsoft.SqlServer.Smo”)
###[System.Reflection.Assembly]::LoadWithPartialName(“Microsoft.SqlServer.SmoExtended”)

# To get access to MSMQ Queues
###[Reflection.Assembly]::LoadWithPartialName("System.Messaging")

# List Existing Databases
# cd SQLSERVER:\sql\localhost\JPDURANDEAU
# Dir Databases | Select Name

# Test to check what is the default path when called
# C:\Users\jpdurandeau\Documents\WindowsPowerShell
# Split-Path $script:MyInvocation.MyCommand.Path | sc E:\CharlesRiver\9146\serverapps\bin\path.txt

#######################
<#
.SYNOPSIS
Runs a T-SQL script.
.DESCRIPTION
Runs a T-SQL script. Invoke-Sqlcmd2 only returns message output, such as the output of PRINT statements when -verbose parameter is specified
.INPUTS
None
    You cannot pipe objects to Invoke-Sqlcmd2
.OUTPUTS
   System.Data.DataTable
.EXAMPLE
Invoke-Sqlcmd2 -ServerInstance "MyComputer\MyInstance" -Query "SELECT login_time AS 'StartTime' FROM sysprocesses WHERE spid = 1"
This example connects to a named instance of the Database Engine on a computer and runs a basic T-SQL query.
StartTime
-----------
2010-08-12 21:21:03.593
.EXAMPLE
Invoke-Sqlcmd2 -ServerInstance "MyComputer\MyInstance" -InputFile "C:\MyFolder\tsqlscript.sql" | Out-File -filePath "C:\MyFolder\tsqlscript.rpt"
This example reads a file containing T-SQL statements, runs the file, and writes the output to another file.
.EXAMPLE
Invoke-Sqlcmd2  -ServerInstance "MyComputer\MyInstance" -Query "PRINT 'hello world'" -Verbose
This example uses the PowerShell -Verbose parameter to return the message output of the PRINT command.
VERBOSE: hello world
.NOTES
Version History
v1.0   - Chad Miller - Initial release
v1.1   - Chad Miller - Fixed Issue with connection closing
v1.2   - Chad Miller - Added inputfile, SQL auth support, connectiontimeout and output message handling. Updated help documentation
v1.3   - Chad Miller - Added As parameter to control DataSet, DataTable or array of DataRow Output type

--------- Examples
Invoke-Sqlcmd2 -ServerInstance JPDURANDEAU -Database SLIH923 -Query "Select count(*) from CSM_SECURITY"
Invoke-Sqlcmd2 -ServerInstance JPDURANDEAU -Database SLIH923 -Query "EXEC JPD_TEST3 'TOTO','TITI','TUTU'"
#>
function Invoke-Sqlcmd2
{
    [CmdletBinding()]
    param(
    [Parameter(Position=0, Mandatory=$true)] [string]$ServerInstance,
    [Parameter(Position=1, Mandatory=$false)] [string]$Database,
    [Parameter(Position=2, Mandatory=$false)] [string]$Query,
    [Parameter(Position=3, Mandatory=$false)] [string]$Username,
    [Parameter(Position=4, Mandatory=$false)] [string]$Password,
    [Parameter(Position=5, Mandatory=$false)] [Int32]$QueryTimeout=600,
    [Parameter(Position=6, Mandatory=$false)] [Int32]$ConnectionTimeout=15,
    [Parameter(Position=7, Mandatory=$false)] [ValidateScript({test-path $_})] [string]$InputFile,
    [Parameter(Position=8, Mandatory=$false)] [ValidateSet("DataSet", "DataTable", "DataRow")] [string]$As="DataRow"
    )

    if ($InputFile)
    {
        $filePath = $(resolve-path $InputFile).path
        $Query =  [System.IO.File]::ReadAllText("$filePath")
    }

    $conn=new-object System.Data.SqlClient.SQLConnection

    if ($Username)
    { $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $ServerInstance,$Database,$Username,$Password,$ConnectionTimeout }
    else
    { $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $ServerInstance,$Database,$ConnectionTimeout }

    $conn.ConnectionString=$ConnectionString

    #Following EventHandler is used for PRINT and RAISERROR T-SQL statements. Executed when -Verbose parameter specified by caller
    if ($PSBoundParameters.Verbose)
    {
        $conn.FireInfoMessageEventOnUserErrors=$true
        $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {Write-Verbose "$($_)"}
        $conn.add_InfoMessage($handler)
    }

    $conn.Open()
    $cmd=new-object system.Data.SqlClient.SqlCommand($Query,$conn)
    $cmd.CommandTimeout=$QueryTimeout
    $ds=New-Object system.Data.DataSet
    $da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd)
    [void]$da.fill($ds)
    $conn.Close()
    switch ($As)
    {
        'DataSet'   { Write-Output ($ds) }
        'DataTable' { Write-Output ($ds.Tables) }
        'DataRow'   { Write-Output ($ds.Tables[0]) }
    }

} #Invoke-Sqlcmd2

Function Get-SQLInstance {  
    <#
        .SYNOPSIS
            Retrieves SQL server information from a local or remote servers.

        .DESCRIPTION
            Retrieves SQL server information from a local or remote servers. Pulls all
            instances from a SQL server and detects if in a cluster or not.

        .PARAMETER Computername
            Local or remote systems to query for SQL information.

        .NOTES
            Name: Get-SQLInstance
            Author: Boe Prox
            DateCreated: 07 SEPT 2013

        .EXAMPLE
            Get-SQLInstance -Computername DC1

            SQLInstance   : MSSQLSERVER
            Version       : 10.0.1600.22
            isCluster     : False
            Computername  : DC1
            FullName      : DC1
            isClusterNode : False
            Edition       : Enterprise Edition
            ClusterName   :
            ClusterNodes  : {}
            Caption       : SQL Server 2008

            SQLInstance   : MINASTIRITH
            Version       : 10.0.1600.22
            isCluster     : False
            Computername  : DC1
            FullName      : DC1\MINASTIRITH
            isClusterNode : False
            Edition       : Enterprise Edition
            ClusterName   :
            ClusterNodes  : {}
            Caption       : SQL Server 2008

            Description
            -----------
            Retrieves the SQL information from DC1
    #>
    [cmdletbinding()] 
    Param (
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('__Server','DNSHostName','IPAddress')]
        [string[]]$ComputerName = $env:COMPUTERNAME
    )
    Process {
        ForEach ($Computer in $Computername) {
            $Computer = $computer -replace '(.*?)\..+','$1'
            Write-Verbose ("Checking {0}" -f $Computer)
            Try {
                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
                $baseKeys = "SOFTWARE\\Microsoft\\Microsoft SQL Server",
                "SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server"
                If ($reg.OpenSubKey($basekeys[0])) {
                    $regPath = $basekeys[0]
                } ElseIf ($reg.OpenSubKey($basekeys[1])) {
                    $regPath = $basekeys[1]
                } Else {
                    Continue
                }
                $regKey= $reg.OpenSubKey("$regPath")
                If ($regKey.GetSubKeyNames() -contains "Instance Names") {
                    $regKey= $reg.OpenSubKey("$regpath\\Instance Names\\SQL" )
                    $instances = @($regkey.GetValueNames())
                } ElseIf ($regKey.GetValueNames() -contains 'InstalledInstances') {
                    $isCluster = $False
                    $instances = $regKey.GetValue('InstalledInstances')
                } Else {
                    Continue
                }
                If ($instances.count -gt 0) {
                    ForEach ($instance in $instances) {
                        $nodes = New-Object System.Collections.Arraylist
                        $clusterName = $Null
                        $isCluster = $False
                        $instanceValue = $regKey.GetValue($instance)
                        $instanceReg = $reg.OpenSubKey("$regpath\\$instanceValue")
                        If ($instanceReg.GetSubKeyNames() -contains "Cluster") {
                            $isCluster = $True
                            $instanceRegCluster = $instanceReg.OpenSubKey('Cluster')
                            $clusterName = $instanceRegCluster.GetValue('ClusterName')
                            $clusterReg = $reg.OpenSubKey("Cluster\\Nodes")
                            $clusterReg.GetSubKeyNames() | ForEach {
                                $null = $nodes.Add($clusterReg.OpenSubKey($_).GetValue('NodeName'))
                            }
                        }
                        $instanceRegSetup = $instanceReg.OpenSubKey("Setup")
                        Try {
                            $edition = $instanceRegSetup.GetValue('Edition')
                        } Catch {
                            $edition = $Null
                        }
                        Try {
                            $ErrorActionPreference = 'Stop'
                            #Get from filename to determine version
                            $servicesReg = $reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Services")
                            $serviceKey = $servicesReg.GetSubKeyNames() | Where {
                                $_ -match "$instance"
                            } | Select -First 1
                            $service = $servicesReg.OpenSubKey($serviceKey).GetValue('ImagePath')
                            $file = $service -replace '^.*(\w:\\.*\\sqlservr.exe).*','$1'
                            $version = (Get-Item ("\\$Computer\$($file -replace ":","$")")).VersionInfo.ProductVersion
                        } Catch {
                            #Use potentially less accurate version from registry
                            $Version = $instanceRegSetup.GetValue('Version')
                        } Finally {
                            $ErrorActionPreference = 'Continue'
                        }
                        New-Object PSObject -Property @{
                            Computername = $Computer
                            SQLInstance = $instance
                            Edition = $edition
                            Version = $version
                            Caption = {Switch -Regex ($version) {
                                "^14" {'SQL Server 2014';Break}
                                "^11" {'SQL Server 2012';Break}
                                "^10\.5" {'SQL Server 2008 R2';Break}
                                "^10" {'SQL Server 2008';Break}
                                "^9"  {'SQL Server 2005';Break}
                                "^8"  {'SQL Server 2000';Break}
                                Default {'Unknown'}
                            }}.InvokeReturnAsIs()
                            isCluster = $isCluster
                            isClusterNode = ($nodes -contains $Computer)
                            ClusterName = $clusterName
                            ClusterNodes = ($nodes -ne $Computer)
                            FullName = {
                                If ($Instance -eq 'MSSQLSERVER') {
                                    $Computer
                                } Else {
                                    "$($Computer)\$($instance)"
                                }
                            }.InvokeReturnAsIs()
                        }
                    }
                }
            } Catch {
                Write-Warning ("{0}: {1}" -f $Computer,$_.Exception.Message)
            }
        }
    }
}

##################################

Set-Alias notep++ 'C:\Program Files (x86)\Notepad++\notepad++.exe'

##################################
<#
	Function to convert a xxx file in CRIMS format with ~ into a CSV file
#>
function ConvCRIMS2CSV {

	param(   [string]$Filename )

	$s = gc -Path (".\"+$Filename)
	$s = $s -replace("~",",")
	$s | sc -Path (".\"+[System.IO.Path]::GetFileNameWithoutExtension($Filename)+".csv")

}
###################################

##################################
<#
	Function to convert a .DAT file in CRIMS format
		with | as a separator
			and "" around each field
	into a CSV file
#>
function ConvDAT2CSV {

	param(   [string]$Filename )

	$s = gc -Path (".\"+$Filename)

	#Get rid of , which appear in some amounts
	$s = $s -replace(",","")

	#Replace the field separator
	$s = $s -replace("\|",",")

	#Eliminate the "
	$s = $s -replace('"','')

	$s | sc -Path (".\"+[System.IO.Path]::GetFileNameWithoutExtension($Filename)+".csv")

}
###################################

<#
.SYNOPSIS
   <A brief description of the script>
   Replacement of a stream of complex and various files/technologies
.DESCRIPTION
   <A detailed description of the script>
   Creation of the XML message after extract from Database and copy of the file to be processed by the Process Server

		<?xml version="1.0" encoding="UTF-8" ?><envelope transaction="multi" ack="details"><auth><user>tm_dev</user><password>tm_dev</password><session>session1</session></auth>
		<cashForecastExt op="create">
		<acctCd>GBPV_61</acctCd><amtType>FTC</amtType><amt>930000.0000</amt><crrncyCd>GBP</crrncyCd><descr>Internal Subscription/Redemption 10344196
		</descr><forecastDate>2014-09-22</forecastDate>
		</cashForecastExt></envelope>

.PARAMETER <paramName>
   <Description of script parameter>
.EXAMPLE
   <An example of using the script>
   ProcessLiquidity -Sec_id 12680230 -Order_id 14563918 -ServerInstance JPDURANDEAU -Database SLIH923 -ProcessServerPath w:\import\OrderImport\
#>

function ProcessLiquidity
{
param(
	[Parameter(Position=0, Mandatory=$true)]  [string]$Sec_id,
	[Parameter(Position=1, Mandatory=$true)]  [string]$Order_id,
	[Parameter(Position=2, Mandatory=$true)]  [string]$ProcessServerPath,
	[Parameter(Position=3, Mandatory=$true)]  [string]$ServerInstance,
	[Parameter(Position=4, Mandatory=$false)] [string]$Database

)

<#
# Example of a typical message to be generated
<?xml version="1.0" encoding="UTF-8" ?><envelope transaction="multi" ack="details"><auth><user>tm_dev</user><password>tm_dev</password><session>session1</session></auth>
<cashForecastExt op="create">
<acctCd>GBPV_61</acctCd><amtType>FTC</amtType><amt>930000.0000</amt><crrncyCd>GBP</crrncyCd><descr>Internal Subscription/Redemption 10344196
</descr><forecastDate>2014-09-22</forecastDate>
</cashForecastExt></envelope>
#>

# Create unique File name to create the XML message
$FileName = Get-Date -format s
$FileName = $FileName -replace ':' -replace '-'
$FileName = $ProcessServerPath + 'X' + $FileName + ".order.xml"

#Debug
#$FileName

# Prepare the SQL to extract the data from the database
$SQLString = "select '<acctCd>',unit_trust_id,'</acctCd>',"
$SQLString = $SQLString + "'<amtType>FTC</amtType>',"
$SQLString = $SQLString + "'<amt>',case trans_type when 'BUYL' then target_amt else -target_amt end,'</amt>',"
$SQLString = $SQLString + "'<crrncyCd>',s.ASSET_CRRNCY_CD,'</crrncyCd>',"
$SQLString = $SQLString + "'<descr>','Internal Subscription/Redemption '+convert(varchar(10),o.order_id),'</descr>',"
$SQLString = $SQLString + "'<forecastDate>',CONVERT(VARCHAR(24),o.SETTLE_DATE,120),'</forecastDate>'"
$SQLString = $SQLString + " FROM csm_security s,ts_order o"
#$SQLString = $SQLString + " WHERE o.sec_id = s.sec_id and s.sec_id = 12680230 and o.order_id = 14563918"
$SQLString = $SQLString + " WHERE o.sec_id = s.sec_id and s.sec_id = " + $Sec_id + " and o.order_id = " + $Order_id

# Debug Display SQL String
#$SQLString

# Extract the data from Database
#Invoke-Sqlcmd2 -ServerInstance JPDURANDEAU -Database SLIH923 -Query $SQLString | export-csv WFR1.csv
Invoke-Sqlcmd2 -ServerInstance $ServerInstance -Database $Database -Query $SQLString | export-csv WFR1.csv

# Process the result to format it accordingly
$a = cat WFR1.csv -Tail 1
# Get rid of " and comma separator
$a = $a -replace '"' -replace ','

# Prepare the wrappers for the XML Message
$XMLHeader = '<?xml version="1.0" encoding="UTF-8" ?><envelope transaction="multi" ack="details"><auth><user>tm_dev</user><password>tm_dev</password><session>session1</session></auth><cashForecastExt op="create">'
$XMLTail = "</cashForecastExt></envelope>"

#Add Beginning/End and create file on the process server directory
$a = $XMLHeader + $a + $XMLTail | sc $FileName

}

# ----------------------------------------------------------------------
# This is a common function i am using which will release excel objects
# ----------------------------------------------------------------------

function Release-Ref ($ref) {

	([System.Runtime.InteropServices.Marshal]::ReleaseComObject( [System.__ComObject]$ref) -gt 0)

	[System.GC]::Collect()

	[System.GC]::WaitForPendingFinalizers()

}


function ExtractDD_Data {

	param( [string]$Filename )

	#Create XL Object
	$objExcel = New-Object -ComObject Excel.Application
	#$objExcel.Visible = $True

	# Get Location of Files. It has to be an absolute path corresponding to where we invoke the function
	$ExcelFilesLocation = Split-Path $script:MyInvocation.MyCommand.Path
	$ExcelFilesLocation = $ExcelFilesLocation + "\"

	# Alternative method to get the path
	#get-location | sc list
	#$text  = (cat list -Head 2)
	#$text

	# Open the excel file and get the nb of worksheets
	$UserWorkBook = $objExcel.Workbooks.Open($ExcelFilesLocation + $Filename)
	$nbWorkSheets = $UserWorkBook.Worksheets.Count

	# The workbook is already saved so no need to ask the question when quitting
	# If not popup message requestion to correct it accordingly
	$UserWorkBook.Saved = $True

	# Constant Rows where to find the information for _61
	$FundNameRow = 2
	$FundCcyRow = 5
	$FundCOBBalanceRow = 30

	# Got through the different worksheets
	for ( $CurrentWorkSheet = 1;
			($CurrentWorkSheet -le $nbWorkSheets) ;
				$CurrentWorkSheet++ ) {

		# Select the next workSheet
		$UserWorksheet = $UserWorkBook.sheets.Item($CurrentWorkSheet)

		# Get the data out of the XL spreadsheet in a CSV like format
		for ( $FundCol = 2 ;($UserWorksheet.Cells.Item($FundNameRow,$FundCol).Value() -ne $null) ; $FundCol++ ) {
			[System.IO.Path]::GetFileNameWithoutExtension($Filename)+","+
			$UserWorksheet.Cells.Item($FundNameRow,$FundCol).text.trim()+"_61"+","+
			$UserWorksheet.Cells.Item($FundCcyRow,$FundCol).text+","+
			$UserWorksheet.Cells.Item($FundCOBBalanceRow,$FundCol).value()
		}

	}

	# Exiting the excel object
	$UserWorkBook.Close()
	$objExcel.Quit()
	#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)

	#Release all the objects used above
	$a = Release-Ref($UserWorksheet)
	$a = Release-Ref($UserWorkBook)
	$a = Release-Ref($objExcel)

}

# Process/split the MEDM_MM_CASH cash file

function SplitCashByAcctType {

	#Extract the data if the .zip file exists
	If (Test-Path .\MEDM_MM_CASH.zip){
		#Cleanup if the data exist but no error message
		rm MEDM_MM_CASH.CASH_FORECAST -ErrorAction SilentlyContinue

		#Extract the data from the zip file
		C:\"Program Files"\7-Zip\7z e MEDM_MM_CASH.zip
	}

	# 1st convert the file obtained into a standard csv file
	ConvCRIMS2CSV .\MEDM_MM_CASH.CASH_FORECAST

	#Split the MEDM_MM_CASH in 3 parts // 61 - 62 - 1
	import-csv .\MEDM_MM_CASH.csv -header ACCT_CD,CCY,DATE,AMT_TYPE,Amount,f1,Descr,f2,f3,f4,f5,f6,f7,f8,f9 | where  {$_.ACCT_CD -like "*_61"} | Export-Csv CASH_61.csv -NoTypeInformation
	import-csv .\MEDM_MM_CASH.csv -header ACCT_CD,CCY,DATE,AMT_TYPE,Amount,f1,Descr,f2,f3,f4,f5,f6,f7,f8,f9 | where  {$_.ACCT_CD -like "*_62"} | Export-Csv CASH_62.csv -NoTypeInformation
	import-csv .\MEDM_MM_CASH.csv -header ACCT_CD,CCY,DATE,AMT_TYPE,Amount,f1,Descr,f2,f3,f4,f5,f6,f7,f8,f9 | where  {$_.ACCT_CD -like "*_1"}  | Export-Csv CASH_1.csv  -NoTypeInformation

	#Split the MEDM_MM_CASH in a last parts // the records not in the 1st 3 groups
	import-csv .\MEDM_MM_CASH.csv -header ACCT_CD,CCY,DATE,AMT_TYPE,Amount,f1,Descr,f2,f3,f4,f5,f6,f7,f8,f9 | where  {$_.ACCT_CD -notlike "*_1" -and $_.ACCT_CD -notlike "*_61" -and $_.ACCT_CD -notlike "*_62"}  | Export-Csv CASH_OTHER.csv  -NoTypeInformation

}

function ExtractDD_SODData {

	param( [string]$Filename )

	#Create XL Object
	$objExcel = New-Object -ComObject Excel.Application
	#$objExcel.Visible = $True

	# Get Location of Files. It has to be an absolute path corresponding to where we invoke the function
	$ExcelFilesLocation = Split-Path $script:MyInvocation.MyCommand.Path
	$ExcelFilesLocation = $ExcelFilesLocation + "\"

	# Alternative method to get the path
	#get-location | sc list
	#$text  = (cat list -Head 2)
	#$text

	# Open the excel file and get the nb of worksheets
	$UserWorkBook = $objExcel.Workbooks.Open($ExcelFilesLocation + $Filename)
	$nbWorkSheets = $UserWorkBook.Worksheets.Count

	# The workbook is already saved so no need to ask the question when quitting
	# If not popup message requestion to correct it accordingly
	$UserWorkBook.Saved = $True

	# Constant Rows where to find the information for _61
	$FundNameRow = 2
	$FundCcyRow = 5
	$FundCOBBalanceRow = 30

	# Got through the different worksheets
	for ( $CurrentWorkSheet = 1;
			($CurrentWorkSheet -le $nbWorkSheets) ;
				$CurrentWorkSheet++ ) {

		# Select the next workSheet
		$UserWorksheet = $UserWorkBook.sheets.Item($CurrentWorkSheet)

		# Get the data out of the XL spreadsheet in a CSV like format
		for ( $FundCol = 2 ;($UserWorksheet.Cells.Item($FundNameRow,$FundCol).Value() -ne $null) ; $FundCol++ ) {
			[System.IO.Path]::GetFileNameWithoutExtension($Filename)+","+
			$UserWorksheet.Cells.Item($FundNameRow,$FundCol).text.trim()+"_61"+","+
			$UserWorksheet.Cells.Item($FundCcyRow,$FundCol).text+","+
			# Add all the movements
			$UserWorksheet.Cells.Item( 8,$FundCol).value()+","+
			$UserWorksheet.Cells.Item( 9,$FundCol).value()+","+
			$UserWorksheet.Cells.Item(10,$FundCol).value()+","+
			$UserWorksheet.Cells.Item(11,$FundCol).value()+","+
			$UserWorksheet.Cells.Item(12,$FundCol).value()+","+
			$UserWorksheet.Cells.Item(13,$FundCol).value()

		}

	}

	# Exiting the excel object
	$UserWorkBook.Close()
	$objExcel.Quit()
	#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)

	#Release all the objects used above
	$a = Release-Ref($UserWorksheet)
	$a = Release-Ref($UserWorkBook)
	$a = Release-Ref($objExcel)

}


<#
# Process Message to CRIMS Queue and execute it on the spot
function SendMSG2CRIMS {

param(
	[Parameter(Position=0, Mandatory=$true)]  [string]$msgContent,
	[Parameter(Position=1, Mandatory=$false)] [string]$fullQueueName = ".\private$\request"
)

# To get access to MSMQ Queues // Usually uploaded by the default profile
#[Reflection.Assembly]::LoadWithPartialName("System.Messaging")

##
$fullQueueName = ".\private$\" + "request"
If ([System.Messaging.MessageQueue]::Exists($fullQueueName))
    {
        Write-Host($fullQueueName + " queue already exists")
	}

#Example of fully formatted message
$msgContent = "<envelope ack=""info"" transaction=""multi""><auth><user>TM_DEV</user><password>resn0_gesl0</password><session>3</session></auth><cashForecastExt op=""create""><acctCd>03_61</acctCd><amtType>FTC</amtType><amt>12500</amt><crrncyCd>GBP</crrncyCd><descr>Cash Receive Today</descr><forecastDate>2014-06-19 00:00:00.000</forecastDate></cashForecastExt><cashForecastExt op=""create""><acctCd>03_61</acctCd><amtType>FTC</amtType><amt>503.67</amt><crrncyCd>GBP</crrncyCd><descr>Intrady ADJUSTMENT</descr><forecastDate>2014-06-19 00:00:00.000</forecastDate></cashForecastExt></envelope>";

##

# Create the message
$msg = new-object System.Messaging.Message;
$msg.Body = $msgContent;
$msg.Formatter = new-object System.Messaging.ActiveXMessageFormatter;

# Opening the queue and writing the message
$queue = new-object -TypeName System.Messaging.MessageQueue -ArgumentList $fullQueueName
$queue.Formatter = new-object System.Messaging.ActiveXMessageFormatter;
$queue.Send($msg);

}
#>

# Chocolatey profile
$ChocolateyProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
if (Test-Path($ChocolateyProfile)) {
  Import-Module "$ChocolateyProfile"
}
