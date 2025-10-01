# ==============================================================================
# Power Query M Documentation Extractor
# ==============================================================================
# This script extracts Power Query M function documentation from Power BI Desktop
# and outputs it in the original format (matching Power BI Desktop structure).
# 
# The output includes ALL fields from Power BI Desktop including:
# - Documentation.Examples (with Description, Code, Result)
# - Documentation.Name, Description, LongDescription, Category
# - Documentation.Caption, DisplayName (when available)
# - Function Parameters and Return Types
#
# This format is designed to be consumed by downstream tools that can transform
# the data as needed (e.g., LSP language server, documentation generators).
# ==============================================================================

param(
    [int]$port = 0, # Default port value (0 = auto-detect)
    [string]$outputJsonFile = "output.json" # Default output file
)

# Import necessary modules
Import-Module SqlServer

# Function to detect Analysis Services port from running Power BI Desktop
function Get-PowerBIAnalysisServicesPort {
    Write-Host "🔍 Detecting Analysis Services port from Power BI Desktop..."
    
    try {
        # Find msmdsrv.exe processes (Analysis Services)
        $msmdsrvProcesses = Get-Process -Name msmdsrv -ErrorAction SilentlyContinue
        
        if (-not $msmdsrvProcesses) {
            Write-Host "❌ No Analysis Services (msmdsrv.exe) processes found"
            Write-Host "💡 Make sure Power BI Desktop is running with a .pbix file open"
            return $null
        }
        
        Write-Host "✅ Found $($msmdsrvProcesses.Count) Analysis Services process(es)"
        
        # Get TCP connections for each msmdsrv process
        $detectedPorts = @()
        foreach ($process in $msmdsrvProcesses) {
            try {
                $connections = Get-NetTCPConnection -OwningProcess $process.Id -State Listen -ErrorAction SilentlyContinue
                foreach ($conn in $connections) {
                    if ($conn.LocalPort -and $conn.LocalPort -ne 0) {
                        $detectedPorts += $conn.LocalPort
                        Write-Host "🔗 Found Analysis Services on port: $($conn.LocalPort) (PID: $($process.Id))"
                    }
                }
            } catch {
                Write-Host "⚠️ Could not get connections for process PID: $($process.Id)"
            }
        }
        
        if ($detectedPorts.Count -eq 0) {
            Write-Host "❌ No listening ports found for Analysis Services processes"
            return $null
        }
        
        # Return the first detected port
        $selectedPort = $detectedPorts | Select-Object -First 1
        Write-Host "✅ Selected port: $selectedPort"
        return $selectedPort
        
    } catch {
        Write-Host "❌ Error detecting Analysis Services port: $_"
        return $null
    }
}

# Determine the port to use
if ($port -eq 0) {
    # Auto-detect port
    $detectedPort = Get-PowerBIAnalysisServicesPort
    if ($detectedPort) {
        $port = $detectedPort
        Write-Host "🎯 Using auto-detected port: $port"
    } else {
        Write-Error "❌ Could not auto-detect Analysis Services port. Please specify -port parameter."
        Write-Host "💡 Example: .\generate-doc-output.ps1 -port 62862"
        exit 1
    }
} else {
    Write-Host "🎯 Using specified port: $port"
}

# Construct Analysis Services instance address
$asInstance = "localhost:$port"

# Invoke the command and expect the result to be in XML format
$xmlResult = Invoke-ASCmd -ConnectionString $asInstance -Query "SELECT [CATALOG_NAME] FROM `$SYSTEM.DBSCHEMA_CATALOGS"

# Parse the XML result
$xml = [xml]$xmlResult

# Now extract the data you need from the XML
# You need to adjust the XPath query according to the actual structure of the XML
$dbName = $xml.return.root.row.CATALOG_NAME

# Process the database
Invoke-ProcessASDatabase -DatabaseName $dbName  -RefreshType "Full" -Server $asInstance

Write-Host "Database $dbName processed."

# Invoke Analysis Services Command to get functions data
[xml]$daxFunctionsResult = Invoke-ASCmd -Server $asInstance -Query "EVALUATE functions"

# Convert the XML result for functions into original format (preserving all fields)
$functions = @()
$daxFunctionsResult.return.root.row | ForEach-Object {
    $functionName = $_.functions_x005B_Function_x005D_
    
    # Skip null or empty function names
    if ([string]::IsNullOrEmpty($functionName)) {
        return
    }
    
    $doc = $_.functions_x005B_Documentation_x005D_ | ConvertFrom-Json
    $params = $_.functions_x005B_Parameters_x005D_ | ConvertFrom-Json
    $returnType = $_.functions_x005B_ReturnType_x005D_
    $requiredParamCount = $_.functions_x005B_RequiredParameters_x005D_
    
    $functions += [pscustomobject]@{
        Name = $functionName
        Documentation = $doc
        ReturnType = $returnType
        Parameters = $params
        RequiredParameters = $requiredParamCount
    }
}

# Invoke Analysis Services Command to get types data
[xml]$daxTypesResult = Invoke-ASCmd -Server $asInstance -Query "EVALUATE types"

# Convert the XML result for types into original format (preserving all fields)
$types = @()
$daxTypesResult.return.root.row | ForEach-Object {
    $typeName = $_.types_x005B_FullType_x005D_
    
    # Skip null or empty type names
    if ([string]::IsNullOrEmpty($typeName)) {
        return
    }
    
    $doc = $_.types_x005B_Documentation_x005D_ | ConvertFrom-Json
    
    $types += [pscustomobject]@{
        FullType = $typeName
        Type = $_.types_x005B_Type_x005D_
        Documentation = $doc
    }
}

# Invoke Analysis Services Command to get enumeration options/values data
[xml]$daxEnumerationOptionsResult = Invoke-ASCmd -Server $asInstance -Query "EVALUATE enum_options"

# Convert the XML result for enumeration options into original format (preserving all fields)
$enumOptions = @()
$daxEnumerationOptionsResult.return.root.row | ForEach-Object {
    $enumName = $_.enum_options_x005B_Enum_x005D_
    $option = $_.enum_options_x005B_Option_x005D_
    $fullOption = $_.enum_options_x005B_FullOption_x005D_
    $value = $_.enum_options_x005B_Value_x005D_
    
    if ($option) {
        $enumOptions += [pscustomobject]@{
            Enum = $enumName
            Option = $option
            FullOption = $fullOption
            Value = $value
        }
    }
}

# Create the output object in original format (matching Power BI Desktop structure)
$outputData = [pscustomobject]@{
    functions = $functions
    types = $types
    enum_options = $enumOptions
}

# Convert to JSON and save to the specified file
$outputData | ConvertTo-Json -Depth 10 | Out-File $outputJsonFile

Write-Host "📊 Generated $($functions.Count) functions, $($types.Count) types, and $($enumOptions.Count) enum options"
Write-Host "💾 Output saved to: $outputJsonFile"


