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

# Convert the XML result for functions into LSP standard library symbols format
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
    $requiredParamCount = [int]$_.functions_x005B_RequiredParameters_x005D_
    
    # Create function parameters array
    # RequiredParameters indicates how many of the first N parameters are required (rest are optional)
    $functionParameters = @()
    if ($params) {
        # Check if params is an object/hashtable (with key-value pairs)
        if ($params -is [PSCustomObject] -or $params -is [System.Collections.Hashtable]) {
            $paramNames = $params.PSObject.Properties.Name
            $paramIndex = 0
            foreach ($paramName in $paramNames) {
                $isRequired = $paramIndex -lt $requiredParamCount
                $functionParameters += [pscustomobject]@{
                    name = $paramName
                    type = $params.$paramName
                    isRequired = $isRequired
                    isNullable = -not $isRequired
                    description = "Parameter $paramName"
                }
                $paramIndex++
            }
        }
        # Fallback: Check if params is an array (legacy format)
        elseif ($params -is [array] -and $params.Count -gt 0) {
            for ($i = 0; $i -lt $params.Count; $i++) {
                $isRequired = $i -lt $requiredParamCount
                $functionParameters += [pscustomobject]@{
                    name = $params[$i]
                    type = "any"
                    isRequired = $isRequired
                    isNullable = -not $isRequired
                    description = "Parameter $($params[$i])"
                }
            }
        }
    }
    
    $functions += [pscustomobject]@{
        name = $functionName
        type = "function"
        isDataSource = ($doc.'Documentation.Category' -eq "Accessing data")
        documentation = [pscustomobject]@{
            description = $doc.'Documentation.Description'
            longDescription = $doc.'Documentation.LongDescription'
            category = $doc.'Documentation.Category'
        }
        functionParameters = $functionParameters
        returnType = if ($returnType) { $returnType.ToLower() } else { "any" }
    }
}

# Invoke Analysis Services Command to get types data
[xml]$daxTypesResult = Invoke-ASCmd -Server $asInstance -Query "EVALUATE types"

# Convert the XML result for types into LSP format
$types = @()
$daxTypesResult.return.root.row | ForEach-Object {
    $typeName = $_.types_x005B_FullType_x005D_
    
    # Skip null or empty type names
    if ([string]::IsNullOrEmpty($typeName)) {
        return
    }
    
    $doc = $_.types_x005B_Documentation_x005D_ | ConvertFrom-Json
    
    $types += [pscustomobject]@{
        name = $typeName
        type = "type"
        baseType = $_.types_x005B_Type_x005D_
        documentation = [pscustomobject]@{
            description = $doc.'Documentation.Description'
            longDescription = $doc.'Documentation.LongDescription'
            category = $doc.'Documentation.Category'
        }
        allowedValues = $doc.'Documentation.AllowedValues'
    }
}

# Invoke Analysis Services Command to get enumeration options/values data
[xml]$daxEnumerationOptionsResult = Invoke-ASCmd -Server $asInstance -Query "EVALUATE enum_options"

# Convert the XML result for enumeration options into LSP format
$enums = @()
$enumGroups = @{}

$daxEnumerationOptionsResult.return.root.row | ForEach-Object {
    $enumName = $_.enum_options_x005B_Enum_x005D_
    $option = $_.enum_options_x005B_Option_x005D_
    $fullOption = $_.enum_options_x005B_FullOption_x005D_
    $value = $_.enum_options_x005B_Value_x005D_
    
    # Group enum options by enum name
    if (-not $enumGroups.ContainsKey($enumName)) {
        $enumGroups[$enumName] = @()
    }
    
    if ($option) {
        $enumGroups[$enumName] += [pscustomobject]@{
            name = $option
            fullName = $fullOption
            value = $value
        }
    }
}

# Convert grouped enums to final format
foreach ($enumName in $enumGroups.Keys) {
    $enums += [pscustomobject]@{
        name = $enumName
        type = "enum"
        options = $enumGroups[$enumName]
    }
}

# Create the output array in standard library symbols format
# Combine all symbols into a single array like the standard-library-symbols-en-us.json
$allSymbols = @()
$allSymbols += $functions
$allSymbols += $types  
$allSymbols += $enums

# Convert the combined symbols array to JSON and save to the specified file
$allSymbols | ConvertTo-Json -Depth 10 | Out-File $outputJsonFile

Write-Host "📊 Generated $($functions.Count) functions, $($types.Count) types, and $($enums.Count) enums"
Write-Host "💾 Total symbols: $($allSymbols.Count)"
