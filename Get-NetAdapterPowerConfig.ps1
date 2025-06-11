# MIT License
# 
# Copyright (c) 2025 noderaven
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
# 
# Requires -Version 5.1
# This script is designed to be run remotely and unattended to check
# network adapter power saving feature configurations.
# It uses the NetAdapter module, available in PowerShell 5.1 and later.
# For remote execution, ensure PowerShell Remoting is enabled on the target machine:
#   Enable-PSRemoting -Force
# And you can invoke it using:
#   Invoke-Command -ComputerName YourRemotePC -ScriptBlock { & 'C:\Path\To\YourScript.ps1' }
# Or copy the function and call it directly in a remote session:
#   Invoke-Command -ComputerName YourRemotePC -ScriptBlock {
#       # Paste the Get-NetAdapterPowerConfig function here
#       Get-NetAdapterPowerConfig
#   }


function Get-NetAdapterPowerConfig {
    # Removed 'VerboseMessage' as it's not a valid parameter for CmdletBindingAttribute.
    # Verbose messages are handled by Write-Verbose calls within the function.
    [CmdletBinding(DefaultParameterSetName='AllAdapters', SupportsShouldProcess=$true)]
    param(
        # Optional: Specify the names of specific network adapters to check.
        # If not provided, the script will check all available network adapters.
        [Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='SpecificAdapter')]
        [string[]]$Name = (Get-NetAdapter -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
    )

    # Initialize an array to store all the results.
    $results = @()

    # Define the advanced power saving features to scan for.
    # Each entry includes the user-friendly feature name and an array of common
    # DisplayName patterns that might appear for this feature in Get-NetAdapterAdvancedProperty.
    # This accounts for variations in driver implementations.
    $featuresToScan = @(
        @{ FeatureName = "Advanced EEE"; DisplayNamePatterns = @("Advanced EEE") },
        @{ FeatureName = "Energy Efficient Ethernet"; DisplayNamePatterns = @("Energy Efficient Ethernet", "EEE", "Energy Efficiency Ethernet") },
        @{ FeatureName = "Ultra Low Power Mode"; DisplayNamePatterns = @("Ultra Low Power mode") },
        @{ FeatureName = "Gigabit Lite"; DisplayNamePatterns = @("Gigabit Lite") },
        @{ FeatureName = "Green Ethernet"; DisplayNamePatterns = @("Green Ethernet") },
        # Updated DisplayNamePatterns for Large Send Offload to include "v2"
        @{ FeatureName = "Large Send Offload (IPv4)"; DisplayNamePatterns = @("Large Send Offload (IPv4)", "LSOv4", "IPv4 Large Send Offload", "Large Send Offload V2 (IPv4)", "Large Send Offload v2 (IPv4)") },
        @{ FeatureName = "Large Send Offload (IPv6)"; DisplayNamePatterns = @("Large Send Offload (IPv6)", "LSOv6", "IPv6 Large Send Offload", "Large Send Offload V2 (IPv6)", "Large Send Offload v2 (IPv6)") }
    )

    # Loop through each network adapter name provided or found.
    foreach ($adapterName in $Name) {
        if ($PSCmdlet.ShouldProcess("network adapter '$adapterName'", "Check Power Configuration")) {
            $adapter = $null
            try {
                # Attempt to retrieve the network adapter object.
                # -ErrorAction Stop ensures that any error in Get-NetAdapter is caught by the try-catch block.
                $adapter = Get-NetAdapter -Name $adapterName -ErrorAction Stop
                Write-Verbose "Successfully retrieved adapter: $($adapter.Name) ($($adapter.InterfaceDescription))"
            }
            catch {
                # If an adapter cannot be retrieved, log a warning and skip to the next one.
                Write-Warning "Could not retrieve network adapter '$adapterName'. Error: $($_.Exception.Message)"
                continue # Skip this adapter and proceed to the next one in the loop.
            }

            # --- 1. Check "Allow the computer to turn off this device to save power" ---
            try {
                # Get the power management settings for the current adapter.
                $powerManagement = Get-NetAdapterPowerManagement -Name $adapter.Name -ErrorAction Stop

                # Determine the status based on the boolean property.
                $featureStatus = if ($powerManagement.AllowComputerToTurnOffDevice) { "Enabled" } else { "Disabled" }

                # Add this feature's status to the results array as a custom object.
                $results += [PSCustomObject]@{
                    AdapterName           = $adapter.Name
                    AdapterDescription    = $adapter.InterfaceDescription
                    Feature               = "Allow the computer to turn off this device to save power"
                    Status                = $featureStatus
                    OriginalPropertyName  = "AllowComputerToTurnOffDevice" # The actual property name from the cmdlet output
                    OriginalPropertyValue = $powerManagement.AllowComputerToTurnOffDevice # The raw value
                }
                Write-Verbose "Checked 'Allow computer to turn off this device to save power' for $($adapter.Name): $($featureStatus)"
            }
            catch {
                # If power management settings cannot be retrieved, indicate an error.
                Write-Warning "Could not get 'Allow the computer to turn off this device to save power' setting for adapter '$($adapter.Name)'. Error: $($_.Exception.Message)"
                $results += [PSCustomObject]@{
                    AdapterName           = $adapter.Name
                    AdapterDescription    = $adapter.InterfaceDescription
                    Feature               = "Allow the computer to turn off this device to save power"
                    Status                = "N/A (Error retrieving)"
                    OriginalPropertyName  = "AllowComputerToTurnOffDevice"
                    OriginalPropertyValue = $null
                }
            }

            # --- 2. Check Advanced Properties (EEE, LSO, etc.) ---
            $advancedProperties = @()
            try {
                # Get all advanced properties for the current adapter.
                # Some adapters might not have advanced properties, or the cmdlet might fail.
                $advancedProperties = Get-NetAdapterAdvancedProperty -Name $adapter.Name -ErrorAction Stop
                Write-Verbose "Retrieved $($advancedProperties.Count) advanced properties for $($adapter.Name)."
            }
            catch {
                # If advanced properties cannot be retrieved, log a warning.
                Write-Warning "Could not get advanced properties for adapter '$($adapter.Name)'. Error: $($_.Exception.Message)"
                # No advanced properties means none of the features below can be checked for this adapter.
            }

            # Iterate through each defined feature we are looking for.
            foreach ($featureDef in $featuresToScan) {
                $foundProperty = $null
                $actualPropertyName = "Not Found" # Default if property isn't found
                $status = "Not Supported"        # Default if property isn't found
                $propertyValue = $null

                # Loop through potential display name patterns for the current feature.
                foreach ($pattern in $featureDef.DisplayNamePatterns) {
                    # Search for an advanced property whose DisplayName matches the pattern.
                    # Select the first match found.
                    $foundProperty = $advancedProperties | Where-Object { $_.DisplayName -like "*$pattern*" } | Select-Object -First 1
                    if ($foundProperty) {
                        break # Found a match, no need to check other patterns for this feature.
                    }
                }

                # If the advanced property was found for this feature...
                if ($foundProperty) {
                    $actualPropertyName = $foundProperty.DisplayName
                    $propertyValue = $foundProperty.RegistryValue

                    # --- Enhanced interpretation of RegistryValue for 'Enabled'/'Disabled' status ---
                    # Common scenarios:
                    # 1. Integer (e.g., 0 or 1)
                    # 2. String (e.g., "0" or "1")
                    # 3. String array with a single element (e.g., {"0"} or {"1"})
                    # 4. Other types/values (pass through)

                    $tempValue = $null
                    if ($propertyValue -is [int]) {
                        $tempValue = $propertyValue
                    }
                    elseif ($propertyValue -is [string]) {
                        [int]::TryParse($propertyValue, [ref]$tempValue) | Out-Null
                    }
                    elseif ($propertyValue -is [System.String[]] -and $propertyValue.Length -gt 0) {
                        [int]::TryParse($propertyValue[0], [ref]$tempValue) | Out-Null
                    }

                    # Determine status based on the interpreted integer value
                    if ($tempValue -ne $null -and ($tempValue -eq 0 -or $tempValue -eq 1)) {
                        $status = if ($tempValue -eq 1) { "Enabled" } else { "Disabled" }
                    }
                    else {
                        # Fallback if interpretation failed or value is not 0/1
                        if ($propertyValue -is [System.Array]) {
                            $status = "Value: '{ $($propertyValue -join ', ') }' (Type: $($propertyValue.GetType().Name))"
                        } else {
                            $status = "Value: '$propertyValue' (Type: $($propertyValue.GetType().Name))"
                        }
                    }
                }
                Write-Verbose "Checked '$($featureDef.FeatureName)' for $($adapter.Name): $($status)"

                # Add the feature's status to the results array.
                $results += [PSCustomObject]@{
                    AdapterName           = $adapter.Name
                    AdapterDescription    = $adapter.InterfaceDescription
                    Feature               = $featureDef.FeatureName
                    Status                = $status
                    OriginalPropertyName  = $actualPropertyName
                    OriginalPropertyValue = $propertyValue
                }
            }
        }
    }

    # Output the collected results, sorted for better readability.
    $results | Sort-Object AdapterName, Feature
}

# --- Script Execution ---
# When this script is run, it will execute the Get-NetAdapterPowerConfig function.
# The output will be a collection of PowerShell objects, suitable for console display
# or piping to other cmdlets like Export-Csv, ConvertTo-Json, etc.

# Example: To run and display the results in the console with better formatting:
# Added -Wrap to ensure long column content is not truncated and wraps to the next line.
Get-NetAdapterPowerConfig | Format-Table -AutoSize -Wrap

# Example: To run and save the results to a CSV file (recommended for unattended execution):
# Get-NetAdapterPowerConfig | Export-Csv -Path "C:\Temp\NetworkPowerConfig_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation -Encoding UTF8

# Example: To run with verbose output for debugging:
# Get-NetAdapterPowerConfig -Verbose

# Example: To check a specific adapter by name:
# Get-NetAdapterPowerConfig -Name "Ethernet" | Format-Table -AutoSize -Wrap
