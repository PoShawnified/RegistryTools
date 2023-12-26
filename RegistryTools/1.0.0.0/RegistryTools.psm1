#=================================================================
#region - Description
#=================================================================
# Tools to work with Microsoft Installer files
#=================================================================
#endregion - Description
#=================================================================

#=================================================================
#region - Define and Export Module Variables / Functions
#=================================================================
#=================================================================
#endregion - Define and Export Module Variables
#=================================================================

#=================================================================
#region - Define PUBLIC Advanced Functions
#=================================================================
function Read-RegFileToPSObject {
    <#
    .SYNOPSIS
        Reads a .REG file and converts the data to PS object
    .DESCRIPTION
        Parses .REG file and converts to PS Object using various schemes
    .EXAMPLE
        Read-RegFileToPSObject -LiteralPath c:\TEMP\RegistryContentFile.reg
        
        LiteralPath                                Value            Data     Type
        -----------                                -----            ----     ----
        HKEY_CURRENT_CONFIG                                                  Key
        HKEY_CURRENT_CONFIG\Software                                         Key
        HKEY_CURRENT_CONFIG\Software\Fonts                                   Key
        HKEY_CURRENT_CONFIG\Software\Fonts         LogPixels        00000060 dword
    .INPUTS
        String
    .OUTPUTS
        System.Object
    #>

    [CmdletBinding()]
    [OutputType([pscustomobject])]

    param (
        $LiteralPath
    )

    $RegFileContent = Get-Content -LiteralPath $LiteralPath

    if ($RegFileContent[0] -notmatch "^Windows Registry Editor Version 5.00$"){
        Write-Error -Message "'$LiteralPath' is not a valid registry file." -ErrorAction Stop
    }

    for ($i = 1; $i -lt $RegFileContent.Count; $i++){
        $ThisLine = $RegFileContent[$i].Trim()
    
        # Discard blank lines
        if ($ThisLine -match "^(\s*|)$"){
            $script:Path = $null
            $Value = $null
            $Data = $null
            $Type = $null

            continue
        }

        # Key
        if ($ThisLine -match "^[HKEY_[\s*,\w,\d,/,`,=,~,!,@,%,&,_,;,:,"",\,,\-,\#,\$,\^,\*,\(,\),\+,\[,\],\\,\{,\},\|,\.,\<,\>,\?]{1,}\]$"){
            # Run sub to build keypath if necessary
            #"Write key: $ThisLine"
            $script:Path = $ThisLine.TrimStart('[').TrimEnd(']')
            $Value = ''
            $Data = ''
            $Type = 'Key'
        }
        # Default Key
        elseif ($ThisLine -match "^@"){
            #"Write Default: $ThisLine"
            $Type = 'Default'
            $Value = '@'
            $Data = $ThisLine.TrimStart('@=').Trim('"')
        }
        # Value / Data
        elseif ($ThisLine -match '^"') {
    
            # Split the Value and Data e.g. "3dword"=dword:cdc69d17
            $Value,$Data = $ThisLine.Split('=', 2)
            $Value = $Value.Trim('"')

            # If the data does NOT start with a quote, we know we've got a non-string data type
            if ($Data -match '^"'){
                $Type = 'String'
                # Remove '\' escapes
                $Data = [regex]::Unescape($Data.Trim('"'))
            }
            else {
                # Split the type from the data e.g. hex(b):99,25,1c
                $Type,$Data = $Data.Split(':', 2)
                $Type = $Type.Trim('"')
                $Data = $Data.TrimEnd('\')
            }

            # Catch multiline Value / Data
            if (($Type -ne 'String') -and ($RegFileContent[$i+1] -match "^\s\s[0-9,A-F,\,]{1,}(\\|)$")){
                #$Value,$Data = $ThisLine
                #"Write data: $ThisLine"
                While ($RegFileContent[$i+1] -match "^\s\s[0-9,A-F,\,]{1,}(\\|)$"){
                    #"More data: $($RegFileContent[$i+1])"
                    $Data += ($RegFileContent[$i+1]).TrimStart(' ').TrimEnd('\')
                    $i++
                }
            }
        }

        # Create object
        [pscustomobject]@{
            LiteralPath  = $script:Path
            Value        = $Value
            Data         = $Data
            Type         = $Type
        }
    }
}

function Write-RegistryEntry {
    <#
    .SYNOPSIS
        Writes PS Object registry data to the Windows Registry
    .DESCRIPTION
        Takes data, in the output format of the Read-RegFileToPSObject cmdlet, and writes it to the registry. 
        (Optionally, verbose for each item written)
    .EXAMPLE
        Write-RegistryEntry -RegistryData (Read-RegFileToPSObject -LiteralPath c:\TEMP\RegistryContentFile.reg)
    .INPUTS
        PSObject
    #>
    [CmdletBinding()]
    
    param (
        # Validate Script
        # $RegData.Type | ForEach-Object {$_ -in ('Key','String','hex','dword','hex(b)','hex(7)','hex(2)','Default')} | Select-Object -Uinque
        [pscustomobject]$RegistryData
    )

    $RegistryData | ForEach-Object {
        # We're either a 'Key' or we're a 'Data'/'Value' pair
        if ($_.Type -eq 'Key'){
            try {
                $NewPathItem = New-Item -Path "Microsoft.PowerShell.Core\Registry::$($_.LiteralPath.Replace('/',"$([char]0x2215)"))" -Force -ErrorAction Stop #-WhatIf
                Write-Verbose -Message "Created Key: $NewPathItem"
            }
            catch {
                Write-Error $_
            }
        }
        else {
            <#
                Adjust the "Type" into a format usable by the PS Reg provider

                "Microsoft.Win32.RegistryValueKind". The possible enumeration values are: 
                    "String, ExpandString, Binary, DWord, MultiString, QWord, Unknown".

                    + "<String value data with escape characters>"
                    + hex:<Binary data (as comma-delimited list of hexadecimal values)>
                    + dword:<DWORD value integer>
                    + hex(0):<REG_NONE (as comma-delimited list of hexadecimal values)>
                    hex(10)
                    hex(100000): Binary Value... Can only find one example in use and it's a "(zero-length binary value)"
                    (Unused?) hex(1):<REG_SZ (as comma-delimited list of hexadecimal values representing a UTF-16LE NULL-terminated string)>
                    + hex(2):<Expandable string value data (as comma-delimited list of hexadecimal values representing a UTF-16LE NULL-terminated string)>
                    (Unused?) hex(3):<Binary data (as comma-delimited list of hexadecimal values)> ; equal to "Value B"
                    hex(4):<DWORD value (as comma-delimited list of 4 hexadecimal values, in little endian byte order)>
                    (Unused?) hex(5):<DWORD value (as comma-delimited list of 4 hexadecimal values, in big endian byte order)>
                    + hex(7):<Multi-string value data (as comma-delimited list of hexadecimal values representing UTF-16LE NULL-terminated strings)>
                    hex(8):<REG_RESOURCE_LIST (as comma-delimited list of hexadecimal values)>
                    hex(a):<REG_RESOURCE_REQUIREMENTS_LIST (as comma-delimited list of hexadecimal values)>
                    + hex(b):<QWORD value (as comma-delimited list of 8 hexadecimal values, in little endian byte order)>
            #>
            switch ($_){
                # hex:<Binary data (as comma-delimited list of hexadecimal values)>
                {$_.Type -eq 'String'}{
                    $Type = 'String'
                    $Data = $_.Data
                }

                # hex:<Binary data (as comma-delimited list of hexadecimal values)>
                {$_.Type -eq 'hex'   }{
                    $Type = 'Binary'
                    $Data =  $_.Data -split ',' | ForEach-Object {[System.Convert]::ToByte($_, 16)}
                }

                # dword:<DWORD value integer>
                {$_.Type -eq 'dword' }{
                    $Type = 'Dword'
                    $Data = [uint32]::Parse($_.Data, [System.Globalization.NumberStyles]::HexNumber)
                }

                # hex(0):<REG_NONE (as comma-delimited list of hexadecimal values)>
                {$_.Type -eq 'hex(0)'}{
                    $Type = ''
                    $Data = $_.Data
                }

                # hex(1):<REG_SZ (as comma-delimited list of hexadecimal values representing a UTF-16LE NULL-terminated string)>
                {$_.Type -eq 'hex(1)'}{
                    $Type = ''
                    $Data = $_.Data
                }

                # hex(10):UNKNOWN
                {$_.Type -eq 'hex(10)'}{
                    Write-Warning -Message "hex(10) is an unknown/unsupported data type at this time)"
                    return
                }

                # hex(10):UNKNOWN
                {$_.Type -eq 'hex(100000)'}{
                    Write-Warning -Message "hex(100000) is an unknown/unsupported data type at this time)"
                    return
                }

                # hex(2):<Expandable string value data (as comma-delimited list of hexadecimal values representing a UTF-16LE NULL-terminated string)>
                {$_.Type -eq 'hex(2)'}{
                    $Type = 'ExpandString'
                    $Data = HexDataToString -HexData $_.Data
                }

                # hex(3):<Binary data (as comma-delimited list of hexadecimal values)> ; equal to "Value B"
                {$_.Type -eq 'hex(3)'}{
                    Write-Warning -Message "hex(3) is an unknown/unsupported  data type at this time)"
                    return
                }

                # hex(4):<DWORD value (as comma-delimited list of 4 hexadecimal values, in little endian byte order)>
                {$_.Type -eq 'hex(4)'}{
                    $Type = 'DWORD'
                    $Data = [uint32]::Parse($_.Data, [System.Globalization.NumberStyles]::HexNumber)
                }

                # hex(5):<DWORD value (as comma-delimited list of 4 hexadecimal values, in big endian byte order)>
                {$_.Type -eq 'hex(5)'}{
                    $Type = 'DWORD'
                    $Data = [uint32]::Parse($_.Data, [System.Globalization.NumberStyles]::HexNumber)
                }

                # hex(7):<Multi-string value data (as comma-delimited list of hexadecimal values representing UTF-16LE NULL-terminated strings)>
                {$_.Type -eq 'hex(7)'}{
                    $Type = 'MultiString'
                    $Data = HexDataToString -HexData $_.Data
                }

                # hex(a):<REG_RESOURCE_REQUIREMENTS_LIST (as comma-delimited list of hexadecimal values)>
                {$_.Type -eq 'hex(a)'}{
                    $Type = ''
                    $Data = $_.Data
                }

                # hex(b):<QWORD value (as comma-delimited list of 8 hexadecimal values, in little endian byte order)>
                {$_.Type -eq 'hex(b)'}{
                    $Type = 'Qword'
                    $Data = $_.Data

                    # Reverse the Array
                    $Data = $Data.Split(',')
                    [array]::Reverse($Data, 0, $Data.Length)
                    $Data = $Data -join ''

                    # Convert to UINT64
                    $Data = [uint64]::Parse($Data, [System.Globalization.NumberStyles]::HexNumber)
                }
            }#EndSwitch

            # Write the Data / Value pair to the registry
            try {
               Write-Verbose -Message "Creating Key: $NewRegItem"
                $Splat = @{
                    LiteralPath = "Microsoft.PowerShell.Core\Registry::$($_.LiteralPath.Replace('/',"$([char]0x2215)"))"
                    Name         = $_.Value
                    Value        = $Data
                    PropertyType = $Type
                    Force        = $true
                    ErrorAction  = 'Stop'
                }

                $NewRegItem = New-ItemProperty @Splat #-WhatIf
            }
            catch {
                Write-Error $_
            }
        }
    }
}

function Get-BackgroundActivityMonitorEntries { 
    Get-ChildItem registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\bam\State\UserSettings | ForEach-Object {
        $keySID = get-item $_.PsPath
        
        $keySID.GetValueNames() | ForEach-Object {

            # We'll omit the 'Version' and 'SequenceNumber' keys. 
            if ($keySID.GetValueKind($_) -eq 'Binary'){

                [pscustomobject]@{
                    UserId        = $(try{(New-Object System.Security.Principal.SecurityIdentifier($keySID.PSChildName) -ErrorAction Stop).Translate([System.Security.Principal.NTAccount])}catch{$null})
                    SID           = $keySID.PSChildName
                    Path          = $_ 
                    ExecutionTime = [DateTime]::FromFileTime(([System.BitConverter]::ToInt64($keySID.GetValue($_), 0)))
                }
            }
        }
    }
}
#=================================================================
#endregion - Define PUBLIC Advanced Functions
#=================================================================

#=================================================================
#region - Define PRIVATE Advanced Functions
#=================================================================
function HexDataToString {
    param (
        $HexData
    )

    $ByteArr = for ($i=0; $i -lt ($HexData.Split(',')).count; $i+=2){
        [System.Convert]::ToByte(($HexData.Split(','))[$i], 16)
    }

    [System.Text.Encoding]::ASCII.GetString($ByteArr)
}

function Convert-RegPathToLiteralPath {
    [cmdletbinding()]
    [OutputType([string[]])]

    Param (
        [Parameter(
            ValueFromPipeline,
            Mandatory=$true
        )]
        [string[]]$Path
    )
    
    process { 
        $Path | ForEach-Object { 
            # If the path comes in the HK*:\ format, we'll adjust it to a more universal standard
            if ($_ -match "^HK(CU|LM|CC|CR|U)\:\\|^HKEY_(CURRENT_USER|LOCAL_MACHINE|CURRENT_CONFIG|CLASSES_ROOT|USER)\\") {

                $thisInPath = $_
            
                switch ($thisInPath.Split('\', 2)[0]){
                    {$_ -eq 'HKCU:'} {
                        $convertedPath = $thisInPath -Replace 'HKCU:', 'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'
                    }
                    {$_ -eq 'HKLM:'} {
                        $convertedPath = $thisInPath -Replace 'HKLM:', 'Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE'
                    }
                    {$_ -eq 'HKCC:'} {
                        $convertedPath = $thisInPath -Replace 'HKCC:', 'Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_CONFIG'
                    }
                    {$_ -eq 'HKCR:'} {
                        $convertedPath = $thisInPath -Replace 'HKCR:', 'Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT'
                    }
                    {$_ -eq 'HKU:'} {
                        $convertedPath = $thisInPath -Replace 'HKU:', 'Microsoft.PowerShell.Core\Registry::HKEY_USERS'
                    }
                    {$_ -in ('HKEY_CURRENT_USER','HKEY_LOCAL_MACHINE','HKEY_CURRENT_CONFIG','HKEY_CLASSES_ROOT','HKEY_USER')} {
                        $convertedPath = "Microsoft.PowerShell.Core\Registry::$($thisInPath)"
                    }
                }
                
                Write-Verbose -Message "Converting: $thisInPath  -->  $convertedPath"
                
                [string]$convertedPath

            }
            else {
                Write-Verbose -Message "Not converting: $_"
                [string]$_
            }
        }
    }
}
#=================================================================
#endregion - Define PRIVATE Advanced Functions
#=================================================================

#=================================================================
#region - Export Modules
#=================================================================
Export-ModuleMember -Function *
#=================================================================
#endregion - Export Modules
#=================================================================