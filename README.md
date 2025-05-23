# Tools to CRUD the Windows Registry
## Problem Statement
Often, when presented with a large number of registry entries to create/delete/modify, the go-to solution is to export to a `.reg` file, and use `reg.exe import <file.reg>` to import the changes, then verifying `reg.exe` exits `0`.

The problem with this approach is there's no proof the import was successful for all changes in the `.reg` file. In all but the most egregious failures, `reg.exe` will return `0`.

In the face of having to deconstruct large `.reg` files and then script each registry change action, many admins will live with this limitation not knowing whether their changes were truly successful.

## CMDLETS
### Public
#### `Read-RegFileToPSObject`
Reads a `.REG` file and converts the data to PS object, ready for immediate use or export to transport (e.g. `Export-Clixml`, `ConvertTo-Json`, etc..).

#### `Write-RegistryEntry`
Takes data, in the output format of the Read-RegFileToPSObject cmdlet, and writes it to the registry. (Optionally, verbose for each item written)

#### `Convert-RegPathToLiteralPath`
Takes a string Registry path and adjusts the prefix to create a fully qualified PowerShell path.

### Private
#### `HexDataToString`
Background data conversion for some hex data.  

## Additional raw notes/info
Starter for a struct type definition.  

Musings about data challenges with different registry types.  

```
<# 
Add-Type -TypeDefinition @"
public struct regdata {
   public string LiteralPath;
   public string Value;
   public string Data;
   public string Type;
}
"@
#>

#New-ItemProperty -Path .\ -Name $bin.Value -Value ($bin.Data.Replace('\  ', '')) -PropertyType $bin.Type -Force

#[System.Convert]::FromBase64String(($bin.Data.Replace('\  ', '').Replace(',',' ')))

# ExpandString
#New-ItemProperty -Path .\ -Name $RegData[6].Value -Value $RegData[6].Data -PropertyType $RegData[6].Type -Force

$ByteArr = ($RegData[5].Data.Split(',') | %{[System.Convert]::ToByte($_, 16)})

$ByteArr = for ($i=0; $i -lt ($RegData[5].Data.Split(',')).count; $i+=2){
    [System.Convert]::ToByte(($RegData[5].Data.Split(','))[$i], 16)
}

$DataOut = [System.Text.Encoding]::ASCII.GetString($ByteArr)

$arr =@(('99,25,1c,40,a4,08,85,00').Split(','))
[array]::Reverse($arr, 0, 8)
$arr

[uint64]::Parse('99,25,1c,40,a4,08,85,00', [System.Globalization.NumberStyles]::HexNumber)
[uint64]::Parse('008508a4401c2599', [System.Globalization.NumberStyles]::HexNumber)
[uint64]::Parse($arr.Replace(',',''), [System.Globalization.NumberStyles]::HexNumber)
```
