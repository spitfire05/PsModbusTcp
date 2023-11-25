
$allReferences = @()
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

Write-Host "Loading module PS.Modbus from path $scriptPath"
foreach($file in (Get-ChildItem -Path $scriptPath -Filter "*-references.json"))
{
    Write-Host "Loading $file"
    $allReferences += Get-Content $file.fullName | ConvertFrom-Json
}

enum ModbusFunction {
    ReadCoilStatus = [byte]0x01 #0-10.000
    ReadInputStatus = [byte]0x02 #10.000-20.000
    ReadHoldingRegister = [byte]0x03 #40.000-50.000
    ReadInputRegister = [byte]0x04 #30.000-40.000
    WriteCoilStatus = [byte]0x05 #0-10.000
    WriteHoldingRegister = [byte]0x06 #40.000-50.000
    WriteMultipleCoilStatus = [byte]0xf #0-10.000
    WriteMultipleRHoldingRegisters = [byte]0x10 #40.000-50.000
}

function Send-ModbusCommand
{
    param (
        [Parameter(Mandatory=$true)]
        [ModbusFunction]
        $Function,
        
        [Parameter(Mandatory=$true)]
        [string]
        $Address,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Port,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Reference,
        
        [Parameter(Mandatory=$true)]
        [uint16[]]
        $Payload,
        
        [byte]
        $Slave = 1
    )

    $ReferenceBytes = [System.BitConverter]::GetBytes($Reference)
    [array]::Reverse($ReferenceBytes)

    [byte[]]$PayloadBytes = @(0) * ($Payload.Length * 2)
    for ($i = 0; $i -lt $Payload.Length; $i++)
    {
        $p = $Payload[$i]
        $pBytes = [System.BitConverter]::GetBytes($p)
        [array]::Reverse($pBytes)
        [array]::Copy($pBytes, 0, $PayloadBytes, $i, 2)
    }

    [byte[]]$tail = $Slave,   # unit id
                    $Function # function id

    $tail += $ReferenceBytes
    if ($PayloadBytes.Length -gt 0)
    {
        if ($Function -eq [ModbusFunction]::WriteMultipleRegisters)
        {
            # add word count
            [uint16]$wordCount = $Payload.Length
            $wordCountBytes = [System.BitConverter]::GetBytes($wordCount)
            [array]::Reverse($wordCountBytes)
            $tail += $wordCountBytes
            $tail += [byte]$PayloadBytes.Length # byte count
        }
        $tail += $PayloadBytes
    }

    [uint16]$length = $tail.Length
    $lengthBytes = [System.BitConverter]::GetBytes($length)
    [array]::Reverse($lengthBytes)

    [byte[]]$data = 0x00, 0x01, # transaction
                    0x00, 0x00  # protocol

    $data += $lengthBytes
    $data += $tail

    #Write-Host $data

    [byte[]] $buffer = @(0) * 1024

    $tcpConnection = New-Object System.Net.Sockets.TcpClient($Address, $Port)
    $tcpConnection.ReceiveTimeout = 50000
    $tcpStream = $tcpConnection.GetStream()
    $tcpStream.Write($data,0,$data.length)
    $tcpStream.Flush()
    $size = $tcpStream.Read($buffer,0,$buffer.length)

    $result = @(0) * $size

    [array]::copy($buffer,$result,$size)

    $tcpConnection.Close()

    return $result
}

function Read-Registers
{
    param (
        [Parameter(Mandatory=$true)]
        [ModbusFunction]
        $Function,

        [Parameter(Mandatory=$true)]
        [string]
        $Address,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Port,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Reference,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Num,
        
        [Parameter()]
        [byte]
        $Slave = 1
    )
    
    if ($Function -ne [ModbusFunction]::ReadCoilStatus -and $Function -ne [ModbusFunction]::ReadHoldingRegister -and $Function -ne [ModbusFunction]::ReadInputRegister -and $Function -ne [ModbusFunction]::ReadInputStatus)
    {
        Write-Error "Function has to be one of the 'read registers' ones"
        return $null
    }

    $startIndex = 9

    $result = Send-ModbusCommand $Function $Address $Port $Reference @($Num) $Slave

    if ($result[7] -ne $Function)
    {
        $errCode = $result[7]
        Write-Error "Response error $errCode"
        return $null
    }
    $length = $result[8]

    [uint16[]]$output = @(0) * ($length / 2)
    if(-not $output) { $output = @(0) }
    
    for ($i = 0; $i -lt $length / 2; $i = $i + 1)
    {
        [byte[]]$slice = @(0) * 2
        [array]::Copy($result, $startIndex + ($i * 2), $slice, 0, $length)
        [array]::Reverse($slice)
        $output[$i] = [System.BitConverter]::ToUInt16($slice, 0)
    }
    return $output

}

function Read-HoldingRegister
{
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Address,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Port,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Reference,
        
        [Parameter(Mandatory=$true)]
        [uint16]$Num,
        
        [Parameter()]
        [byte]
        $Slave = 1
    )
    
    return Read-Registers ReadHoldingRegister $Address $Port $Reference $Num $Slave
}

function Read-InputRegister
{
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Address,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Port,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Reference,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Num,
        
        [Parameter()]
        [byte]
        $Slave = 1
    )
    
    return Read-Registers ReadInputRegiste $Address $Port $Reference $Num $Slave
}

function Write-SingleRegister
{
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Address,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Port,
        
        [Parameter(Mandatory=$true)]
        [ModbusFunction]
        $Function,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Reference,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Data,
        
        [Parameter()]
        [byte]
        $Slave = 1
    )
    
    $result = Send-ModbusCommand $Function $Address $Port $Reference @($Data) $Slave
    if ($result[7] -ne $Function)
    {
        $errCode = $result[7]
        Write-Error "Response error $errCode"
    }
}

function Write-MultipleRegisters
{
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Address,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Port,
        
        [Parameter(Mandatory=$true)]
        [uint16]
        $Reference,
        
        [Parameter(Mandatory=$true)]
        [uint16[]]
        $Data,
        
        [Parameter()]
        [byte]
        $Slave = 1
    )
    
    $result = Send-ModbusCommand WriteMultipleRegisters $Address $Port $Reference $Data $Slave
    if ($result[7] -ne [ModbusFunction]::WriteMultipleRegisters)
    {
        $errCode = $result[7]
        Write-Error "Response error $errCode"
    }
}

function Get-Value
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $ReferenceName,

        [Parameter(Mandatory=$true)]
        [string]
        $Ip
    )

    $reference = $allReferences | ? Name -eq $ReferenceName
    
    $value = Read-Registers -function $reference.ReadFunction -address $ip -port 502 -reference $reference.id -num 1
    
    if($reference.scale -gt 1)
    {
        $value = [int]$value / $reference.scale
    }

    if($reference.MeasuringUnit)
    {
        [string]$value += ' ' + $reference.MeasuringUnit
    }
    
    if($reference.ValueNames)
    {
        $value = ($reference.ValueNames | ? id -eq $value).value
    }
    
    return $value
}

function Set-Value
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $ReferenceName,

        [Parameter(Mandatory=$true)]
        [string]
        $Data,

        [Parameter(Mandatory=$true)]
        [string]
        $Ip
    )

    $reference = $allReferences | ? Name -eq $ReferenceName

    if($reference.Scale -gt 1) { $calculatedData = [int]$Data * $reference.Scale }
    elseif($data -eq "1") { $calculatedData = 65280 }
    else { $calculatedData = $Data }

    Write-Host "Old value: $(Get-Value -ReferenceName $reference.Name -Ip $Ip)"
    Write-SingleRegister -Function $reference.WriteFunction -address $ip -port 502 -Reference $reference.Id -Data $calculatedData
    Start-Sleep -Seconds 1
    Write-Host "New value: $(Get-Value -ReferenceName $reference.Name -Ip $Ip)"
}

function Get-Reference
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $ReferenceName
    )

    $reference = $allReferences | ? Name -eq $ReferenceName
      
    return $reference
}

