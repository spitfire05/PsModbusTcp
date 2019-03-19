Set-StrictMode -Version 5.0

enum ModbusFunction {
    ReadHoldingRegisters = [byte]0x03
    ReadInputRegisters = [byte]0x04
    WriteSingleRegister = [byte]0x06
    WriteMultipleRegisters = [byte]0x10
}

function Send-ModbusCommand {
    param (
        [Parameter(Mandatory=$true)][ModbusFunction]$Function,
        [Parameter(Mandatory=$true)][string]$Address,
        [Parameter(Mandatory=$true)][uint16]$Port,
        [Parameter(Mandatory=$true)][uint16]$Reference,
        [Parameter(Mandatory=$true)][uint16[]]$Payload,
        [byte]$Slave = 1
    )

    $ReferenceBytes = [System.BitConverter]::GetBytes($Reference)
    [array]::Reverse($ReferenceBytes)

    [byte[]]$PayloadBytes = @(0) * ($Payload.Length * 2)
    for ($i = 0; $i -lt $Payload.Length; $i++)
    {
        $p = $Payload[$i]
        $pBytes = [System.BitConverter]::GetBytes($p)
        [array]::Reverse($pBytes)
        [array]::Copy($pBytes, 0, $PayloadBytes, $i *2, 2)
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

    [byte[]] $buffer = @(0) * 1024

    $tcpConnection = New-Object System.Net.Sockets.TcpClient($Address, $Port)
    $tcpConnection.ReceiveTimeout = 5000
    $tcpStream = $tcpConnection.GetStream()
    $tcpStream.Write($data,0,$data.length)
    $tcpStream.Flush()
    $size = $tcpStream.Read($buffer,0,$buffer.length)

    $result = @(0) * $size

    [array]::copy($buffer,$result,$size)

    $tcpConnection.Close()

    return $result
}

function Read-Registers {
    param (
        [Parameter(Mandatory=$true)][ModbusFunction]$Function,
        [Parameter(Mandatory=$true)][string]$Address,
        [Parameter(Mandatory=$true)][uint16]$Port,
        [Parameter(Mandatory=$true)][uint16]$Reference,
        [Parameter(Mandatory=$true)][uint16]$Num,
        [byte]$Slave = 1
    )
    
    if ($Function -ne [ModbusFunction]::ReadHoldingRegisters -and $Function -ne [ModbusFunction]::ReadInputRegisters)
    {
        Write-Error "Function has to be one of the 'read registers' ones"
        return $null
    }

    $result = Send-ModbusCommand $Function $Address $Port $Reference @($Num) $Slave
    if ($result[7] -ne $Function)
    {
        Write-Error "Response error " + $result[7]
        return $null
    }
    $length = $result[8]
    [uint16[]]$output = @(0) * ($length / 2)
    for ($i = 0; $i -lt $length / 2; $i = $i + 1)
    {
        [byte[]]$slice = @(0) * 2
        [array]::Copy($result, 9 + ($i * 2), $slice, 0, 2)
        [array]::Reverse($slice)
        $output[$i] = [System.BitConverter]::ToUInt16($slice, 0)
    }
    return $output

}

function Read-HoldingRegisters {
    param (
        [Parameter(Mandatory=$true)][string]$Address,
        [Parameter(Mandatory=$true)][uint16]$Port,
        [Parameter(Mandatory=$true)][uint16]$Reference,
        [Parameter(Mandatory=$true)][uint16]$Num,
        [byte]$Slave = 1
    )
    
    return Read-Registers ReadHoldingRegisters $Address $Port $Reference $Num $Slave
}

function Read-InputRegisters {
    param (
        [Parameter(Mandatory=$true)][string]$Address,
        [Parameter(Mandatory=$true)][uint16]$Port,
        [Parameter(Mandatory=$true)][uint16]$Reference,
        [Parameter(Mandatory=$true)][uint16]$Num,
        [byte]$Slave = 1
    )
    
    return Read-Registers ReadInputRegisters $Address $Port $Reference $Num $Slave
}

function Write-SingleRegister {
    param (
        [Parameter(Mandatory=$true)][string]$Address,
        [Parameter(Mandatory=$true)][uint16]$Port,
        [Parameter(Mandatory=$true)][uint16]$Reference,
        [Parameter(Mandatory=$true)][uint16]$Data,
        [byte]$Slave = 1
    )
    
    $result = Send-ModbusCommand WriteSingleRegister $Address $Port $Reference @($Data) $Slave
    if ($result[7] -ne [ModbusFunction]::WriteSingleRegister)
    {
        Write-Error "Response error " + $result[7]
    }
}

function Write-MultipleRegisters {
    param (
        [Parameter(Mandatory=$true)][string]$Address,
        [Parameter(Mandatory=$true)][uint16]$Port,
        [Parameter(Mandatory=$true)][uint16]$Reference,
        [Parameter(Mandatory=$true)][uint16[]]$Data,
        [byte]$Slave = 1
    )
    
    $result = Send-ModbusCommand WriteMultipleRegisters $Address $Port $Reference $Data $Slave
    if ($result[7] -ne [ModbusFunction]::WriteMultipleRegisters)
    {
        Write-Error "Response error " + $result[7]
    }
}
