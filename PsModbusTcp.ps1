Set-StrictMode -Version 5.0

enum ModbusFunction {
    ReadHoldingRegisters = [byte]0x03
    ReadInputRegisters = [byte]0x04
}

function Send-ModbusCommand {
    param (
        [byte]$function,
        [string]$address,
        [uint16]$port,
        [uint16]$reference,
        [uint16]$num = 0
    )

    $tcpConnection = New-Object System.Net.Sockets.TcpClient($address, $port)
    $tcpStream = $tcpConnection.GetStream()

    $referenceBytes = [System.BitConverter]::GetBytes($reference)
    [array]::Reverse($referenceBytes)

    $numBytes = [System.BitConverter]::GetBytes($num)
    [array]::Reverse($numBytes)

    $data = 0x00, 0x01, # transaction
            0x00, 0x00, # protocol
            0x00, 0x06, # length
            0x01,       # unit id
            $function  # function id

    $data = $data + $referenceBytes + $numBytes

    if ($function -eq 0x03 -or $function -eq 0x04)
    {
        $data = $data + $numBytes
    }

    [byte[]] $buffer = @(0) * 30

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
        [byte]$function,
        [string]$address,
        [uint16]$port,
        [uint16]$reference,
        [uint16]$num
    )
    
    if ($function -ne 0x03 -and $function -ne 0x04)
    {
        Write-Error "Function has to be one of the 'read registers' ones"
        return $null
    }

    $result = Send-ModbusCommand $function $address $port $reference $num

    if ($result[7] -ne $function)
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
        $output[$i] = [System.BitConverter]::ToInt16($slice, 0)
    }
    return $output

}

function Read-HoldingRegisters {
    param (
        [string]$address,
        [uint16]$port,
        [uint16]$reference,
        [uint16]$num
    )
    
    return Read-Registers 0x03 $address $port $reference $num
}

function Read-InputRegisters {
    param (
        [string]$address,
        [uint16]$port,
        [uint16]$reference,
        [uint16]$num
    )
    
    return Read-Registers 0x04 $address $port $reference $num
}
