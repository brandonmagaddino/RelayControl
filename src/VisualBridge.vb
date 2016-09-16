Imports FTD2XX_NET.FTDI

Module Module1
    Dim Board As FTD2XX_NET.FTDI

    Dim devices As Integer
    Dim deviceToOpen As Integer
    Dim serial As String
    Dim description As String

    Sub Main()
        If (My.Application.CommandLineArgs.Count = 0 Or My.Application.CommandLineArgs.Count > 4) Then
            help()
            Exit Sub
        End If

        If My.Application.CommandLineArgs(0).ToString().Equals("help") Or My.Application.CommandLineArgs(0).ToString().Equals("-help") Or My.Application.CommandLineArgs(0).ToString().Equals("h") Or My.Application.CommandLineArgs(0).ToString().Equals("-h") Then
            help()
            Exit Sub
        End If
        If My.Application.CommandLineArgs(0).ToString().Equals("-count") Then
            count()
            Exit Sub
        End If
        If My.Application.CommandLineArgs(0).ToString().Equals("-list") Then
            list()
            Exit Sub
        End If
        If My.Application.CommandLineArgs(0).ToString().Equals("-cycle") Then
            cycleport()
            Exit Sub
        End If
        If My.Application.CommandLineArgs(0).ToString().Equals("-state") Then
            pinState()
            Exit Sub
        End If
        If My.Application.CommandLineArgs(0).ToString().Equals("-d") Then
            knownDevice()
            Exit Sub
        End If

        'if you got here, you didn't enter in a real command
        Console.WriteLine("Error: Unknown command")
        Console.WriteLine("VBridge.exe -help   -- for help with commands")
    End Sub
    Sub load(ByVal device As Integer)
        Board = New FTD2XX_NET.FTDI()
        FT_Status = Board.GetNumberOfDevices(devices)
        If FT_Status <> FT_OK Then
            Console.WriteLine("Failed to find number of devices")
            Exit Sub
        End If

        FT_Status = Board.OpenByIndex(device)
        If FT_Status <> FT_OK Then
            Console.WriteLine("Device " + device.ToString + " Failed to open!")
            Console.WriteLine("")
            Exit Sub
        End If

        ' Get serial number of device with index 0
        FT_Status = Board.GetSerialNumber(serial)
        If FT_Status <> FT_OK Then
            Console.WriteLine("Failed to find serial number")
            Exit Sub
        End If

        ' Get description of device with index 0
        Board.GetDescription(description)
        If FT_Status <> FT_OK Then
            Console.WriteLine("Failed to find device description")
            Exit Sub
        End If

        ' Reset device
        FT_Status = Board.ResetDevice()
        If FT_Status <> FT_OK Then
            Console.WriteLine("Failed to initial reset")
            Exit Sub
        End If

        ' Set Baud Rate
        FT_Status = Board.SetBaudRate(FT_BAUD_921600)
        If FT_Status <> FT_OK Then
            Console.WriteLine("Failed to set baud rate")
            Exit Sub
        End If

        Dim mask As Byte = 255 '1111 1111 (8 channel relay)

        ' Set Bit Bang
        FT_Status = Board.SetBitMode(mask, FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG)
        If FT_Status <> FT_OK Then
            Console.WriteLine("Failed to set bit mode")
            Board.Close()
            Exit Sub
        End If

        ' Set timeout sequance, 10second read, infinte load
        FT_Status = Board.SetTimeouts(10000, 0)
        If FT_Status <> FT_OK Then
            Console.WriteLine("Failed to set time out")
            Board.Close()
            Exit Sub
        End If
    End Sub
    Sub list()
        Board = New FTD2XX_NET.FTDI()

        Board.GetNumberOfDevices(devices)
        Console.WriteLine("=================")
        Console.WriteLine("Total Devices: " + devices.ToString)
        Console.WriteLine("=================")
        For i As Integer = 0 To devices - 1
            Board.OpenByIndex(i)
            Board.GetSerialNumber(serial)
            Board.GetDescription(description)
            Dim state As Byte = 0
            Board.GetPinStates(state)
            Board.ResetDevice()

            Console.WriteLine("Device: " + i.ToString)
            Console.WriteLine("Serial: " + serial.ToString)
            Console.WriteLine("Description: " + description.ToString)
            Console.WriteLine("Pin State: " + state.ToString)
        Next
        Board.Close()
    End Sub
    Sub count()
        Board = New FTD2XX_NET.FTDI()

        ' Get the number of devices attached
        FT_Status = Board.GetNumberOfDevices(devices)
        If FT_Status <> FT_OK Then
            Console.WriteLine("Failed to find number of devices")
            Board.Close()
            Exit Sub
        End If
        Console.WriteLine(devices.ToString)
        Board.Close()
    End Sub
    Sub knownDevice()
        If My.Application.CommandLineArgs.Count < 4 Then
            Console.WriteLine("erroginous usage: VBridge.exe.exe -d 0 -c 255")
            Console.WriteLine("-- Turns everything on")
            Board.Close()
            Exit Sub
        End If
        If Not My.Application.CommandLineArgs(0).ToString().Equals("-d") Or Not My.Application.CommandLineArgs(2).ToString().Equals("-c") Then
            Console.WriteLine("erroginous usage: VBridge.exe.exe -d 0 -c 255")
            Console.WriteLine("-- Turns everything on")
            Board.Close()
            Exit Sub
        End If
        Dim WriteInt As Integer = My.Application.CommandLineArgs(3).ToString
        If My.Application.CommandLineArgs(3).ToString().Count = 8 Then
            WriteInt = Convert.ToInt32(My.Application.CommandLineArgs(3).ToString(), 2)
        End If

        If WriteInt > 255 Or WriteInt < 0 Then
            Console.WriteLine("Error out of bounds: VBridge.exe -d <0-devices> -c <0-255>")
            Board.Close()
            Exit Sub
        End If
        load(My.Application.CommandLineArgs(1).ToString()) 'load all the drivers and state of the device

        If FT_Status <> FT_OK Then
            Console.WriteLine("Could not write to device, an error has occered in loading")
            Board.Close()
            Exit Sub
        End If

        Console.Write("triggering " + WriteInt.ToString)

        Dim arrBytes(1) As Byte
        Dim BytesWritten As Byte = 0
        arrBytes(0) = WriteInt
        ' Set Relays on
        FT_Status = Board.Write(arrBytes, 1, BytesWritten)

        Board.Close()
    End Sub
    Sub pinState()
        If (My.Application.CommandLineArgs.Count = 2) Then
            load(My.Application.CommandLineArgs(1).ToString())
            If FT_Status <> FT_OK Then
                Board.Close()
                Exit Sub
            End If

            Dim state As Byte = 0
            Board.GetPinStates(state)
            Console.WriteLine(state.ToString)
            Board.Close()
            Exit Sub
        Else
            Console.WriteLine("ERROR: VBridge.exe -state <Device number>")
            Exit Sub
        End If
    End Sub
    Sub cycleport()
        If (My.Application.CommandLineArgs.Count = 2) Then
            load(My.Application.CommandLineArgs(1).ToString())
            If FT_Status <> FT_OK Then
                Console.WriteLine("Error cycling device!")
                Board.Close()
                Exit Sub
            End If
            Board.CyclePort()
            Board.Close()
            Exit Sub
        Else
            Board = New FTD2XX_NET.FTDI()
            Board.CyclePort()

            Board.Close()

        End If
    End Sub
    Sub help()
        Console.WriteLine("=========")
        Console.WriteLine("Commands")
        Console.WriteLine("=========")
        Console.WriteLine("VBridge.exe -help -- Lists commands")
        Console.WriteLine("VBridge.exe -d <device #> -c <0-255> -- Triggers the relay base on integer")
        Console.WriteLine("VBridge.exe -d <device #> -c <00000000-11111111> -- Triggers the relay base on binary")
        Console.WriteLine("VBridge.exe -list -- lists devices and info")
        Console.WriteLine("VBridge.exe -state <device #> -- returns the pin state of the relay")
        Console.WriteLine("VBridge.exe -count -- lists amount of devices")
        Console.WriteLine("VBridge.exe -cycle <device #> -- restarts specific device, like unplugging/pluggin the USB")
        Console.WriteLine("VBridge.exe -cycle -- restarts all devices, like unplugging/pluggin the USB")

    End Sub
End Module
