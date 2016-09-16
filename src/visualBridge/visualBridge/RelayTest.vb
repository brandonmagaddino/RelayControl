Module RelayTest
    '===========================================================================================================================
    ' This module was written to be compatible with version 2.1.3.1 of FTDI FTD2XX.DLL
    ' Created January 2005
    ' FTDI
    ' www.ftdichip.com
    ' Modified Brian Hughe 01 Dec 2011
    '===========================================================================================================================


    '===========================================================================================================================
    ' FTD2XX Return codes
    '===========================================================================================================================
    Public Const FT_OK = 0
    Public Const FT_INVALID_HANDLE = 1
    Public Const FT_DEVICE_NOT_FOUND = 2
    Public Const FT_DEVICE_NOT_OPENED = 3
    Public Const FT_IO_ERROR = 4
    Public Const FT_INSUFFICIENT_RESOURCES = 5
    Public Const FT_INVALID_PARAMETER = 6
    Public Const FT_INVALID_BAUD_RATE = 7
    Public Const FT_DEVICE_NOT_OPENED_FOR_ERASE = 8
    Public Const FT_DEVICE_NOT_OPENED_FOR_WRITE = 9
    Public Const FT_FAILED_TO_WRITE_DEVICE = 10
    Public Const FT_EEPROM_READ_FAILED = 11
    Public Const FT_EEPROM_WRITE_FAILED = 12
    Public Const FT_EEPROM_ERASE_FAILED = 13
    Public Const FT_EEPROM_NOT_PRESENT = 14
    Public Const FT_EEPROM_NOT_PROGRAMMED = 15
    Public Const FT_INVALID_ARGS = 16
    Public Const FT_OTHER_ERROR = 17


    '===========================================================================================================================
    ' FTD2XX Flags - These are only used in this module
    '===========================================================================================================================
    ' FT_OpenEx Flags (See FT_OpenEx)
    Public Const FT_OPEN_BY_SERIAL_NUMBER = 1
    Public Const FT_OPEN_BY_DESCRIPTION = 2

    ' FT_ListDevices Flags (See FT_ListDevices)
    Public Const FT_LIST_NUMBER_ONLY = &H80000000
    Public Const FT_LIST_BY_INDEX = &H40000000
    Public Const FT_LIST_ALL = &H20000000


    '===========================================================================================================================
    ' FTD2XX Buffer Constants - These are only used in this module
    '===========================================================================================================================
    Const FT_In_Buffer_Size = &H100000                  ' 1024K
    Const FT_In_Buffer_Index = FT_In_Buffer_Size - 1
    Const FT_Out_Buffer_Size = &H10000                  ' 64K
    Const FT_Out_Buffer_Index = FT_Out_Buffer_Size - 1


    '===========================================================================================================================
    ' FTD2XX Constants
    '===========================================================================================================================
    ' FT Standard Baud Rates (See FT_SetBaudrate)
    Public Const FT_BAUD_300 = 300
    Public Const FT_BAUD_600 = 600
    Public Const FT_BAUD_1200 = 1200
    Public Const FT_BAUD_2400 = 2400
    Public Const FT_BAUD_4800 = 4800
    Public Const FT_BAUD_9600 = 9600
    Public Const FT_BAUD_14400 = 14400
    Public Const FT_BAUD_19200 = 19200
    Public Const FT_BAUD_38400 = 38400
    Public Const FT_BAUD_57600 = 57600
    Public Const FT_BAUD_115200 = 115200
    Public Const FT_BAUD_230400 = 230400
    Public Const FT_BAUD_460800 = 460800
    Public Const FT_BAUD_921600 = 921600

    ' FT Data Bits (See FT_SetDataCharacteristics)
    Public Const FT_DATA_BITS_7 = 7
    Public Const FT_DATA_BITS_8 = 8

    ' FT Stop Bits (See FT_SetDataCharacteristics)
    Public Const FT_STOP_BITS_1 = 0
    Public Const FT_STOP_BITS_2 = 2

    ' FT Parity (See FT_SetDataCharacteristics)
    Public Const FT_PARITY_NONE = 0
    Public Const FT_PARITY_ODD = 1
    Public Const FT_PARITY_EVEN = 2
    Public Const FT_PARITY_MARK = 3
    Public Const FT_PARITY_SPACE = 4

    ' FT Flow Control (See FT_SetFlowControl)
    Public Const FT_FLOW_NONE = &H0
    Public Const FT_FLOW_RTS_CTS = &H100
    Public Const FT_FLOW_DTR_DSR = &H200
    Public Const FT_FLOW_XON_XOFF = &H400

    ' Modem Status
    Public Const FT_MODEM_STATUS_CTS = &H10
    Public Const FT_MODEM_STATUS_DSR = &H20
    Public Const FT_MODEM_STATUS_RI = &H40
    Public Const FT_MODEM_STATUS_DCD = &H80

    ' FT Purge Commands (See FT_Purge)
    Public Const FT_PURGE_RX = 1
    Public Const FT_PURGE_TX = 2

    ' FT Bit Mode (See FT_SetBitMode)
    Public Const FT_RESET_BITMODE = &H0
    Public Const FT_ASYNCHRONOUS_BIT_BANG = &H1
    Public Const FT_MPSSE = &H2
    Public Const FT_SYNCHRONOUS_BIT_BANG = &H4
    Public Const FT_MCU_HOST = &H8
    Public Const FT_OPTO_ISOLATE = &H10

    ' FT Notification Event Masks (See FT_SetEventNotification)
    Public Const FT_EVENT_RXCHAR = 1
    Public Const FT_EVENT_MODEM_STATUS = 2
    Public Const WAIT_ABANDONED As Integer = &H80
    'Public Const WAIT_FAILED As Integer = &HFFFFFFFF
    Public Const WAIT_OBJECT_0 As Integer = &H0
    Public Const WAIT_TIMEOUT As Integer = &H102

    Public Const INFINITE As Integer = &HFFFFFFFF


    '===========================================================================================================================
    'Type definition for EEPROM as equivalent for C-structure "FT_PROGRAM_DATA"
    '===========================================================================================================================

    'Define string as Integer and use FT_EE_ReadEx and FT_EE_ProgramEx functions


    Public Structure FT_PROGRAM_DATA
        Dim Signature1 As Integer                  ' // Header - must be 0x00000000
        Dim Signature2 As Integer                  ' // Header - must be 0xFFFFFFFF
        Dim Version As Integer                     ' // 0 = original, 1 = FT2232C extensions
        Dim VendorID As Short                 ' // 0x0403
        Dim ProductID As Short                ' // 0x6001
        Dim Manufacturer As Integer                ' // "FTDI" (32 bytes allocated)
        Dim ManufacturerID As Integer              ' // "FT" (16 bytes allocated)
        Dim Description As Integer                 ' // "USB HS Serial Converter" (64 bytes allocated)
        Dim SerialNumber As Integer                ' // "FT000001" if fixed, or NULL (16 bytes allocated)
        Dim MaxPower As Short                 ' // 0 < MaxPower <= 500
        Dim PnP As Short                      ' // 0 = disabled, 1 = enabled
        Dim SelfPowered As Short              ' // 0 = bus powered, 1 = self powered
        Dim RemoteWakeup As Short             ' // 0 = not capable, 1 = capable
        ' Rev4 extensions:
        Dim Rev4 As Byte                        ' // true if Rev4 chip, false otherwise
        Dim IsoIn As Byte                       ' // true if in endpoint is isochronous
        Dim IsoOut As Byte                      ' // true if out endpoint is isochronous
        Dim PullDownEnable As Byte              ' // true if pull down enabled
        Dim SerNumEnable As Byte                ' // true if serial number to be used
        Dim USBVersionEnable As Byte            ' // true if chip uses USBVersion
        Dim USBVersion As Short               ' // BCD (0x0200 => USB2)
        ' FT2232C extensions:
        Dim Rev5 As Byte                        ' // non-zero if Rev5 chip, zero otherwise
        Dim IsoInA As Byte                      ' // non-zero if in endpoint is isochronous
        Dim IsoInB As Byte                      ' // non-zero if in endpoint is isochronous
        Dim IsoOutA As Byte                     ' // non-zero if out endpoint is isochronous
        Dim IsoOutB As Byte                     ' // non-zero if out endpoint is isochronous
        Dim PullDownEnable5 As Byte             ' // non-zero if pull down enabled
        Dim SerNumEnable5 As Byte               ' // non-zero if serial number to be used
        Dim USBVersionEnable5 As Byte           ' // non-zero if chip uses USBVersion
        Dim USBVersion5 As Short              ' // BCD (0x0200 => USB2)
        Dim AIsHighCurrent As Byte              ' // non-zero if interface is high current
        Dim BIsHighCurrent As Byte              ' // non-zero if interface is high current
        Dim IFAIsFifo As Byte                   ' // non-zero if interface is 245 FIFO
        Dim IFAIsFifoTar As Byte                ' // non-zero if interface is 245 FIFO CPU target
        Dim IFAIsFastSer As Byte                ' // non-zero if interface is Fast serial
        Dim AIsVCP As Byte                      ' // non-zero if interface is to use VCP drivers
        Dim IFBIsFifo As Byte                   ' // non-zero if interface is 245 FIFO
        Dim IFBIsFifoTar As Byte                ' // non-zero if interface is 245 FIFO CPU target
        Dim IFBIsFastSer As Byte                ' // non-zero if interface is Fast serial
        Dim BIsVCP As Byte                      ' // non-zero if interface is to use VCP drivers
    End Structure




    '===========================================================================================================================
    ' Public declarations of variables for external modules to access from FTD2XX.dll
    '===========================================================================================================================
    Public FT_Status As Integer
    Public FT_Device_Count As Integer
    Public FT_Serial_Number As String
    Public FT_Description As String
    Public FT_Handle As Integer
    Public FT_Type As String
    Public FT_VID_PID As String
    Public FT_Current_Baud As Integer
    Public FT_Current_DataBits As Byte
    Public FT_Current_StopBits As Byte
    Public FT_Current_Parity As Byte
    Public FT_Current_FlowControl As Integer
    Public FT_Current_XOn_Char As Byte
    Public FT_Current_XOff_Char As Byte
    Public FT_ModemStatus As Integer
    Public FT_RxQ_Bytes As Integer
    Public FT_TxQ_Bytes As Integer
    Public FT_EventStatus As Integer
    Public FT_Event_On As Boolean
    Public FT_Error_On As Boolean
    Public FT_Event_Value As Byte
    Public FT_Error_Value As Byte
    Public FT_In_Buffer(FT_In_Buffer_Index) As Byte
    Public FT_Out_Buffer(FT_Out_Buffer_Index) As Byte
    Public FT_Latency As Byte
    Public FT_EEPROM_DataBuffer As FT_PROGRAM_DATA
    Public FT_EEPROM_Manufacturer As String
    Public FT_EEPROM_ManufacturerID As String
    Public FT_EEPROM_Description As String
    Public FT_EEPROM_SerialNumber As String
    Public FT_UA_Size As Integer
    Public FT_UA_Data() As Byte     'NOTE: when using Read_EEPROM_UA and Write_EEPROM_UA get size of user area first,
    '      then use the command ReDim FT_UA_Data(0 to FT_UA_Size-1) As Byte

End Module
