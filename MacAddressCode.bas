Attribute VB_Name = "MacAddressCode"
Option Explicit

' MAC address handling and generation methods v1.1.2
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.MacAddress
'
' Set of functions to retrieve, create, parse, verify, and format
' MAC addresses and their main properties.
' Also, generate BSSIDs derived from a MAC address and list these.
'
' Limitation: Only IPv4 is handled. Any IPv6 information is ignored or unhandled.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Required references:
'   Microsoft Scripting Runtime
'
' Required modules:
'   Internet


' Constants.
'
    ' Count of bytes/octets in a MAC48 address.
    Const OctetCount    As Integer = 6
    ' Length of a byte (octet).
    Const OctetLength   As Integer = 2
    ' Count of digits in a MAC48 address.
    Const DigitCount    As Integer = OctetCount * OctetLength
    ' Possible count of frames.
    Const FrameCount1   As Integer = 1
    Const FrameCount3   As Integer = 3
    Const FrameCount6   As Integer = 6
    ' Length of the frames in a MAC address having one, three, or six frames.
    Const FrameLength1  As Integer = DigitCount / FrameCount1
    Const FrameLength3  As Integer = DigitCount / FrameCount3
    Const FrameLength6  As Integer = DigitCount / FrameCount6
    ' Lenght of a MAC address having one, three, or six frames.
    Const TotalLength1  As Integer = DigitCount + FrameCount1 - 1
    Const TotalLength3  As Integer = DigitCount + FrameCount3 - 1
    Const TotalLength6  As Integer = DigitCount + FrameCount6 - 1
    ' HexPattern of one hex digit ignoring case.
    Const HexPattern    As String = "[0-9,A-Fa-f]"
'

' Enums.
'
    ' Enum for MAC address parsing and formatting.
    Public Enum IpMacAddressDelimiter
        ' Commonly accepted delimiters.
        ipMacNone = 0
        ipMacColon = 1
        ipMacDash = 2
        ipMacDot = 3
        ' Delimiter for temporary use only.
        ipMacStar = 4
    End Enum
    
    ' Enum for MAC address creation and verification.
    Public Enum IpMacAddressTransmissionType
        ipMacUnicast = 0
        ipMacMulticast = 1
    End Enum
    
    ' Enum for MAC address creation and verification.
    Public Enum IpMacAddressAministration
        ipMacUniversal = 0
        ipMacLocal = 1
    End Enum
    
    ' Enum for array to hold NIC information.
    '   0: Byte array. MAC address of NIC
    '   1: Boolean. NIC is IP enabled
    '   2: Boolean. NIC has been assigned a default IP gateway
    '   3: String. IP address (first if several)
    '   4: String. NIC description
    Public Enum IpNicInformation
        [_First] = 0
        ipNicMacAddress = 0
        ipNicIpEnabled = 1
        ipNicHasDefaultIpGateway = 2
        ipNicIpAddress = 3
        ipNicDescription = 4
        [_Last] = 4
    End Enum
'

' Retrieves from the local computer an array of the NICs having a MAC address.
'
' The array has four dimensions:
'   0: Byte array. MAC address of the NIC
'   1: Boolean. The NIC is IP enabled
'   2: Boolean. The NIC has been assigned a default IP gateway
'   3: String. IP address (first if several)
'   4: String. Description of the NIC.

' Returns the MAC address of the first network adapter having a gateway.
' If no adapter has a gateway, the MAC address of the first adapter is returned.
' If no IP enabled network adapter is found, a neutral MAC address is returned.
'
' Reference:
'   https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-networkadapterconfiguration
'
' 2019-09-23, Cactus Data ApS, Gustav Brock
'
Public Function GetMacAddresses() As Variant()

    ' This computer.
    Const Computer      As String = "."
    ' Namespace to access the Win32_NetworkAdapterConfiguration class.
    Const NameSpace     As String = "\root\cimv2"
    ' Query to list IP enabled network adapters.
    Const Sql           As String = "Select * From Win32_NetworkAdapterConfiguration Where MACAddress Is Not Null"
    
    Dim WMIService      As Object
    Dim Adapters        As Object
    Dim Adapter         As Object
    
    Dim Octets(0 To OctetCount - 1) As Byte
    Dim Nics()          As Variant
    Dim PathName        As String
    Dim AdapterCount    As Long
    Dim Index           As Long
    
    PathName = "winmgmts:" & "{impersonationLevel=impersonate}!\\" & Computer & NameSpace
    Set WMIService = GetObject(PathName)
    
    ' Retrieve the list of network adapters having a MAC address.
    Set Adapters = WMIService.ExecQuery(Sql)
    AdapterCount = Adapters.Count
    
    If AdapterCount > 0 Then
        ' Array to hold:
        '   0: Byte array. MAC address of NIC
        '   1: Boolean. NIC is IP enabled
        '   2: Boolean. NIC has been assigned a default IP gateway
        '   3: String. IP address (first if several)
        '   4: String. NIC description
        ReDim Nics(0 To AdapterCount - 1, IpNicInformation.[_First] To IpNicInformation.[_Last])
        
        ' Loop the network adapters to fill the array.
        For Each Adapter In Adapters
            Nics(Index, IpNicInformation.ipNicMacAddress) = MacAddressParse(Adapter.MacAddress)
            Nics(Index, IpNicInformation.ipNicIpEnabled) = Adapter.IPEnabled
            Nics(Index, IpNicInformation.ipNicHasDefaultIpGateway) = Not IsNull(Adapter.DefaultIPGateway)
            If Not IsNull(Adapter.IPAddress) Then
                Nics(Index, IpNicInformation.ipNicIpAddress) = Adapter.IPAddress(0)
            End If
            Nics(Index, IpNicInformation.ipNicDescription) = Adapter.Description
            Index = Index + 1
        Next
    Else
        ' No adapter having a MAC address was found.
        ReDim Nics(0, IpNicInformation.[_First] To IpNicInformation.[_Last])
        Nics(Index, IpNicInformation.ipNicMacAddress) = Octets()
        Nics(Index, IpNicInformation.ipNicIpEnabled) = False
        Nics(Index, IpNicInformation.ipNicHasDefaultIpGateway) = False
        Nics(Index, IpNicInformation.ipNicIpAddress) = ""
        Nics(Index, IpNicInformation.ipNicDescription) = "N/A"
    End If
    
    GetMacAddresses = Nics

End Function

' Retrieves the MAC address of the local computer as a byte array.
'
' Returns the MAC address of the first IP enabled network adapter having a gateway.
' If no adapter has a gateway, the MAC address of the first adapter is returned.
' If no IP enabled network adapter is found, a neutral MAC address is returned.
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Function MacAddressLocal() As Byte()

    Dim Nics()      As Variant
    Dim Index       As Long
    Dim Octets()    As Byte
    
    ' Retrieve array with list of adapters having a MAC address.
    Nics = GetMacAddresses()
    
    For Index = LBound(Nics) To UBound(Nics)
        If Not Nics(Index, IpNicInformation.ipNicIpEnabled) Then
            ' Ignore adapters that not are IP enabled.
        Else
            If Not IsMacAddress(Octets()) Then ' CBool(Len(CStr(Octets()))) Then
                ' First IP enabled NIC is found.
                Octets() = Nics(Index, IpNicInformation.ipNicMacAddress)
            End If
            If Nics(Index, IpNicInformation.ipNicHasDefaultIpGateway) Then
                ' First NIC assigned a gateway is found.
                Octets() = Nics(Index, IpNicInformation.ipNicMacAddress)
                Exit For
            End If
        End If
    Next
    
    If Not IsMacAddress(Octets()) Then ' CBool(Len(CStr(Octets()))) Then
        ' No IP enabled NIC was found.
        ' Return neutral MAC address.
        ReDim Octets(0 To OctetCount - 1)
    End If
    
    MacAddressLocal = Octets()

End Function

' Parses a string formatted MAC address and returns it as a Byte array.
' Parsing is not case sensitive.
' Will by default only accept the four de facto standard formats used widely.
'
' Examples:
'   "1234567890AB"          ->  1234567890AB
'   "1234.5678.90AB"        ->  1234567890AB
'   "12-34-56-78-90-AB"     ->  1234567890AB
'   "12:34:56:78:90:AB"     ->  1234567890AB
'
' If argument Exact is False, a wider variation of formats will be accepted:
'   "12-34:56-78:90-AB"     ->  1234567890AB
'   "12 34 56-78 90 AB"     ->  1234567890AB
'   "56 78 90 AB"           ->  0000567890AB
'   "1234567890ABDE34A0"    ->  1234567890AB
'
' For unparsable values, the neutral MAC address is returned:
'   "1K34567890ABDEA0"      ->  000000000000
'
' 2019-09-23, Cactus Data ApS, Gustav Brock
'
Public Function MacAddressParse( _
    ByVal MacAddress As String, _
    Optional Exact As Boolean = True) _
    As Byte()
        
    Dim Octets()    As Byte
    Dim Index       As Integer
    Dim Expression  As String
    Dim Match       As Boolean
    
    ' Delimiters.
    Dim Colon       As String
    Dim Dash        As String
    Dim Dot         As String
    Dim Star        As String
    
    ' Create neutral MAC address.
    ReDim Octets(0 To OctetCount - 1)
    
    ' Retrieve delimiter symbols.
    Colon = DelimiterSymbol(ipMacColon)
    Dash = DelimiterSymbol(ipMacDash)
    Dot = DelimiterSymbol(ipMacDot)
    Star = DelimiterSymbol(ipMacStar)
    
    If Exact = True Then
        ' Verify exact pattern of the passed MAC address.
        Select Case Len(MacAddress)
            Case TotalLength1
                ' One frame of six octets (no delimiter).
                Expression = Replace(Space(DigitCount), Space(1), HexPattern)
                Match = MacAddress Like Expression
                If Match = True Then
                    ' MAC address formatted as: 0123456789AB.
                End If
            Case TotalLength3
                ' Three frames of two octets.
                Expression = Replace(Replace(Replace(Space(DigitCount / FrameLength3), Space(1), Replace(Replace(Space(FrameLength3), Space(1), HexPattern), "][", "]" & Star & "[")), "][", "]" & Dot & "["), Star, "")
                Match = MacAddress Like Expression
                If Match = True Then
                    ' MAC address formatted as: 0123.4567.89AB.
                    MacAddress = Replace(MacAddress, Dot, "")
                End If
            Case TotalLength6
                ' Six frames of one octets.
                Expression = Replace(Replace(Replace(Space(DigitCount / FrameLength6), Space(1), Replace(Replace(Space(FrameLength6), Space(1), HexPattern), "][", "]" & Star & "[")), "][", "]" & Colon & "["), Star, "")
                Match = MacAddress Like Expression
                If Match = True Then
                    ' MAC address formatted as: 01:23:45:67:89:AB.
                    MacAddress = Replace(MacAddress, Colon, "")
                Else
                    Expression = Replace(Expression, Colon, Dash)
                    Match = MacAddress Like Expression
                    If Match = True Then
                        ' MAC address formatted as: 01-23-45-67-89-AB.
                        MacAddress = Replace(MacAddress, Dash, "")
                    End If
                End If
        End Select
    Else
        ' Non-standard format.
        ' Clean MacAddress and try to extract six octets.
        MacAddress = Replace(Replace(Replace(Replace(MacAddress, Colon, ""), Dash, ""), Dot, ""), Space(1), "")
        Select Case Len(MacAddress)
            Case Is > DigitCount
                ' Pick leading characters.
                MacAddress = Left(MacAddress, DigitCount)
            Case Is < DigitCount
                ' Fill with leading zeros.
                MacAddress = Right(String(DigitCount, "0") & MacAddress, DigitCount)
        End Select
        
        ' One frame of six possible octets.
        Expression = Replace(Space(DigitCount), Space(1), HexPattern)
        Match = MacAddress Like Expression
        If Match = True Then
            ' MAC address formatted as: 0123456789AB.
        End If
    End If
        
    If Match = True Then
        ' Fill array Octets.
        For Index = LBound(Octets) To UBound(Octets)
            Octets(Index) = Val("&H" & Mid(MacAddress, 1 + Index * OctetLength, OctetLength))
        Next
    End If
    
    MacAddressParse = Octets
    
End Function

' Formats a MAC address using one of the four de facto formats used widely.
' Thus, the format can and will be defined by the specified delimiter to use.
' The default is no delimiter and uppercase.
' Optionally, the case of the returned string can be specified as lowercase.
'
' Examples:
'   None        ->  "1234567890AB"
'   Dot         ->  "1234.5678.90AB"
'   Dash        ->  "12-34-56-78-90-AB"
'   Colon       ->  "12:34:56:78:90:AB"
'
'   Lowercase   ->  "1234567890ab"
'
' 2019-09-23, Cactus Data ApS, Gustav Brock
'
Public Function FormatMacAddress( _
    ByRef Octets() As Byte, _
    Optional Delimiter As IpMacAddressDelimiter, _
    Optional TextCase As VbStrConv = VbStrConv.vbProperCase) _
    As String
    
    Dim LastFrame   As Integer
    Dim ThisFrame   As Integer
    Dim FrameLength As Integer
    Dim Index       As Integer
    Dim Symbol      As String
    Dim MacAddress  As String
    
    ' Only accept an array with six octets.
    If LBound(Octets) = 0 And UBound(Octets) = OctetCount - 1 Then
    
        ' Calculate the frame length.
        FrameLength = DigitCount / DelimiterFrameCount(Delimiter)
        ' Format the octets using the specified delimiter.
        For Index = LBound(Octets) To UBound(Octets)
            ThisFrame = (Index * OctetLength) \ FrameLength
            Symbol = ""
            If LastFrame < ThisFrame Then
                Symbol = DelimiterSymbol(Delimiter)
                LastFrame = ThisFrame
            End If
            MacAddress = MacAddress & Symbol & Right("0" & Hex(Octets(Index)), OctetLength)
        Next
    End If
    
    If MacAddress <> "" Then
        Select Case TextCase
            Case VbStrConv.vbLowerCase
                MacAddress = StrConv(MacAddress, TextCase)
            Case Else
                ' Leave MacAddress in uppercase.
        End Select
    End If
    
    FormatMacAddress = MacAddress

End Function

' Returns the description of an IpMacAddressAministration value.
' By default, the description will be in proper case.
' Optionally, the description can be specified as upper- or lowercase.
'
' For invalid values, an empty string is returned.
'
' Examples:
'   IpMacAddressAministration.ipMacUniversal:
'       Default     ->  "Universal"
'       Uppercase   ->  "UNIVERSAL"
'       Lowercase   ->  "universal"
'
' 2019-09-23, Cactus Data ApS, Gustav Brock
'
Public Function FormatAdministration( _
    ByVal Administration As IpMacAddressAministration, _
    Optional TextCase As VbStrConv = VbStrConv.vbProperCase) _
    As String
    
    Dim Name        As String
    
    Select Case Administration
        Case IpMacAddressAministration.ipMacLocal
            Name = "Local"
        Case IpMacAddressAministration.ipMacUniversal
            Name = "Universal"
    End Select
    
    If Name <> "" Then
        Select Case TextCase
            Case VbStrConv.vbLowerCase, VbStrConv.vbUpperCase
                Name = StrConv(Name, TextCase)
        End Select
    End If
    
    FormatAdministration = Name
        
End Function

' Returns the description of an IpMacAddressTransmissionType value.
' By default, the description will be in proper case.
' Optionally, the description can be specified as upper- or lowercase.
'
' For invalid values, an empty string is returned.
'
' Examples:
'   IpMacAddressTransmissionType.ipMacUnicast:
'       Default     ->  "Unicast"
'       Uppercase   ->  "UNICAST"
'       Lowercase   ->  "unicast"
'
' 2019-09-23, Cactus Data ApS, Gustav Brock
'
Public Function FormatTransmissionType( _
    ByVal TransmissionType As IpMacAddressTransmissionType, _
    Optional TextCase As VbStrConv = VbStrConv.vbProperCase) _
    As String
    
    Dim Name        As String
    
    Select Case TransmissionType
        Case IpMacAddressTransmissionType.ipMacMulticast
            Name = "Multicast"
        Case IpMacAddressTransmissionType.ipMacUnicast
            Name = "Unicast"
    End Select
    
    If Name <> "" Then
        Select Case TextCase
            Case VbStrConv.vbLowerCase, VbStrConv.vbUpperCase
                Name = StrConv(Name, TextCase)
        End Select
    End If
    
    FormatTransmissionType = Name
        
End Function

' Returns the symbol of an IpMacAddressDelimiter value.
'
' For invalid values, an empty string is returned.
'
' 2019-09-23, Cactus Data ApS, Gustav Brock
'
Public Function DelimiterSymbol( _
    ByVal Delimiter As IpMacAddressDelimiter) _
    As String
    
    Dim Symbol      As String
    
    Select Case Delimiter
        ' Valid delimiters.
        Case IpMacAddressDelimiter.ipMacColon
            Symbol = ":"
        Case IpMacAddressDelimiter.ipMacDash
            Symbol = "-"
        Case IpMacAddressDelimiter.ipMacDot
            Symbol = "."
        ' Temporary delimiter only.
        Case IpMacAddressDelimiter.ipMacStar
            Symbol = "*"
    End Select
    
    DelimiterSymbol = Symbol
        
End Function

' Returns for an IpMacAddressDelimiter value the frame count of a MAC address.
'
' For invalid values, a count of 1 is returned.
'
' 2019-09-23, Cactus Data ApS, Gustav Brock
'
Public Function DelimiterFrameCount( _
    ByVal Delimiter As IpMacAddressDelimiter) _
    As Integer
    
    Dim FrameCount  As Integer
    
    Select Case Delimiter
        Case IpMacAddressDelimiter.ipMacColon
            FrameCount = 6
        Case IpMacAddressDelimiter.ipMacDash
            FrameCount = 6
        Case IpMacAddressDelimiter.ipMacDot
            FrameCount = 3
        Case Else
            FrameCount = 1
    End Select
    
    DelimiterFrameCount = FrameCount
        
End Function

' Creates a pseudo-random MAC address as a byte array.
' By default, it will be a multicast address marked as locally administered.
' Optionally, it can be a unicast address or marked as universally administered.
'
' Examples:
'   Default                 ->  7F:57:FA:7C:DD:7A
'   Unicast                 ->  42:7E:8A:9B:75:B2
'   Universal               ->  39:F3:7C:AD:ED:22
'   Unicast and Universal   ->  74:85:57:51:C4:37
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Function MacAddressRandom( _
    Optional TransmissionType As IpMacAddressTransmissionType = IpMacAddressTransmissionType.ipMacMulticast, _
    Optional Administration As IpMacAddressAministration = IpMacAddressAministration.ipMacLocal) _
    As Byte()
    
    Dim Octets(0 To OctetCount - 1) As Byte
    Dim Index       As Integer
    Dim Bit         As Byte
    Dim Octet       As Byte
    
    Randomize
    For Index = LBound(Octets) To UBound(Octets)
        ' Get random octet.
        Octet = Rnd * &HFF
        If Index = 0 Then
        
            ' Set transmission type.
            Bit = 2 ^ 0
            If TransmissionType = IpMacAddressTransmissionType.ipMacMulticast Then
                Octet = Octet Or Bit
            Else
                Octet = Octet And Not Bit
            End If
            
            ' Set administration.
            Bit = 2 ^ 1
            If Administration = IpMacAddressAministration.ipMacLocal Then
                Octet = Octet Or Bit
            Else
                Octet = Octet And Not Bit
            End If
            
        End If
        Octets(Index) = Octet
    Next
    
    MacAddressRandom = Octets
    
End Function

' Creates the neutral MAC address as a byte array.
'
' Example:
'   FormatMacAddress(MacAddressNeutral())   -> "000000000000"
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Function MacAddressNeutral() As Byte()

    Dim Octets(0 To OctetCount - 1) As Byte

    MacAddressNeutral = Octets()

End Function

' Retrieves the transmission type of a MAC address.
'
' Examples:
'   TransmissionTypeMacAddress(MacAddressParse("7F:57:FA:7C:DD:7A", True))  ->  1
'   TransmissionTypeMacAddress(MacAddressParse("42:7E:8A:9B:75:B2", True))  ->  0
'   TransmissionTypeMacAddress(MacAddressParse("39:F3:7C:AD:ED:22", True))  ->  1
'   TransmissionTypeMacAddress(MacAddressParse("74:85:57:51:C4:37", True))  ->  0
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Function TransmissionTypeMacAddress( _
    ByRef Octets() As Byte) _
    As IpMacAddressTransmissionType
    
    Const Bit       As Byte = 2 ^ 0
    
    Dim TransmissionType    As IpMacAddressTransmissionType
    
    If (Octets(0) And Bit) = Bit Then
        TransmissionType = ipMacMulticast
    Else
        TransmissionType = ipMacUnicast
    End If
    
    TransmissionTypeMacAddress = TransmissionType

End Function

' Retrieves the administration of a MAC address.
'
' Examples:
'   AdministrationMacAddress(MacAddressParse("7F:57:FA:7C:DD:7A", True))    ->  1
'   AdministrationMacAddress(MacAddressParse("42:7E:8A:9B:75:B2", True))    ->  1
'   AdministrationMacAddress(MacAddressParse("39:F3:7C:AD:ED:22", True))    ->  0
'   AdministrationMacAddress(MacAddressParse("74:85:57:51:C4:37", True))    ->  0
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Function AdministrationMacAddress( _
    ByRef Octets() As Byte) _
    As IpMacAddressAministration
    
    Const Bit       As Byte = 2 ^ 1
    
    Dim Administration      As IpMacAddressAministration
    
    If (Octets(0) And Bit) = Bit Then
        Administration = ipMacLocal
    Else
        Administration = ipMacUniversal
    End If
    
    AdministrationMacAddress = Administration

End Function

' Returns True if the passed MAC address is locally administered.
'
' Examples:
'   IsMacAddressAdministrationLocal(MacAddressParse("7F:57:FA:7C:DD:7A", True))     ->  True
'   IsMacAddressAdministrationLocal(MacAddressParse("42:7E:8A:9B:75:B2", True))     ->  True
'   IsMacAddressAdministrationLocal(MacAddressParse("39:F3:7C:AD:ED:22", True))     ->  False
'   IsMacAddressAdministrationLocal(MacAddressParse("74:85:57:51:C4:37", True))     ->  False
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Function IsMacAddressAdministrationLocal( _
    ByRef Octets() As Byte) _
    As Boolean
    
    Dim IsLocal     As Boolean
    
    If AdministrationMacAddress(Octets) = ipMacLocal Then
        IsLocal = True
    End If
    
    IsMacAddressAdministrationLocal = IsLocal

End Function

' Returns True if the passed MAC address is universally administered.
'
' Examples:
'   IsMacAddressAdministrationUniversal(MacAddressParse("7F:57:FA:7C:DD:7A", True)) ->  False
'   IsMacAddressAdministrationUniversal(MacAddressParse("42:7E:8A:9B:75:B2", True)) ->  False
'   IsMacAddressAdministrationUniversal(MacAddressParse("39:F3:7C:AD:ED:22", True)) ->  True
'   IsMacAddressAdministrationUniversal(MacAddressParse("74:85:57:51:C4:37", True)) ->  True
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Function IsMacAddressAdministrationUniversal( _
    ByRef Octets() As Byte) _
    As Boolean
    
    Dim IsUniversal As Boolean
    
    If AdministrationMacAddress(Octets) = ipMacUniversal Then
        IsUniversal = True
    End If
    
    IsMacAddressAdministrationUniversal = IsUniversal

End Function

' Returns True if the transmission type of the passed MAC address is multicast.
'
' Examples:
'   IsMacAddressTransmissionTypeMulticast(MacAddressParse("7F:57:FA:7C:DD:7A", True))   ->  True
'   IsMacAddressTransmissionTypeMulticast(MacAddressParse("42:7E:8A:9B:75:B2", True))   ->  False
'   IsMacAddressTransmissionTypeMulticast(MacAddressParse("39:F3:7C:AD:ED:22", True))   ->  True
'   IsMacAddressTransmissionTypeMulticast(MacAddressParse("74:85:57:51:C4:37", True))   ->  False
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Function IsMacAddressTransmissionTypeMulticast( _
    ByRef Octets() As Byte) _
    As Boolean
    
    Dim IsMulticast As Boolean
    
    If TransmissionTypeMacAddress(Octets) = ipMacMulticast Then
        IsMulticast = True
    End If
    
    IsMacAddressTransmissionTypeMulticast = IsMulticast

End Function

' Returns True if the transmission type of the passed MAC address is unicast.
'
' Examples:
'   IsMacAddressTransmissionTypeUnicast(MacAddressParse("7F:57:FA:7C:DD:7A", True))     ->  False
'   IsMacAddressTransmissionTypeUnicast(MacAddressParse("42:7E:8A:9B:75:B2", True))     ->  True
'   IsMacAddressTransmissionTypeUnicast(MacAddressParse("39:F3:7C:AD:ED:22", True))     ->  False
'   IsMacAddressTransmissionTypeUnicast(MacAddressParse("74:85:57:51:C4:37", True))     ->  True
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Function IsMacAddressTransmissionTypeUnicast( _
    ByRef Octets() As Byte) _
    As Boolean
    
    Dim IsUnicast   As Boolean
    
    If TransmissionTypeMacAddress(Octets) = ipMacUnicast Then
        IsUnicast = True
    End If
    
    IsMacAddressTransmissionTypeUnicast = IsUnicast

End Function

' Returns True if the passed byte array can hold a MAC address, and
' that this is not the neutral MAC address (00:00:00:00:00:00).
'
' Examples:
'   IsMacAddress(MacAddressRandom())    ->  True
'   IsMacAddress(MacAddressNeutral())   ->  False
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Function IsMacAddress( _
    ByRef Octets() As Byte) _
    As Boolean
    
    Dim Result      As Boolean
    
    If CBool(Len(Replace(CStr(Octets()), vbNullChar, vbNullString))) Then
        If LBound(Octets) = 0 And UBound(Octets) = OctetCount - 1 Then
            Result = True
        End If
    End If
    
    IsMacAddress = Result

End Function

' Retrieves as a string the vendor name from the OUI of a MAC address.
' Optionally, the vendor name is abbreviated to eight characters.
' For MAC addresses with an unknown OUI, and empty string is returned.
'
' Reference:
'   Open the URL (it is a text file) and study the header section.
'
' Examples:
'   GetMacAddressVendor(MacAddressParse("74:85:57:51:C4:37", True)) ->  ""
'   GetMacAddressVendor(MacAddressParse("44:37:E6:82:18:BB", True)) ->  "Hon Hai Precision Ind. Co.,Ltd."
'
' Requires:
'   Module: Internet
'
' 2019-10-02, Cactus Data ApS, Gustav Brock
'
Public Function GetMacAddressVendor( _
    ByRef Octets() As Byte, _
    Optional ByVal Abbreviated As Boolean) _
    As String

    ' Wireshark 'manuf' file.
    Const Url               As String = "https://code.wireshark.org/review/gitweb?p=wireshark.git;a=blob_plain;f=manuf;hb=HEAD"

    Static Path             As String
    
    Dim FileSystemObject    As Scripting.FileSystemObject
    Dim TextStream          As Scripting.TextStream
    Dim OuiFrame            As String
    Dim Line                As String
    Dim Vendor              As String
        
    If Dir(Path, vbNormal) = "" Then
        ' Caching of the downloaded file has timed out.
        Path = ""
    End If
    If Path = "" Then
        ' Download and cache file with list of vendors.
        Path = DownloadCacheFile(Url)
    Else
        ' File has been downloaded and saved in this session.
    End If

    OuiFrame = Left(FormatMacAddress(Octets, ipMacColon), 8)
    
    Set FileSystemObject = New Scripting.FileSystemObject
    Set TextStream = FileSystemObject.OpenTextFile(Path, ForReading)
    
    Do While Not TextStream.AtEndOfStream
        Line = TextStream.ReadLine
        If InStr(1, Line, OuiFrame, vbTextCompare) = 1 Then
            Vendor = Split(Line, vbTab)(2 - Abs(Abbreviated))
            Exit Do
        End If
    Loop
    TextStream.Close
    
    Set TextStream = Nothing
    Set FileSystemObject = Nothing

    GetMacAddressVendor = Vendor

End Function

' Returns one BSSID of the possible 32 BSSIDs derived from the passed MAC address.
' By default, the first BSSID is returned.
' Optionally, the Id argument (0 to 31) specifies which of the possible BSSIDs to be returned.
'
' Examples:
'   ' Octets() holds the MAC address d8:C7:C8:cc:43:24
'   FormatMacAddress(BssidMacAddress(Octets()), ipMacColon)     -> D8:C7:C8:44:32:40
'   FormatMacAddress(BssidMacAddress(Octets()) 12, ipMacColon)  -> D8:C7:C8:44:32:4C
'
' 2019-10-02, Cactus Data ApS, Gustav Brock
'
Public Function BssidMacAddress( _
    ByRef Octets() As Byte, _
    Optional Id As Byte) _
    As Byte()

    ' Maximum count of SSIDs.
    Const MaxSsid   As Integer = &H20
    
    Dim Bssid(0 To OctetCount - 1)  As Byte
    
    Dim Index       As Integer
    
    For Index = LBound(Octets) To UBound(Octets)
        Select Case Index
            Case 0 To 2
                ' Copy OUI.
                Bssid(Index) = Octets(Index)
            Case 3
                Bssid(Index) = (Octets(Index) And &HF) * &H10 Xor &H80 + Octets(Index + 1) / &H10
            Case 4
                Bssid(Index) = (Octets(Index) And &HF) * &H10 + Octets(Index + 1) / &H10
            Case 5
                Bssid(Index) = (Octets(Index) And &HF) * &H10 + (Id Mod MaxSsid)
        End Select
    Next
    
    BssidMacAddress = Bssid()

End Function

' Returns the possible 32 BSSIDs derived from the passed MAC address as an array of octets.
' By default, only the first BSSID is returned.
' Optionally, any other range of the possible BSSID can be returned:
'   Argument IdBase specifies the first BSSID to be returned.
'   Argument IdCount specifies the count of BSSIDs to be returned.
'
' Examples:
'   ' Octets() holds the MAC address d8:C7:C8:cc:43:24
'
'   Bssids() = BssidsMacAddress(Octets())
'   Bssids(0)   ->  D8:C7:C8:44:32:40
'
'   Bssids() = BssidsMacAddress(Octets(), 4, 3)
'   Bssids(0)   ->  D8:C7:C8:44:32:44
'   Bssids(1)   ->  D8:C7:C8:44:32:45
'   Bssids(2)   ->  D8:C7:C8:44:32:46
'
' 2019-10-02, Cactus Data ApS, Gustav Brock
'
Public Function BssidsMacAddress( _
    ByRef Octets() As Byte, _
    Optional ByVal IdBase As Byte, _
    Optional ByVal IdCount As Byte) _
    As Variant()

    Dim Bssids()    As Variant
    
    Dim Index       As Byte
    
    If IdCount = 0 Then
        ' Return minimum one BSSID.
        IdCount = 1
    End If
    
    ReDim Bssids(IdBase To IdBase + IdCount - 1)
    
    For Index = LBound(Bssids) To UBound(Bssids)
        Bssids(Index) = BssidMacAddress(Octets(), Index)
    Next
    
    BssidsMacAddress = Bssids()

End Function

