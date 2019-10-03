Attribute VB_Name = "MacAddressDemo"
Option Explicit

' MAC address listing examples v1.0.1
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.MacAddress
'
' Set of example functions to list various information for network adapters of the local computer.
'
' Limitation: Only IPv4 is handled. Any IPv6 information is ignored or unhandled.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Requires:
'   Module MacAddressCode


' Constants.
'
    ' Row dimension of array.
    Const RowDimension      As Long = 1
'

' Lists general information for each of the network adapters of the local computer.
'
' Example:
'   MAC address   IP Enabled    Has gateway   IP address       Description
'   4437E68218AB  True          True          192.168.100.26   Hyper-V Virtual Ethernet Adapter
'   00155D011500  True          False         169.254.80.80    Hyper-V Virtual Ethernet Adapter #2
'   00155D4DB442  True          False         192.168.96.211   Hyper-V Virtual Ethernet Adapter #3
'   4437E68218AB  False         False                          Intel(R) 82579LM Gigabit Network Connection
'   E0FB20524153  False         False                          WAN Miniport (IP)
'   E0FB20524153  False         False                          WAN Miniport (IPv6)
'   E45E20524153  False         False                          WAN Miniport (Network Monitor)
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Sub ListLocalMacAddressesInfo()

    Const IpAddressWidth    As Long = 17

    Dim MacAddresses()      As Variant
    Dim Index               As Long
    Dim NicInformation      As IpNicInformation
    Dim Octets()            As Byte
    
    ' Retrieve the MAC addresses.
    MacAddresses = GetMacAddresses()
    
    ' Print a header line.
    Debug.Print "MAC address", "IP Enabled", "Has gateway", "IP address       Description"
    ' Loop the adapters.
    For Index = LBound(MacAddresses, RowDimension) To UBound(MacAddresses, RowDimension)
        For NicInformation = IpNicInformation.[_First] To IpNicInformation.[_Last]
            Select Case NicInformation
                Case IpNicInformation.ipNicMacAddress
                    Octets() = MacAddresses(Index, NicInformation)
                    Debug.Print FormatMacAddress(Octets()), ;
                Case IpNicInformation.ipNicIpAddress
                    Debug.Print Left(MacAddresses(Index, NicInformation) & Space(IpAddressWidth), IpAddressWidth);
                Case Else
                    Debug.Print MacAddresses(Index, NicInformation), ;
            End Select
        Next
        Debug.Print
    Next

End Sub

' Lists MAC address and vendor for each of the network adapters of the local computer.
'
' Example:
'   MAC address   Vendor
'   4437E68218AB  Hon Hai Precision Ind. Co.,Ltd.
'   00155D011500  Microsoft Corporation
'   00155D4DB442  Microsoft Corporation
'   4437E68218AB  Hon Hai Precision Ind. Co.,Ltd.
'   E0FB20524153
'   E0FB20524153
'   E45E20524153
'
' 2019-09-21, Cactus Data ApS, Gustav Brock
'
Public Sub ListLocalMacAddressesVendor()

    Dim MacAddresses()      As Variant
    Dim Index               As Long
    Dim Vendor              As String
    Dim Octets()            As Byte
    
    ' Retrieve the MAC addresses.
    MacAddresses = GetMacAddresses()
    
    ' Print a header line.
    Debug.Print "MAC address", "Vendor"
    ' Loop the adapters.
    For Index = LBound(MacAddresses, RowDimension) To UBound(MacAddresses, RowDimension)
        Octets() = MacAddresses(Index, IpNicInformation.ipNicMacAddress)
        Debug.Print FormatMacAddress(Octets()), ;
        Vendor = GetMacAddressVendor(Octets())
        Debug.Print Vendor
    Next

End Sub

' Lists the derived BSSID based on a MAC address.
' By default, the first BSSID is listed.
' Optionally, any of the other BSSIDs can be listed.
'
' Examples:
'   ' Octets() holds the MAC address d8:C7:C8:cc:43:24
'
'   ListBssid Octets()      ->
'       NIC           d8:c7:c8:cc:43:24
'        0            d8:c7:c8:44:32:40
'
'   ListBssid Octets(), 7   ->
'       NIC           d8:c7:c8:cc:43:24
'        0            d8:c7:c8:44:32:47
'
' 2019-10-02, Cactus Data ApS, Gustav Brock
'
Public Sub ListBssid( _
    ByRef Octets() As Byte, _
    Optional Id As Byte)

    Dim Bssid()     As Byte
    
    ' Retrieve BSSID.
    Bssid() = BssidMacAddress(Octets(), Id)
    
    ' Print the base MAC address.
    Debug.Print "NIC", FormatMacAddress(Octets(), ipMacColon, vbLowerCase)
    ' Print the derived BSSID.
    Debug.Print Id, FormatMacAddress(Bssid(), ipMacColon, vbLowerCase)

End Sub

' Lists derived BSSIDs based on a MAC address.
' The range of BBSIDs is controlled by the constants.
'
' Example:
'   ' Octets() holds the MAC address d8:C7:C8:cc:43:24
'
'   ListBssids Octets()     ->
'   NIC           d8:c7:c8:cc:43:24
'    0            d8:c7:c8:44:32:40
'    1            d8:c7:c8:44:32:41
'    2            d8:c7:c8:44:32:42
'    3            d8:c7:c8:44:32:43
'    4            d8:c7:c8:44:32:44
'    5            d8:c7:c8:44:32:45
'    6            d8:c7:c8:44:32:46
'    7            d8:c7:c8:44:32:47
'    8            d8:c7:c8:44:32:48
'    9            d8:c7:c8:44:32:49
'    10           d8:c7:c8:44:32:4a
'    11           d8:c7:c8:44:32:4b
'    12           d8:c7:c8:44:32:4c
'    13           d8:c7:c8:44:32:4d
'    14           d8:c7:c8:44:32:4e
'    15           d8:c7:c8:44:32:4f
'    16           d8:c7:c8:44:32:50
'    17           d8:c7:c8:44:32:51
'
' 2019-10-02, Cactus Data ApS, Gustav Brock
'
Public Sub ListBssids(ByRef Octets() As Byte)

    Const IdBase    As Byte = 0
    Const IdCount   As Byte = 18
    
    Dim Bssids()    As Variant
    Dim Bssid()     As Byte
    Dim Id          As Byte
    
    ' Retrieve array of BSSIDs.
    Bssids() = BssidsMacAddress(Octets(), IdBase, IdCount)
    
    ' Print the base MAC address.
    Debug.Print "NIC", FormatMacAddress(Octets(), ipMacColon, vbLowerCase)
    ' List the derived BSSIDs.
    For Id = LBound(Bssids) To UBound(Bssids)
        Bssid() = Bssids(Id)
        Debug.Print Id, FormatMacAddress(Bssid(), ipMacColon, vbLowerCase)
    Next

End Sub

