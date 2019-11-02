# VBA.MacAddress

![Help](https://raw.githubusercontent.com/GustavBrock/VBA.MacAddress/master/images/EE%20Header.png)

### MAC addresses in Access and Excel
The block of six octets of a MAC address represents a lot of challenges when it comes to reading, formatting, parsing, validation, and look up of vendor information. The functions presented here let you read, generate, format, store, list, and report MAC addresses and derived BSSIDs for most tasks. 

### Challenges
This is a typical task that should be simple, but surprisingly is not. The reason is that a computer may have several other interfaces than the network card (NIC) that connects it the LAN, be it wired or wireless. Also, if the computer hosts virtual machines, at least one virtual network card exists in addition to the physical network cards.

Thus, first step is to retrieve the list of MAC addresses of the interface cards, and the function **GetMacAddresses** does that.

You will notice, that the retrieved information of the NICs are collected and returned in an array where the first item, the MAC address, is yet an array - an array of octets. 
This method is used throughout the functions as a convenient method that doesn't require a specific format. Formatting (to a human readable string) is first done when an address is displayed.

### Presenting a MAC address

Having the array of octets, we now need it formatted for humans to read it. That can be done in any way that fits a specific purpose, but for general use four common formats exist designated by the separator used:



|Separator|Display|
|---------|:-----:|
|None|1234567890AB|
|Dot|1234.5678.90AB|
|Dash|12-34-56-78-90-AB|
|Colon|12:34:56:78:90:AB|

The format for the purpose is **FormatMacAddress**, and it can return the formatted string in uppercase or (as Microsoft often uses) lowercase.

### Code ###
Code has been tested with both 32-bit and 64-bit *Microsoft Access 2019* and *365*.

### Documentation ###
Full documentation can be found here:

![EE Logo](https://raw.githubusercontent.com/GustavBrock/VBA.Quartiles/master/images/EE%20Logo.png) 

[MAC addresses in Access and Excel](https://www.experts-exchange.com/articles/33827/MAC-addresses-in-Access-and-Excel.html)

Included is a Microsoft Access example application and Microsoft Excel example workbook.

<hr>

*If you wish to support my work or need extended support or advice, feel free to:*

<p>

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.MacAddress/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)