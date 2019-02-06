# KukavarProxy

**KukavarProxy**  is a TCP/IP server that enables reading and writing robot variables over the network.

![KukavarProxy Form](https://github.com/lionpeloux/KukavarProxy/blob/dev/kukavarproxy.png)

## Credits

* Developed by [IMTS](www.imts.eu) - opensource release in january 2019
* Enhanced by Lionel du Peloux - february 2019

## Related documents
* Controlling Kuka Industrial Robots : Flexible Communication Interface JOpenShowVar [PDF](http://filipposanfilippo.inspitivity.com/publications/controlling-kuka-industrial-robots-flexible-communication-interface-jopenshowvar.pdf)

## Related repo
* [JOpenShowVar](https://github.com/aauc-mechlab/JOpenShowVar)
* [kukavarproxy-msg-format](https://github.com/akselov/kukavarproxy-msg-format)

## Contibuting

To contribute to this project, you should fork the repo and submit pull-request. **KukavarProxy**  is coded in Visual Basic 6.

To get a working development environment this is the steps you might follow :
* get a working installation of windows compatible with VB6 (mine is Win7 Pro x64)
* install [VB6 IDE](https://stackoverflow.com/questions/8029122/where-can-i-get-a-vb6-ide)
* install [VB6 SPS](https://www.microsoft.com/fr-fr/download/details.aspx?id=5721)

Prior to load the VB project, you need to register the crosscom components. Do do that, copy 
the content of the `lib` folder of this repo in the following directory :

```
C:\Windows\System32 (for a x86 os)
C:\Windows\SysWOW64 (for a x64 os)
```

Then, register the components with `regsvr32` utility. Start a `cmd` prompt with administrator rights (start > search 'cmd' > right-click on 'cmd' program > select 'run as admin'). Then run the following lines in the cmd prompt :

```
C:\Windows> cd C:\Windows\SysWOW64
C:\Windows\SysWOW64> regsvr32 Cross.ocx
C:\Windows\SysWOW64> regsvr32 cswsk32.ocx
```

Now, you should be able to open the visual basic project (.vbp) and reference the components needed to build the project.

## Install

To install **KukavarProxy** just copy the `KukavarProxy.exe` somewhere on your robot contorler (desktop is fine). Doubleclick to launch the server.

## Communicate with KukavarProxy

The proxy reveives read/write requests from clients over the network. The proxy has the ip address of the robot and is listening on port 7000.

The format of the messages exchange between clients and proxy must follow this protocol :

```
Read Request Message Format
---------------------------
2 bytes Id (uint16)
2 bytes for content length (uint16)
1 byte for read/write mode (0=Read)
2 bytes for the variable name length (uint16)
N bytes for the variable name to be read (ASCII)

Write Request Message Format
---------------------------
2 bytes Id (uint16)
2 bytes for content length (uint16)
1 byte for read/write mode (1=Write)
2 bytes for the variable name length (uint16)
N bytes for the variable name to be written (ASCII)
2 bytes for the variable value length (uint16)
M bytes for the variable value to be written (ASCII)

Answer Message Format
---------------------------
2 bytes Id (uint16)
2 bytes for content length (uint16)
1 byte for read/write mode (0=Read, 1=Write, 2=ReadArray, 3=WriteArray)
2 bytes for the variable value length (uint16)
N bytes for the variable value (ASCII)
3 bytes for tail (000 on error, 011 on success)
```



