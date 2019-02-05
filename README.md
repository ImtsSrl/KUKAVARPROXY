# KukavarProxy

KukavarProxy is a TCP/IP server that enables reading and writing robot variables over the network.

## Credits

* Developed by [IMTS](www.imts.eu) - opensource release in january 2019
* Enhanced by Lionel du Peloux - february 2019

## Contibuting

To contribute to this project, you should fork the repo and submit pull-request. KukavarProxy is coded in Visual Basic 6.

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
