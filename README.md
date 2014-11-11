Rubberduck
==============

A COM Add-In for the VBA IDE that makes VBA development even more enjoyable:

 - Unit testing
 - To-do items
 - Refactoring
 - ...

##Registry Keys

    [HKEY_CLASSES_ROOT\CLSID\{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}]
     ~> [@] ("Rubberduck.Extension")
     
    [HKEY_CLASSES_ROOT\CLSID\{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}\InprocServer32]
     ~> [@] ("mscoree.dll")
     ~> [ThreadingModel] ("Both")
     ~> [Class] ("Rubberduck.Extension")
     ~> [Assembly] ("Rubberduck")
     ~> [RuntimeVersion] ("v2.0.50727")
     ~> [CodeBase] ("file:///C:\Dev\RetailCoder\RetailCoder.VBE\RetailCoder.VBE\bin\Debug\Rubberduck.dll")
   
    [HKEY_CLASSES_ROOT\CLSID\{8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66}\InprocServer32]
     ~> [@] ("Rubberduck.Extension")

###Office x64

    [HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins64\Rubberduck]
     ~> [CommandLineSafe] (DWORD:00000000)
     ~> [Description] ("Rubberduck add-in for VBA IDE.")
     ~> [LoadBehavior] (DWORD:00000003)
     ~> [FriendlyName] ("Rubberduck")
   
###Office x86

    [HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0\Addins\Rubberduck]
     ~> [CommandLineSafe] (DWORD:00000000)
     ~> [Description] ("Rubberduck add-in for VBA IDE.")
     ~> [LoadBehavior] (DWORD:00000003)
     ~> [FriendlyName] ("Rubberduck")
   
---   

###Icons attribution

Fugue Icons

(C) 2012 Yusuke Kamiyamane. All rights reserved.
These icons are licensed under a Creative Commons
Attribution 3.0 License.
<http://creativecommons.org/licenses/by/3.0/>

If you can't or don't want to provide attribution, please
purchase a royalty-free license.
<http://p.yusukekamiyamane.com/>
