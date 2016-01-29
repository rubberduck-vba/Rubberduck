![Rubberduck](http://i.stack.imgur.com/vmqXM.png)

| Branch     | Build Status |
|------------|--------------|
| **master** | [![master branch build status][masterBuildStatus]][masterBuild] |
| **next**   | [![next branch build status][nextBuildStatus]][nextBuild] |

[nextBuild]:https://ci.appveyor.com/project/ckuhn203/rubberduck-3v9qv/branch/next
[nextBuildStatus]:https://ci.appveyor.com/api/projects/status/bfwl1pwu9eeqd11o/branch/next?svg=true
[masterBuild]:https://ci.appveyor.com/project/ckuhn203/rubberduck-3v9qv/branch/master
[masterBuildStatus]:https://ci.appveyor.com/api/projects/status/bfwl1pwu9eeqd11o/branch/master?svg=true

Rubberduck is a COM Add-In for the VBA IDE that makes VBA development even more enjoyable, by extending the Visual Basic Editor (VBE) with menus, toolbars and toolwindows that enable things we didn't even think were possible when we first started this project.

If you're learning VBA, Rubberduck can help you avoid a few common beginner mistakes, and can probably show you a trick or two - even if you're only ever writing *macros*. If you're a more advanced programmer, you will appreciate the richness of [Rubberduck's feature set](https://github.com/retailcoder/Rubberduck/wiki/Features).

[**Follow us on Twitter!**](https://twitter.com/rubberduckvba)

[**Rubberduck Wiki**](https://github.com/retailcoder/Rubberduck/wiki)

---

#Contributing

If you're a C# developer looking for a fun project to contribute to, feel free to fork the project and 
[come meet the devs in Code Review's "VBA" chatroom][chat] - we'll be happy to answer your questions and help you help us take the VBE into the 21st century!

Some issues are tagged with [help-wanted][helpwanted], but that doesn't mean we can't use some help with anything else in the project - if this project interests you, we want to hear from you!

We follow a [development branch workflow][branch], so please submit any Pull Requests to the `next` branch.

There is additonal information about [building the project in our wiki][install].

  [chat]:http://chat.stackexchange.com/rooms/14929/vba
  [helpwanted]:https://github.com/retailcoder/Rubberduck/labels/help-wanted
  [install]:https://github.com/retailcoder/Rubberduck/wiki/Building-&-Installation
  [branch]:https://github.com/retailcoder/Rubberduck/issues/288

---

#License

Rubberduck is a COM add-in for the VBA IDE (VBE).

Copyright (C) 2014-2015 Mathieu Guindon & Christopher McClellan

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see http://www.gnu.org/licenses/.

---

#Installation

Visit our releases page, [download the installer](https://github.com/retailcoder/Rubberduck/releases/latest), and run the Setup.exe.

If you're **upgrading** from version 1.0, you will need to completely uninstall it before installing the newest release. This isn't necessary when upgrading from newer versions. Also, be sure to back up the `rubberduck.config` file in the `\AppData\Roaming\Rubberduck\` directory prior to installation.

##System Requirements

- Windows Vista or more recent (tested on Win7 and Win8.1)
- .Net Framework 4.5
- Microsoft Office 97-2003 or higher

Please feel free to test it on other versions and [submit any bugs on our issue tracker](https://github.com/retailcoder/Rubberduck/issues).

**Note:** The 1.2 installer doesn't work for 64 bit versions of office. You can find the [work around here](https://rubberduckvba.wordpress.com/2015/03/11/64-bit-rubberduck-a-work-around-until-release-1-21-soon/).

---

#Attributions

##Software & Libraries

###[ANTLR](http://www.antlr.org/)

As of v1.2, Rubberduck is empowered by the awesomeness of ANTLR.

> **What is ANTLR?**

> *ANTLR (ANother Tool for Language Recognition) is a powerful parser generator for reading, processing, executing, or translating structured text or binary files. It's widely used to build languages, tools, and frameworks. From a grammar, ANTLR generates a parser that can build and walk parse trees.*

We're not doing half of what we could be doing with this amazing tool. Try it, see for yourself!

###[LibGit2Sharp](https://github.com/libgit2/libgit2sharp)

**What is LibGit2Sharp?**

LibGit2Sharp is the library that has allowed us to integrate Git right into the VBA IDE (and as a nice bonus, expose a nice API that handles the nitty gritty of importing source files to and from the IDE to a repo for you).

> LibGit2Sharp brings all the might and speed of libgit2, a native Git implementation, to the managed world of .Net and Mono.

**Okay, so what is [libgit2](https://libgit2.github.com/)?**

> libgit2 is a portable, pure C implementation of the Git core methods provided as a re-entrant linkable library with a solid API, allowing you to write native speed custom Git applications in any language which supports C bindings.

Which basically means it's a reimplementation of Git in C. It also [happens to be the technology Microsoft uses for their own Git integration with Visual Studio](http://www.hanselman.com/blog/GitSupportForVisualStudioGitTFSAndVSPutIntoContext.aspx).

##Icons

We didn't come up with these icons ourselves! Here's who did what:

###[Fugue Icons](http://p.yusukekamiyamane.com/)

This beautiful suite of professional-grade icons packs over 3,570 icons (16x16). You name it, there's an icon for that.

> (C) 2012 Yusuke Kamiyamane. All rights reserved. 
These icons are licensed under a [Creative Commons Attribution 3.0 License](http://creativecommons.org/licenses/by/3.0/).
If you can't or don't want to provide attribution, please [purchase a royalty-free license](http://p.yusukekamiyamane.com/).

###[Microsoft Visual Studio Image Library](http://www.microsoft.com/en-ca/download/details.aspx?id=35825)

Icons in the `./Resources/Microsoft/` directory are licensed under Microsoft's Software License Terms, must be used accordingly with their meaning / file name.

> You have a right to Use and Distribute these files. This means that you are free to copy and use these images in documents and projects that you create, but you may not modify them in anyway.

For more information, please see the EULAs in the [./Resources/Microsoft/ directory](https://github.com/retailcoder/Rubberduck/tree/master/RetailCoder.VBE/Resources/Microsoft).

 * [Visual Studio 2013 Image Library EULA](https://github.com/retailcoder/Rubberduck/blob/master/RetailCoder.VBE/Resources/Microsoft/Visual%20Studio%202013%20Image%20Library%20EULA.rtf)
 * [Visual Studio 2012 Image Library EULA](https://github.com/retailcoder/Rubberduck/blob/master/RetailCoder.VBE/Resources/Microsoft/Visual%20Studio%202012%20Image%20Library%20EULA.rtf)
