[back to readme.md](https://github.com/rubberduck-vba/Rubberduck/blob/next/README.md)

## Software & Libraries

### [Ninject](http://www.ninject.org)

**What is Ninject?**

> *Ninject is a lightning-fast, ultra-lightweight dependency injector for .NET applications. It helps you split your application into a collection of loosely-coupled, highly-cohesive pieces, and then glue them back together in a flexible manner. By using Ninject to support your software's architecture, your code will become easier to write, reuse, test, and modify.*

If you're into *Dependency Injection* and *Inversion of Control*, you'll appreciate Ninject's simple configuration API. Other than in the very specialized startup code, Ninject seems absent of Rubberduck's code base, which indeed knows nothing of Ninject or any IoC framework: don't let it fool you - Ninject is responsible for instantiating essentially.. every single class in the solution.

Rubberduck uses the following Ninject [extensions](http://www.ninject.org/extensions):

 - Ninject.Extensions.Conventions
 - Ninject.Extensions.Factory
 - Ninject.Extensions.Interception
 - Ninject.Extensions.Interception.DynamicProxy
 - Ninject.Extensions.NamedScope

### [ANTLR](http://www.antlr.org/)

As of v1.2, Rubberduck is empowered by the awesomeness of ANTLR.

> **What is ANTLR?**

> *ANTLR (ANother Tool for Language Recognition) is a powerful parser generator for reading, processing, executing, or translating structured text or binary files. It's widely used to build languages, tools, and frameworks. From a grammar, ANTLR generates a parser that can build and walk parse trees.*

We're not doing half of what we could be doing with this amazing tool. Try it, see for yourself!

### [LibGit2Sharp](https://github.com/libgit2/libgit2sharp)

**What is LibGit2Sharp?**

LibGit2Sharp is the library that has allowed us to integrate Git right into the VBA IDE (and as a nice bonus, expose a nice API that handles the nitty gritty of importing source files to and from the IDE to a repo for you).

> LibGit2Sharp brings all the might and speed of libgit2, a native Git implementation, to the managed world of .Net and Mono.

**Okay, so what is [libgit2](https://libgit2.github.com/)?**

> libgit2 is a portable, pure C implementation of the Git core methods provided as a re-entrant linkable library with a solid API, allowing you to write native speed custom Git applications in any language which supports C bindings.

Which basically means it's a reimplementation of Git in C. It also [happens to be the technology Microsoft uses for their own Git integration with Visual Studio](http://www.hanselman.com/blog/GitSupportForVisualStudioGitTFSAndVSPutIntoContext.aspx).

### [AvalonEdit](http://avalonedit.net)

Source code looks a lot better with syntax highlighting, and AvalonEdit excels at it. 

> AvalonEdit is a WPF-based text editor component. It was written by [Daniel Grunwald](https://github.com/dgrunwald) for the [SharpDevelop](http://www.icsharpcode.net/OpenSource/SD/) IDE. Starting with version 5.0, AvalonEdit is released under the [MIT license](http://opensource.org/licenses/MIT).

We're currently only using a tiny bit of this code editor's functionality (more to come!).

### [EasyHook](http://easyhook.github.io/index.html)

Without the EasyHook library, many of our more advanced Unit Testing features would simply not be possible.  This library really lives up to its name, and allows us to intercept and inspect traffic through VBE7.dll and other unmanged libraries.
 
> EasyHook makes it possible to extend (via hooking) unmanaged code APIs with pure managed functions, from within a fully managed environment on 32- or 64-bit Windows XP SP2, Windows Vista x64, Windows Server 2008 x64, Windows 7, Windows 8.1, and Windows 10. 

EasyHook is released under the [MIT license](https://github.com/EasyHook/EasyHook#license).

### [WPF Localization Using RESX Files](http://www.codeproject.com/Articles/35159/WPF-Localization-Using-RESX-Files)

This library makes localizing WPF applications at runtime using resx files a breeze. Thank you [Grant Frisken](http://www.codeproject.com/script/Membership/View.aspx?mid=1079060)!

> Licensed under [The Code Project Open License](http://www.codeproject.com/info/cpol10.aspx) with the [author's permission](http://www.codeproject.com/Messages/5272045/Re-License.aspx) to re-release under the GPLv3.

## Icons

We didn't come up with these icons ourselves! Here's who did what:

### [Fugue Icons](http://p.yusukekamiyamane.com/)

This beautiful suite of professional-grade icons packs over 3,570 icons (16x16). You name it, there's an icon for that.

> (C) 2012 Yusuke Kamiyamane. All rights reserved. 
These icons are licensed under a [Creative Commons Attribution 3.0 License](http://creativecommons.org/licenses/by/3.0/).
If you can't or don't want to provide attribution, please [purchase a royalty-free license](http://p.yusukekamiyamane.com/).

### [SharpDevelop](https://github.com/icsharpcode/SharpDevelop.git)

Icons in the `./Resources/Custom/` directory were created by (or modified using elements from) the SharpDevelop icon set licensed under the [MIT license](https://opensource.org/licenses/MIT).
