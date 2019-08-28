[back to readme.md](https://github.com/rubberduck-vba/Rubberduck/blob/next/README.md)

## Software & Libraries

### [ANTLR](http://www.antlr.org/)

Since v1.2, tokenizing and parsing the VBA code is left to the parsing masters behind Antlr. We write the language's grammatical/syntactical rules into a file that Antlr processes to generate a lexer that can turn a string into a stream of tokens, a parser that turns that stream into a tree structure Rubberduck can work with. Everything starts with Antlr.

> *ANTLR (ANother Tool for Language Recognition) is a powerful parser generator for reading, processing, executing, or translating structured text or binary files. It's widely used to build languages, tools, and frameworks. From a grammar, ANTLR generates a parser that can build and walk parse trees.*

### [AvalonEdit](http://avalonedit.net)

Source code looks a lot better with syntax highlighting, and AvalonEdit excels at it. 

> AvalonEdit is a WPF-based text editor component. It was written by [Daniel Grunwald](https://github.com/dgrunwald) for the [SharpDevelop](http://www.icsharpcode.net/OpenSource/SD/) IDE. Starting with version 5.0, AvalonEdit is released under the [MIT license](http://opensource.org/licenses/MIT).



### [EasyHook](http://easyhook.github.io/index.html)

Without the EasyHook library, many of our more advanced Unit Testing features would simply not be possible.  This library really lives up to its name, and allows us to intercept and inspect traffic through VBE7.dll and other unmanaged libraries.
 
> EasyHook makes it possible to extend (via hooking) unmanaged code APIs with pure managed functions, from within a fully managed environment on 32- or 64-bit Windows XP SP2, Windows Vista x64, Windows Server 2008 x64, Windows 7, Windows 8.1, and Windows 10. 

EasyHook is released under the [MIT license](https://github.com/EasyHook/EasyHook#license).

### [WPF Localization Using RESX Files](http://www.codeproject.com/Articles/35159/WPF-Localization-Using-RESX-Files)

This library makes localizing WPF applications at runtime using resx files a breeze. Thank you [Grant Frisken](http://www.codeproject.com/script/Membership/View.aspx?mid=1079060)!

> Licensed under [The Code Project Open License](http://www.codeproject.com/info/cpol10.aspx) with the [author's permission](http://www.codeproject.com/Messages/5272045/Re-License.aspx) to re-release under the GPLv3.

### [Moq](https://github.com/moq)

Moq has always been powering Rubberduck's own unit test mocks, but as of v2.5 our VBA unit testing API includes a wrapper that basically lets you use Moq for your VBA unit tests, to configure a mock implementation of any class or interface your VBA code might depend on.

> **What is ANTLR?**

> *ANTLR (ANother Tool for Language Recognition) is a powerful parser generator for reading, processing, executing, or translating structured text or binary files. It's widely used to build languages, tools, and frameworks. From a grammar, ANTLR generates a parser that can build and walk parse trees.*

We're not doing half of what we could be doing with this amazing tool. Try it, see for yourself!

## Icons

We didn't come up with these icons ourselves! Here's who did what:

### [Fugue Icons](http://p.yusukekamiyamane.com/)

This beautiful suite of professional-grade icons packs over 3,570 icons (16x16). You name it, there's an icon for that.

> (C) 2012 Yusuke Kamiyamane. All rights reserved. 
These icons are licensed under a [Creative Commons Attribution 3.0 License](http://creativecommons.org/licenses/by/3.0/).
If you can't or don't want to provide attribution, please [purchase a royalty-free license](http://p.yusukekamiyamane.com/).

### [SharpDevelop](https://github.com/icsharpcode/SharpDevelop.git)

Icons in the `./Resources/Custom/` directory were created by (or modified using elements from) the SharpDevelop icon set licensed under the [MIT license](https://opensource.org/licenses/MIT).
