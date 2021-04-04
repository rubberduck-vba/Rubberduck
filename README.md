![banner](https://user-images.githubusercontent.com/5751684/113501222-8edfe880-94f1-11eb-99a9-64583e413ef3.png)

[**Installing**](https://github.com/rubberduck-vba/Rubberduck/wiki/Installing) • [Contributing](https://github.com/rubberduck-vba/Rubberduck/blob/next/CONTRIBUTING.md) • [Attributions](https://github.com/rubberduck-vba/Rubberduck/blob/next/docs/Attributions.md) • [Blog](https://rubberduckvba.blog) • [Wiki](https://github.com/rubberduck-vba/Rubberduck/wiki) • [rubberduckvba.com](https://rubberduckvba.com)

## Build Status

|Branch     | Build Status | Release notes &amp; Download Links | Donate |
|------------|--------------|-|:---:|
| **main** | ![main branch build status][mainBuildStatus] | [latest release](https://github.com/rubberduck-vba/Rubberduck/releases/latest) | <a href="https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=UY5K5X36B7T2S&currency_code=CAD&source=url"><img alt="Donate via PayPal" src="https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif"></a>|
| **next**   | ![next branch build status][nextBuildStatus] | [pre-releases](https://github.com/rubberduck-vba/Rubberduck/releases) | <p>via PayPal</p><sup>Pays for website and blog hosting fees. Donations in excess of our needs are going to <a href="https://mssociety.ca/">Multiple Sclerosis Society of Canada</a>.</sup> |

[nextBuildStatus]:https://ci.appveyor.com/api/projects/status/we3pdnkeebo4nlck/branch/next?svg=true
[mainBuildStatus]:https://ci.appveyor.com/api/projects/status/we3pdnkeebo4nlck/branch/main?svg=true

---

## [License (GPLv3)](https://github.com/rubberduck-vba/Rubberduck/blob/next/LICENSE)

Copyright &copy; 2014-2021 Rubberduck project contributors.

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the [GNU General Public License](https://www.gnu.org/licenses/gpl-3.0.en.html) for more details.

---

## [JetBrains](https://www.jetbrains.com) | [ReSharper](https://www.jetbrains.com/resharper/)

[![JetBrains ReSharper logo](https://cloud.githubusercontent.com/assets/5751684/20271309/616bb740-aa58-11e6-91c9-65287b740985.png)](https://www.jetbrains.com/resharper/)

Since the project's early days, JetBrains' Open-Source team has been supporting Rubberduck with free OSS licenses for all core contributors - and we deeply thank them for that. ReSharper has been not only a tool we couldn't do without; it's been an inspiration, the ultimate level of polished perfection to strive for in our own IDE add-in project. So just like you're missing out if you write VBA and you're not using Rubberduck, you're missing out if you write C# and aren't using ReSharper.

<sub>Note: Rubberduck is not a JetBrains product. JetBrains does not contribute and is not affiliated to the Rubberduck project in any way.</sub>

---

## What is Rubberduck?

The Visual Basic Editor (VBE) has stood still for over 20 years, and there is no chance a first-party update to the legacy IDE ever brings it up to speed with modern-day tooling. Rubberduck aims to bring the VBE into this century by doing exactly that.

Read more about contributing here:

[![contribute!](https://user-images.githubusercontent.com/5751684/113513709-071dcc80-9539-11eb-833d-d21532065306.png)](https://github.com/rubberduck-vba/Rubberduck/blob/next/CONTRIBUTING.md)

The add-in has *many* features - below is a quick overview.

### Enhanced Navigation

The Rubberduck *command bar* displays docstring for the current member

![command bar](https://user-images.githubusercontent.com/5751684/113501975-25fb6f00-94f7-11eb-9189-fcf2a0dd98da.png)

The *Code Explorer* drills down to member level, has a search bar, and lets you visualize your project as a virtual folder hierarchy organized just the way you need it.

All references to any identifier, whether defined in your project or any of its library references, are one click away. If it has a name, it can be navigated to.

### Static Code Analysis, Refactorings

Rubberduck analyses your code in various configurable ways and can help avoiding beginner mistakes, keeping a consistent programming style, and finding all sorts of potential bugs and problems. Many code inspections were implemented as a result of frequently-asked [VBA questions on Stack Overflow](https://stackoverflow.com/questions/tagged/vba), and in many occasions an automatic quick-fix is available.

Rename variables to meaningful identifiers without worrying about breaking something. Promote local variables to parameters, extract interfaces and methods out of a selection, encapsulate fields into properties; reorder and/or delete parameters, and automatically update all callers. 

### Unit Testing

Write code that *provably* works, by invoking it from small test procedures that setup the conditions for a test case and assert that the expected outcome happened. Rubberduck provides a rich MSTest-inspired API, and soon an experimental mocking framework (a COM-visible wrapper around [Moq](https://github.com/Moq)) that can automatically implement VBA interfaces and configure mock objects.

![test explorer](https://user-images.githubusercontent.com/5751684/113502368-fa2db880-94f9-11eb-954f-5735c15d4c3e.png)

### Smart Indenter

A port of the popular 32-bit add-in by Office Automation Ltd., whose legacy VB6 source code was generously made [freely available](https://github.com/rubberduck-vba/Rubberduck/tree/next/Rubberduck.SmartIndenter/Legacy) for Rubberduck under GPLv3 by the legendary Stephen Bullen and Rob Bovey themselves! Rubberduck will prompt to import your *Smart Indenter* settings on first load if detected.

### Annotations

Special comments that become a game changer with Rubberduck processing them: organize modules in your project using `@Folder` annotations, synchronize `VB_Description` and `VB_PredeclaredId` hidden attributes without manually exporting, editing, and re-importing modules with `@Description` and `@PredeclaredId` annotations.

### More?

Of course there's more! There's tooling to help synchronizing the project with files in a folder (for source/version control), some auto-completion features like self-closing parentheses and quotes; there's a regular expression assistant, a replacement for the VBE's *add/remove references* dialog, and so many other things to discover, and yet even more to implement.

---

## Tips

Rubberduck isn't a lightweight add-in and consumes a large amount of memory. So much, that working with a very large project could be problematic with a 32-bit host, and sometimes even with a 64-bit host. Here are a few tips to get the best out of your ducky.

- **Start small**: explore the features with a small test project first.
- **Refresh often**: refresh Rubberduck's parser every time you pause to read after modifying a module. The fewer files modified since the last parse, the faster the next parse.
- **Review inspection settings**: there are *many* inspections, and some of them may produce *a lot* of results if they're directly targeting something that's part of your coding style. Spawning tens of thousands of inspection results can significantly degrade performance.
- **Avoid late binding**: Rubberduck cannot resolve identifier references and thus cannot understand the code as well if even the VBA compiler is deferring all validations to run-time. Declare and return explicit data types, and cast from `Object` and `Variant` to a more specific type whenever possible.

Feel free to [ask for support](https://github.com/rubberduck-vba/Rubberduck/issues/new?assignees=&labels=support&template=support.md), we're always happy to help. You may also want to browse [Rubberduck questions on Stack Overflow](https://stackoverflow.com/questions/tagged/rubberduck) (mind post dates!) or ask a new one. If you have 20+ reputation on Stack Exchange, you can join the [dev chat](https://chat.stackexchange.com/rooms/14929/vba-rubberducking) and ask anyone anything about Rubberduck or VBA.

For more information please see [Getting Started](https://github.com/rubberduck-vba/Rubberduck/blob/next/docs/GettingStarted.md) in the project's wiki, and follow the project's blog for project updates and advanced VBA OOP reading material.
