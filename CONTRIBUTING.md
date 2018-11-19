# Contributing to Rubberduck

Thank you very much for taking the time to contribute!

The following is a set of guidelines and links to helpful resources to get you started.

## Code of Conduct

Rubberduck as of now has no formalized Code of Conduct.
Yet the following guidelines apply in no particular order:

- Apply common sense?!
- Show and deserve respect
- Be nice
- We don't bite :)
- Assume people's best intentions unless proven otherwise

## How can I contribute?

### Bug Reports

The development team is always happy to hear about bugs. Reported bugs and other issues are opportunities to make Rubberduck even better.

Before you report your bug, please check the [GitHub Issues](https://github.com/rubberduck-vba/Rubberduck/issues).
The team spends quite some time to label the issues nicely so that you can find already known problems.

To enable the fastest response possible, please include the following information in your Bug report wherever applicable:

- The Version of Rubberduck you are running.
- Steps to reproduce the issue (Code?).
- A **full and complete** logfile, preferrably at TRACE-level
- Screenshots?
- ...

When posting a CLR/JIT stack trace, please avoid cluttering the issue with the "loaded assemblies" part; that (very) long list of loaded DLL's isn't relevant.

### Suggesting Enhancements

The team is always happy to hear your ideas. Please do make sure you checked, whether someone else already requested something similar.
The label [enhancement] will be a good guide.

Do note that despite the rather large scope of Rubberduck, some things are just **too large** or **too complex** for Rubberduck to effectively support.
Please don't be discouraged if your idea gets declined, your feedback is very valuable.

As with Bug Reports the more exact you define what's your gripe, the quicker we can get back to you.

### The Wiki

The Rubberduck Wiki is open for editing to everybody that has a Github account.
This is very much intentional.
Everybody is welcome to improve and expand the wiki.

### Coding and Related

So you want to get your hands dirty and fix that obnoxious bug in the last release?
Well **great**, we like seeing new faces in the fray :)
Or you just want to find something that you can contribute to to learn C#? Or ...
Well whatever your motivation is, we recommend taking a look at issues that are labeled [\[up-for-grabs\]](https://github.com/rubberduck-vba/Rubberduck/issues?q=is%3Aissue+is%3Aopen+label%3Aup-for-grabs).

Our Ducky has come quite the long way and is intimidatingly large for a first-time (or even long-time) contributor.
To make it easier to find something you can do, all issues with that label are additionally labelled with a difficulty.
That difficulty level is a best-guess by the dev-team how complex or evil an issue will be.

If you're new to C#, we recommend you start with \[difficulty-01-duckling\]. The next step \[difficulty-02-ducky\] requires some knowledge about C#, but not so much about how Rubberduck works - 02/ducky issues usually make a good introduction to using Rubberduck's internal API.

Then there's \[difficulty-03-duck\] which requires a good handle on C# and how the components of Rubberduck play toghether.
And finally there's \[difficulty-04-quackhead\] for when you really want to bang your head against the wall.
They require both good knowledge of C# and a deep understanding of how not only Rubberduck, but also COM and well, the VBE, works.

In addition to this, a few people have put together some helpful resources about how Rubberduck works internally and help contributors tell left from right in the ~250k Lines of Code.

You can find their work in the [GitHub wiki](https://github.com/rubberduck-vba/Rubberduck/wiki). If something in there doesn't quite seem to match, it's probably because nobody got around to updating it yet. You're very welcome to **improve** the wiki by adding and correcting information. (see right above)

In case this doesn't quite help you or the information you need hasn't been added to the wiki, you can **always** ask questions.
Please do so in the ["War room"](https://chat.stackexchange.com/rooms/14929).
Note that you will need an account on any of the Stack Exchange sites with at least 20 reputation points to ask questions there.
Other than that, you can also open an issue with the label \[support\], but it may take longer to get back to that.

In that room the core team talks about the duck and whatever else comes up.
N.B.: The rules of the Stack Exchange Network apply to everything you say in that room. But basically those are the same rules as those in the [CoC](#Code_of_Conduct).

Whether you create an issue or a pull request, please avoid using `[` square brackets `]` in the title, and try to be as descriptive (but succinct) as possible. Square brackets in titles confuse our chat-bot, and end up rendering in weird broken ways in SE chat. Avoid riddles, bad puns and other would-be funny (or NSFW) titles that don't really *describe the issue or PR*.

### Translations

Rubberduck is localized in multiple languages.
All these translations have been provided by volunteers.
We welcome both new translations as well as improvements to current translations very much.

The resource files are RESX/XML files easily editable in any text editor, but our recommendation goes to the [ResXManager](https://marketplace.visualstudio.com/items?itemName=TomEnglert.ResXManager) Visual Studio plug-in:

![ResXManager in Visual Studio](https://user-images.githubusercontent.com/5751684/48683625-1a07f000-eb7c-11e8-914a-10ed80e85e29.png)

If you contribute a translation for a brand new previously unsupported language, please keep in touch with the dev team so that new resource keys can be translated when a new release is coming up; abandoned languages/translations will end up getting dropped.

## What comes out of it for me?

Well... the eternal gratitude of all Rubberduck users for one :wink:
Aside from that, all contributors are explicitly listed by name (or alias) in the "About Rubberduck" window.
