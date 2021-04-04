![banner-contributing](https://user-images.githubusercontent.com/5751684/113513403-93c78b00-9537-11eb-8262-9d47bc6b8d54.png)

Thank you for your interest in contributing to improve Rubberduck!

Below should be everything you need to know to get started with contributing to this project.

---

## GitHub

[Get started with GitHub](https://docs.github.com/en/github/getting-started-with-github) first, star ⭐ us if you haven't already (it helps boosting the project's visibility on the GitHub platform!), and you'll hit the ground running in no time.

## Non-Code Contributions

Don't know any C#? No problem, there is still plenty you can do to help!

### Issues

This project uses the GitHub issue tracker to document ideas and suggest enhancements, as well as bugs. Because this creates _lots_ of issues, we take the time to carefully label each one, to facilitate searching and browsing.

Please review the existing open [![bug](https://user-images.githubusercontent.com/5751684/113514335-ef941300-953b-11eb-9255-b070b79d90db.png)](https://github.com/rubberduck-vba/Rubberduck/issues?q=is%3Aissue+is%3Aopen+label%3Abug) issues (search filter `is:issue is:open label:bug`) to avoid duplicating a known issue.
 - 👍 upvote existing issues and ideas you'd wish were fixed or implemented.
 - 💬 share your feedback and ideas, even (especially!) if it's "just a small thing".
 - ✔️ submitting a bug? Please provide _and test_ clear steps to reproduce the problem, upload screenshots and/or include the logged exception stack trace(s), as applicable: the quicker we can track down a bug, the quicker we can fix it!
 - ❌ some excellent ideas are out-of-scope for this project; don't fret a rejected idea for an enhancement that is deemed too large or too complex to maintain. *It's not you, it's us!*
 - ❌ when posting a CLR/JIT stack trace, please avoid cluttering the issue with the "loaded assemblies" part; that (very) long list of loaded DLLs is not relevant.

### Wiki

The [wiki](https://github.com/rubberduck-vba/Rubberduck/wiki) is open for editing, and contributions are welcome.

 - ✏️ edit existing content to fix typos, update screenshots.
 - 📄 add new pages to document the source code and internal API.

### XML Documentation

The source code contains XML documentation that is compiled as an asset and that is ultimately used to generate content (English only) that is rendered on the website and sometimes accessible as an online resource from within Rubberduck.

Each page on the website generated from XML documentation in the source code, includes a link to 'edit this page' that points to the corresponding C# source file. With a GitHub login, it is easy to contribute enhancements to the documentation and content presented on the project's website.

You will find the XML documentation in the following folders/namespaces:

 - [Rubberduck.CodeAnalysis.Inspections.Concrete](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.CodeAnalysis/Inspections/Concrete/README.md)
 - [Rubberduck.CodeAnalysis.QuickFixes.Concrete](https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.CodeAnalysis/QuickFixes/Concrete/README.md)
 - [Rubberduck.Parsing.Annotations.Concrete](https://github.com/rubberduck-vba/Rubberduck/tree/next/Rubberduck.Parsing/Annotations/Concrete)

### Just Saying Thanks

Maintaining open-source software is often a thankless, selfless activity that involves a small and volunteering contributor force. If you ❤️ what we're doing here, we want to hear from you! Sure every ⭐ star gives us a warm fuzzy feeling of making a bit of a difference somewhere, but reading how our work impacts yours beats everything else.

[Tell us](https://github.com/rubberduck-vba/Rubberduck/issues/new?assignees=&labels=thanks&template=thanks-.md&title=) how you found out about Rubberduck, share that story about an obscure bug you'd been chasing for years that a Rubberduck inspection just single-handedly pointed out in a minute.

We don't do any telemetry, so we don't know how our users are using Rubberduck. Tell us what your favorite features are, what features you're perhaps not using but wish you'd be able to find more information about.

---

## Code Contributions

So you want to get your hands dirty and fix that obnoxious bug in the last release? Well that's great, we like seeing new faces in the fray 😄 Or you just want to find something that you can contribute to to learn C#? Or... Whatever your motivation is, if you're looking to add something new we recommend taking a look at issues that are labeled ![up for grabs](https://user-images.githubusercontent.com/5751684/113517076-004c8500-954c-11eb-97de-53fd714862b6.png); if you're looking to take down a known bug, you will want to pick from issues that are labeled with [![bug](https://user-images.githubusercontent.com/5751684/113514335-ef941300-953b-11eb-9255-b070b79d90db.png)](https://github.com/rubberduck-vba/Rubberduck/issues?q=is%3Aissue+is%3Aopen+label%3Abug).

Rubberduck is intimidatingly large project, even for long-time contributors. When an issue is deemed fairly simple to fix or implement, we label it with ![01 duckling](https://user-images.githubusercontent.com/5751684/113517296-0d1da880-954d-11eb-8350-edf758e9c485.png), ...or with increasing but always very subjective difficulty/complexity levels up to 
![04 quackhead](https://user-images.githubusercontent.com/5751684/113517348-65ed4100-954d-11eb-9133-c11459aa76e1.png) if you dare.

Level 3 and above demand a deeper knowledge of the inner workings of VBA (or VB6), the VBIDE, and Rubberduck's internal APIs - and we're always happy to help a contributor gain that knowledge!

We use GitHub's _Pull Request_ (PR) mechanism to share our work, so the first thing to do is to [fork the repository](https://docs.github.com/en/github/collaborating-with-issues-and-pull-requests/about-forks) and then [cloning](https://docs.github.com/en/github/getting-started-with-github/about-remote-repositories#cloning-with-https-urls) it to your local machine.

Once you have a local copy of the repository, use Git commands (or GitHub tooling in Visual Studio) to [commit](https://docs.github.com/en/github/getting-started-with-github/github-glossary#commit) your local changes, then [push](https://docs.github.com/en/github/getting-started-with-github/github-glossary#push) your local branch to your GitHub fork, and from there you can use GitHub to create your PR.

When you start work on an existing issue, comment on that issue to let us know you're on it and a pull request will be incoming as a result: don't hesitate to ask any questions you need answered to keep progressing, "issue" pages _exist_ for that type of discussion!

### Translations &amp; Resources

Rubberduck doesn't speak your language? Get [ResX Manager](https://marketplace.visualstudio.com/items?itemName=TomEnglert.ResXManager) (available both as a Visual Studio plug-in and as a standalone application) and have it create all the necessary resource keys and help you navigate and track the many .resx files in the solution.

> ⚠️ If you contribute a translation for a brand new previously unsupported language, please **keep in touch** with the dev team so that new resource keys can be translated when a new release is coming up; abandoned languages/translations **will** end up getting dropped.

Pull requests that improve wordings and translations are always welcome, too!
