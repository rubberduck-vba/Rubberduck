using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeFixes;
using Microsoft.CodeAnalysis.Diagnostics;
using NUnit.Framework;
using TestHelper;

namespace RubberduckCodeAnalysis.Test
{
    [TestFixture]
    public class ComVisibilityUnitTests : ComManagementAnalyzer
    {
        //No diagnostics expected to show up
        [Test]
        [Category("ComVisible")]
        public void NoCode()
        {
            var test = @"";

            VerifyCSharpDiagnostic(test);
        }

        //No diagnostics expected to show up
        [Test]
        [Category("ComVisible")]
        public void NotComVisible()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(false)]
  public class Bar
  {
  }
}
";

            VerifyCSharpDiagnostic(test);
        }

        //No diagnostics expected to show up
        [Test]
        [Category("ComVisible")]
        public void NoComVisibleAttribute()
        {
            var test = @"using Rubberduck.RegexAssistant.i18n;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant
{
    public interface IRegularExpression : IDescribable
    {
        IList<IRegularExpression> Subexpressions { get; }
    }

    public class ConcatenatedExpression : IRegularExpression
    {
        public ConcatenatedExpression(IList<IRegularExpression> subexpressions)
        {
            Subexpressions = subexpressions ?? throw new ArgumentNullException();
        }

        public string Description => AssistantResources.ExpressionDescription_ConcatenatedExpression;

        public IList<IRegularExpression> Subexpressions { get; }
    }

    public class AlternativesExpression : IRegularExpression
    {
        public AlternativesExpression(IList<IRegularExpression> subexpressions)
        {
            Subexpressions = subexpressions ?? throw new ArgumentNullException();
        }

        public string Description => string.Format(AssistantResources.ExpressionDescription_AlternativesExpression, Subexpressions.Count);

        public IList<IRegularExpression> Subexpressions { get; }
    }

    public class SingleAtomExpression : IRegularExpression
    {
        public readonly IAtom Atom;

        public SingleAtomExpression(IAtom atom)
        {
            Atom = atom ?? throw new ArgumentNullException();
        }

        public string Description => $""{Atom.Description} {Atom.Quantifier.HumanReadable()}."";

        public IList<IRegularExpression> Subexpressions => new List<IRegularExpression>(Enumerable.Empty<IRegularExpression>());

        public override bool Equals(object obj)
        {
            return obj is SingleAtomExpression other
                && other.Atom.Equals(Atom);
        }

        public override int GetHashCode()
        {
            return Atom.GetHashCode();
        }

    }

    public class ErrorExpression : IRegularExpression
    {
        private readonly string _errorToken;

        public ErrorExpression(string errorToken)
        {
            _errorToken = errorToken ?? throw new ArgumentNullException();
        }

        public string Description => string.Format(AssistantResources.ExpressionDescription_ErrorExpression, _errorToken);

        public IList<IRegularExpression> Subexpressions => new List<IRegularExpression>();
    }

    internal static class RegularExpression
    {

        /// <summary>
        /// We basically run a Chain of Responsibility here. At first we try to parse the whole specifier as one Atom.
        /// If this fails, we assume it's a ConcatenatedExpression and proceed to create one of these.
        /// That works well until we encounter a non-escaped '|' outside of a CharacterClass. Then we know that we actually have an AlternativesExpression.
        /// This means we have to check what we got back and add it to a List of subexpressions to the AlternativesExpression. 
        /// We then proceed to the next alternative (ParseIntoConcatenatedExpression consumes the tokens it uses) and keep adding to our subexpressions.
        /// 
        /// Note that Atoms (or more specifically Groups) can request a Parse of their subexpressions. 
        /// Also note that TryParseAtom is responsible for grabbing an Atom <b>and</b> it's Quantifier.
        /// </summary>
        /// <param name=""specifier"">The full Regular Expression specifier to Parse</param>
        /// <returns>An IRegularExpression that encompasses the complete given specifier</returns>
        public static IRegularExpression Parse(string specifier)
        {
            if (specifier == null)
            {
                throw new ArgumentNullException();
            }

            // ByRef requires us to hack around here, because TryParseAsAtom doesn't fail when it doesn't consume the specifier anymore
            var specifierCopy = specifier;
            if (TryParseAsAtom(ref specifierCopy, out var expression) && specifierCopy.Length == 0)
            {
                return expression;
            }
            var subexpressions = new List<IRegularExpression>();
            while (specifier.Length != 0)
            {
                expression = ParseIntoConcatenatedExpression(ref specifier);
                // ! actually an AlternativesExpression
                if (specifier.Length != 0 || subexpressions.Count != 0)
                {
                    // flatten hierarchy
                    var parsedSubexpressions = (expression as ConcatenatedExpression).Subexpressions;
                    if (parsedSubexpressions.Count == 1)
                    {
                        expression = parsedSubexpressions[0];
                    }
                    subexpressions.Add(expression);
                }
            }
            return (subexpressions.Count == 0) ? expression : new AlternativesExpression(subexpressions);
        }
        /// <summary>
        /// Successively parses the complete specifer into Atoms and returns a ConcatenatedExpression after the specifier has been exhausted or a single '|' is encountered at the start of the remaining specifier.
        /// Note: this may fail to work if the last encountered token cannot be parsed into an Atom, but the remaining specifier has nonzero lenght
        /// </summary>
        /// <param name=""specifier"">The specifier to Parse into a concatenated expression</param>
        /// <returns>The ConcatenatedExpression resulting from parsing the given specifier, either completely or up to the first encountered '|'</returns>
        private static IRegularExpression ParseIntoConcatenatedExpression(ref string specifier)
        {
            var subexpressions = new List<IRegularExpression>();
            var currentSpecifier = specifier;
            var oldSpecifierLength = currentSpecifier.Length + 1;
            while (currentSpecifier.Length > 0 && currentSpecifier.Length < oldSpecifierLength)
            {
                oldSpecifierLength = currentSpecifier.Length;
                // we actually have an AlternativesExpression, return the current status to Parse after updating the specifier
                if (currentSpecifier[0].Equals('|'))
                {
                    specifier = currentSpecifier.Substring(1); // skip leading |
                    return new ConcatenatedExpression(subexpressions);
                }
                if (TryParseAsAtom(ref currentSpecifier, out var expression))
                {
                    subexpressions.Add(expression);
                }
                else if (currentSpecifier.Length == oldSpecifierLength)
                {
                    subexpressions.Add(new ErrorExpression(currentSpecifier.Substring(0, 1)));
                    currentSpecifier = currentSpecifier.Substring(1);
                }
            }
            specifier = """"; // we've exhausted the specifier, tell Parse about it to prevent infinite looping
            return new ConcatenatedExpression(subexpressions);
        }

        private static readonly Regex groupWithQuantifier = new Regex($""^{Group.Pattern}{Quantifier.Pattern}?"", RegexOptions.Compiled);
        private static readonly Regex characterClassWithQuantifier = new Regex($""^{CharacterClass.Pattern}{Quantifier.Pattern}?"", RegexOptions.Compiled);
        private static readonly Regex literalWithQuantifier = new Regex($""^{Literal.Pattern}{Quantifier.Pattern}?"", RegexOptions.Compiled);
        /// <summary>
        /// Tries to parse the given specifier into an Atom. For that all categories of Atoms are checked in the following order:
        ///  1. Group
        ///  2. Class
        ///  3. Literal
        /// When it succeeds, the given expression will be assigned a SingleAtomExpression containing the Atom and it's Quantifier.
        /// The parsed atom will be removed from the specifier and the method returns true. To check whether the complete specifier was an Atom, 
        /// one needs to examine the specifier after calling this method. If it was, the specifier is empty after calling.
        /// </summary>
        /// <param name=""specifier"">The specifier to extract the leading Atom out of. Will be shortened if an Atom was successfully extracted</param>
        /// <param name=""expression"">The resulting SingleAtomExpression</param>
        /// <returns>True, if an Atom could be extracted, false otherwise</returns>
        // Note: could be rewritten to not consume the specifier and instead return an integer specifying the consumed length of specifier. This would remove the by-ref passed string hack
        // internal for testing
        internal static bool TryParseAsAtom(ref string specifier, out IRegularExpression expression)
        {
            var m = groupWithQuantifier.Match(specifier);
            if (m.Success)
            {
                var atom = m.Groups[""expression""].Value;
                var quantifier = m.Groups[""quantifier""].Value;
                specifier = specifier.Substring(atom.Length + 2 + quantifier.Length);
                expression = new SingleAtomExpression(new Group($""({atom})"", new Quantifier(quantifier)));
                return true;
            }
            m = characterClassWithQuantifier.Match(specifier);
            if (m.Success)
            {
                var atom = m.Groups[""expression""].Value;
                var quantifier = m.Groups[""quantifier""].Value;
                specifier = specifier.Substring(atom.Length + 2 + quantifier.Length);
                expression = new SingleAtomExpression(new CharacterClass($""[{atom}]"", new Quantifier(quantifier)));
                return true;
            }
            m = literalWithQuantifier.Match(specifier);
            if (m.Success)
            {
                var atom = m.Groups[""expression""].Value;
                var quantifier = m.Groups[""quantifier""].Value;
                specifier = specifier.Substring(atom.Length + quantifier.Length);
                expression = new SingleAtomExpression(new Literal(atom, new Quantifier(quantifier)));
                return true;
            }
            expression = null;
            return false;
        }
    }
}";

            VerifyCSharpDiagnostic(test);
        }
    }

    [TestFixture]
    public class MissingGuidUnitTests : ComManagementAnalyzer
    {
        //No diagnostics expected to show up
        [Test]
        [Category("MissingGuid")]
        public void ClassContainsGuid_SeparateAttribute()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(RubberduckGuid.BarGuid)]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] {test}, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.All(d => d.Descriptor.Id != "MissingGuid"));
        }

        //No diagnostics expected to show up
        [Test]
        [Category("MissingGuid")]
        public void ClassContainsGuid_AttributesOneLine()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true), Guid(RubberduckGuid.BarGuid)]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] {test}, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.All(d => d.Descriptor.Id != "MissingGuid"));
        }

        //No diagnostics expected to show up
        [Test]
        [Category("MissingGuid")]
        public void ClassContainsGuid_AttributesMultipleLines()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  
    [
        ComVisible(true), 
        Guid(RubberduckGuid.AccessibilityGuid)
    ]
    public class Bar
    {
    }
}
";
            var diagnostics = GetSortedDiagnostics(new[] {test}, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.All(d => d.Descriptor.Id != "MissingGuid"));
        }

        //No diagnostics expected to show up
        [Test]
        [Category("MissingGuid")]
        public void ClassContainsGuid_ConstantReference()
        {
            var test = @"using System.Runtime.InteropServices;
using Source = Rubberduck.Parsing.Symbols;

namespace Rubberduck.API.VBA
{
    [
        ComVisible(true), 
        Guid(RubberduckGuid.AccessibilityGuid)
    ]
    public enum Accessibility
    {
        Private = Source.Accessibility.Private,
        Friend = Source.Accessibility.Friend,
        Global = Source.Accessibility.Global,
        Implicit = Source.Accessibility.Implicit,
        Public = Source.Accessibility.Public,
        Static = Source.Accessibility.Static
    }
}
";
            VerifyCSharpDiagnostic(test);
        }

        [Test]
        [Category("MissingGuid")]
        public void LiteralGuid()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true), Guid(""69E0F697-43F0-3B33-B105-9B8188A6F040"")]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.Any(d => d.Descriptor.Id == "MissingGuid"));
        }

        [Test]
        [Category("MissingGuid")]
        public void MissingGuid()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] {test}, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.Any(d => d.Descriptor.Id == "MissingGuid"));
            /*
            var fixtest = @"
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Diagnostics;

    namespace ConsoleApplication1
    {
        class TYPENAME
        {   
        }
    }";
            VerifyCSharpFix(test, fixtest);
            */
        }

        //Diganostic triggered
        [Test]
        [Category("MissingGuid")]
        public void MissingGuid_CommentedOut()
        {
            var test = @"using System.Runtime.InteropServices;
using Source = Rubberduck.Parsing.Symbols;

namespace Rubberduck.API.VBA
{
    [
        ComVisible(true), 
        // Guid(RubberduckGuid.AccessibilityGuid)
    ]
    public enum Accessibility
    {
        Private = Source.Accessibility.Private,
        Friend = Source.Accessibility.Friend,
        Global = Source.Accessibility.Global,
        Implicit = Source.Accessibility.Implicit,
        Public = Source.Accessibility.Public,
        Static = Source.Accessibility.Static
    }
}
";
            var expected = new DiagnosticResult
            {
                Id = "MissingGuid",
                Message = string.Format("COM-visible type '{0}' does not have an explicit Guid attribute that references a RubberduckGuid constant.",
                    "Accessibility"),
                Severity = DiagnosticSeverity.Error,
                Locations =
                    new[]
                    {
                        new DiagnosticResultLocation("Test0.cs", 10, 17)
                    }
            };

            VerifyCSharpDiagnostic(test, expected);
        }
    }

    [TestFixture]
    public class MissingClassInterfaceUnitTests : ComManagementAnalyzer
    {
        //No diagnostics expected to show up
        [Test]
        [Category("MissingGuid")]
        public void ClassContainsInterfaceType()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(""69E0F697-43F0-3B33-B105-9B8188A6F040"")]
  [ClassInterface(ClassInterfaceType.None)]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] {test}, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.All(d => d.Descriptor.Id != "MissingClassInterface"));
        }
        
        [Test]
        [Category("MissingGuid")]
        public void ClassContainsInterfaceType_Wrong()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(""69E0F697-43F0-3B33-B105-9B8188A6F040"")]
  [ClassInterface(ClassInterfaceType.AutoDual)]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.Any(d => d.Descriptor.Id == "MissingClassInterface"));
        }

        [Test]
        [Category("MissingGuid")]
        public void ClassDoesNotContainsInterfaceType()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(""69E0F697-43F0-3B33-B105-9B8188A6F040"")]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.Any(d => d.Descriptor.Id == "MissingClassInterface"));
        }
    }

    [TestFixture]
    public class MissingProgIdUnitTests : ComManagementAnalyzer
    {
        //No diagnostics expected to show up
        [Test]
        [Category("MissingProgId")]
        public void ClassContainsProgId()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(RubberduckGuid.BarGuid)]
  [ClassInterface(ClassInterfaceType.None)]
  [ProgId(RubberduckProgId.BarProgId)]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.All(d => d.Descriptor.Id != "MissingProgId"));
        }

        [Test]
        [Category("MissingProgId")]
        public void ClassContainsProgId_Wrong()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(""69E0F697-43F0-3B33-B105-9B8188A6F040"")]
  [ProgId(""derp.derp"")]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.Any(d => d.Descriptor.Id == "MissingProgId"));
        }

        [Test]
        [Category("MissingProgId")]
        public void ClassDoesNotContainsProgId()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(""69E0F697-43F0-3B33-B105-9B8188A6F040"")]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.Any(d => d.Descriptor.Id == "MissingProgId"));
        }
    }

    [TestFixture]
    public class MissingComDefaultInterfaceUnitTests : ComManagementAnalyzer
    {
        //No diagnostics expected to show up
        [Test]
        [Category("MissingComDefaultInterface")]
        public void ClassContainsComDefaultInterface()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(RubberduckGuid.BarGuid)]
  [ClassInterface(ClassInterfaceType.None)]
  [ProgId(RubberduckProgId.BarProgId)]
  [ComDefaultInterface(typeof(IBar))]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.All(d => d.Descriptor.Id != "MissingComDefaultInterface"));
        }

        [Test]
        [Category("MissingComDefaultInterface")]
        public void ClassContainsComDefaultInterface_Wrong()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(""69E0F697-43F0-3B33-B105-9B8188A6F040"")]
  [ProgId(""derp.derp"")]
  [ComDefaultInterface(""IBar"")
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.Any(d => d.Descriptor.Id == "MissingComDefaultInterface"));
        }

        [Test]
        [Category("MissingComDefaultInterface")]
        public void ClassDoesNotContainsComDefaultInterface()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(""69E0F697-43F0-3B33-B105-9B8188A6F040"")]
  public class Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.Any(d => d.Descriptor.Id == "MissingComDefaultInterface"));
        }
    }

    [TestFixture]
    public class MissingInterfaceTypeUnitTests : ComManagementAnalyzer
    {
        //No diagnostics expected to show up
        [Test]
        [Category("MissingInterfaceType")]
        public void InterfaceContainsInterfaceType()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(RubberduckGuid.BarGuid)]
  [InterfaceType(ComInterfaceType.InterfaceIsDual)]
  public interface Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.All(d => d.Descriptor.Id != "MissingInterfaceType"));
        }

        [Test]
        [Category("MissingInterfaceType")]
        public void InterfaceContainsInterfaceType_Wrong()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [InterfaceType(blah)]
  public interface Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.Any(d => d.Descriptor.Id == "MissingInterfaceType"));
        }

        [Test]
        [Category("MissingInterfaceType")]
        public void InterfaceDoesNotContainsInterfaceType()
        {
            var test = @"using System.Runtime.InteropServices;

namespace Foo
{
  [ComVisible(true)]
  [Guid(""69E0F697-43F0-3B33-B105-9B8188A6F040"")]
  public interface Bar
  {
  }
}
";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.IsTrue(diagnostics.Any(d => d.Descriptor.Id == "MissingInterfaceType"));
        }
    }

    public interface IFoo
    {
        FooImp Execute();
    }

    public abstract class Foo : IFoo
    {
        public virtual FooImp Execute() { return new FooImp(); }
    }

    public class FooImp : Foo
    {
        public override FooImp Execute() { return base.Execute(); }
    }

    [TestFixture]
    public class ChainedWrapperUnitTests : ChainedWrapperAnalyzer
    {
        [Test]
        [Category("ChainedWrappers")]
        public void InterfaceContainsInterfaceType()
        {
            var test = @"namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ISafeComWrapper
    {
        FooImp Execute();
    }

    public abstract class Foo : ISafeComWrapper
    {
        public virtual FooImp Execute() { return new FooImp(); }
    }

    public class FooImp : Foo
    {
        public override FooImp Execute() { return base.Execute(); }
    }

    public class D
    {
        public void B()
        {
            var v = new FooImp();
            v.Execute().Execute();
        }
    }
}";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.AreEqual("ChainedWrapper", diagnostics.Single().Descriptor.Id);
        }
        [Test]
        [Category("ChainedWrappers")]
        public void InterfaceContainsInterfaceType_Property()
        {
            var test = @"namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ISafeComWrapper
    {
        FooImp Value { get; }
    }

    public abstract class Foo : ISafeComWrapper
    {
        public virtual FooImp Value => new FooImp();
    }

    public class FooImp : Foo
    {
        public override FooImp Value => base.Value;
    }

    public class D
    {
        public void B()
        {
            var v = new FooImp();
            var x = v.Value.Value;
        }
    }
}";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.AreEqual("ChainedWrapper", diagnostics.Single().Descriptor.Id);
        }
        [Test]
        [Category("ChainedWrappers")]
        public void InterfaceContainsInterfaceType_Property_ReverseAssignment()
        {
            var test = @"namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ISafeComWrapper
    {
        FooImp Value { get; set; }
    }

    public abstract class Foo : ISafeComWrapper
    {
        public virtual FooImp Value
        {
            get => new FooImp();
            set => new FooImp().Value = value;
        }
    }

    public class FooImp : Foo
    {
        public override FooImp Value => base.Value;
    }

    public class D
    {
        public void B()
        {
            var v = new FooImp();
            v.Value.Value = new FooImp();
        }
    }
}";
            var diagnostics = GetSortedDiagnostics(new[] { test }, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer());
            Assert.AreEqual("ChainedWrapper", diagnostics.Single().Descriptor.Id);
        }
    }

    public class ComManagementAnalyzer : CodeFixVerifier
    {
        protected override CodeFixProvider GetCSharpCodeFixProvider()
        {
            return null;
        }

        protected override DiagnosticAnalyzer GetCSharpDiagnosticAnalyzer()
        {
            return new ComVisibleTypeAnalyzer();

        }
    }

    public class ChainedWrapperAnalyzer : CodeFixVerifier
    {
        protected override CodeFixProvider GetCSharpCodeFixProvider()
        {
            return null;
        }

        protected override DiagnosticAnalyzer GetCSharpDiagnosticAnalyzer()
        {
            return new RubberduckCodeAnalysis.ChainedWrapperAnalyzer();

        }
    }
}

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract1
{
    public interface ISafeComWrapper
    {
        FooImp Value { get; set; }
    }

    public abstract class Foo : ISafeComWrapper
    {
        public virtual FooImp Value
        {
            get => new FooImp();
            set => new FooImp().Value = value;
        }
    }

    public class FooImp : Foo
    {
        public override FooImp Value => base.Value;
    }

    public class D
    {
        public void B()
        {
            var v = new FooImp();
            v.Value.Value = new FooImp();
        }
    }
}