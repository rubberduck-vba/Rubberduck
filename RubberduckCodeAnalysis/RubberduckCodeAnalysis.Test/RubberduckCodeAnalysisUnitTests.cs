using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeFixes;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestHelper;

namespace RubberduckCodeAnalysis.Test
{
    [TestClass]
    public class UnitTest : CodeFixVerifier
    {
        //No diagnostics expected to show up
        [TestMethod]
        public void NoCode()
        {
            var test = @"";

            VerifyCSharpDiagnostic(test);
        }

        //No diagnostics expected to show up
        [TestMethod]
        public void GoodCode()
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

            VerifyCSharpDiagnostic(test);
        }

        //No diagnostics expected to show up
        [TestMethod]
        public void GoodCode2()
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

            VerifyCSharpDiagnostic(test);
        }

        //No diagnostics expected to show up
        [TestMethod]
        public void GoodCode3()
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

            VerifyCSharpDiagnostic(test);
        }

        //No diagnostics expected to show up
        [TestMethod]
        public void GoodCode4()
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
        
        //No diagnostics expected to show up
        [TestMethod]
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

        //Diagnostic triggered
        [TestMethod]
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
            var expected = new DiagnosticResult
            {
                Id = "ComVisibleClassMustHaveGuidAnalyzer",
                Message = string.Format("COM-visible type name '{0}' does not have an explicit Guid attribute", "Bar"),
                Severity = DiagnosticSeverity.Error,
                Locations =
                    new[] {
                            new DiagnosticResultLocation("Test0.cs", 6, 16)
                        }
            };

            VerifyCSharpDiagnostic(test, expected);
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

        //No diagnostics expected to show up
        [TestMethod]
        public void MissingGuid2()
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
                Id = "ComVisibleClassMustHaveGuidAnalyzer",
                Message = string.Format("COM-visible type name '{0}' does not have an explicit Guid attribute", "Accessibility"),
                Severity = DiagnosticSeverity.Error,
                Locations =
                    new[] {
                        new DiagnosticResultLocation("Test0.cs", 10, 17)
                    }
            };

            VerifyCSharpDiagnostic(test, expected);
        }

        protected override CodeFixProvider GetCSharpCodeFixProvider()
        {
            return null;
        }

        protected override DiagnosticAnalyzer GetCSharpDiagnosticAnalyzer()
        {
            return new ComVisibleClassMustHaveGuidAnalyzer();
        }
    }
}
