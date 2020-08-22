using System;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.Common;

namespace Rubberduck.UnitTesting
{
    public class TestMethod : IEquatable<TestMethod>, INavigateSource
    {
        public TestMethod(Declaration declaration)
        {
            Declaration = declaration;
            TestCode = declaration.Context.GetText();
        }
        public Declaration Declaration { get; }

        public string TestCode { get; }

        public TestCategory Category
        {
            get
            {
                var testMethodAnnotation = Declaration.Annotations.First(pta => pta.Annotation is TestMethodAnnotation);
                var argument = testMethodAnnotation.AnnotationArguments.FirstOrDefault()?.FromVbaStringLiteral();

                var categorization = string.IsNullOrWhiteSpace(argument)
                    ? TestExplorer.TestExplorer_Uncategorized 
                    : argument;
                return new TestCategory(categorization);
            }
        }

        public NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(new QualifiedSelection(Declaration.QualifiedName.QualifiedModuleName, Declaration.Context.GetSelection()));
        }

        public bool IsIgnored => Declaration.Annotations.Any(a => a.Annotation is IgnoreTestAnnotation);
        
        public bool Equals(TestMethod other) => other != null && Declaration.QualifiedName.Equals(other.Declaration.QualifiedName) && TestCode.Equals(other.TestCode);

        public override bool Equals(object obj) => obj is TestMethod method && Equals(method);

        public override int GetHashCode() => Declaration.QualifiedName.GetHashCode();
    }
}
