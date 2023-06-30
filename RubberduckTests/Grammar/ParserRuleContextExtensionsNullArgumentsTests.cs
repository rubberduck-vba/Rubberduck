using Antlr4.Runtime;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Grammar
{
    [TestFixture]
    public class ParserRuleContextExtensionsNullArgumentsTests
    {
        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetSelection_nullContext_returnsHome()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.GetSelection();

            //Assert
            Assert.AreEqual(Selection.Home, actual);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void Contains_nullContainedContext_returnsFalse(bool containingContextIsNull)
        {
            //Arrange
            var inputCode =
@"
Public Function Foo(arg1 As Long) As Long
End Function"
;
            var aContext = GetUserDeclaration(inputCode, "Foo").Context;
            var containingContext = containingContextIsNull ? null : GetUserDeclaration(inputCode, "Foo").Context;
            var containedContext = containingContextIsNull ? GetUserDeclaration(inputCode, "Foo").Context : null;

            //Act
            var actual = containingContext.Contains(containedContext);

            //Assert
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetTokens_nullContext_returnsEmptyList()
        {
            //Arrange
            var mockTokenStream = new Mock<ITokenSource>();
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.GetTokens(new CommonTokenStream(mockTokenStream.Object)).Count();

            //Assert
            Assert.AreEqual(0, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetText_nullContext_returnsEmptyString()
        {
            //Arrange
            var mockCharStream = new Mock<ICharStream>();
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.GetText(mockCharStream.Object);

            //Assert
            Assert.AreEqual(string.Empty, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetChild_nullContext_returnsNull()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.GetChild<ParserRuleContext>();

            //Assert
            Assert.IsNull(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void IsDescendentOf_nullContext_returnsFalse()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.IsDescendentOf<ParserRuleContext>();

            //Assert
            Assert.IsFalse(actual);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void IsDescendentOf_Ancestor_nullContext_returnsFalse(bool ancestorIsNull)
        {
            //Arrange
            var inputCode =
@"
Public Function Foo(arg1 As Long) As Long
End Function"
;
            var aValidContext = GetUserDeclaration(inputCode, "Foo").Context;
            ParserRuleContext nullContext = ancestorIsNull ? aValidContext : null;
            ParserRuleContext ancestor = ancestorIsNull ? null : aValidContext;

            //Act
            var actual = nullContext.IsDescendentOf<ParserRuleContext>(ancestor);

            //Assert
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetAncestor_nullContext_returnsNull()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.GetAncestor<ParserRuleContext>();

            //Assert
            Assert.IsNull(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void TryGetAncestor_nullContext_returnsFalse()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.TryGetAncestor<ParserRuleContext>(out _);

            //Assert
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetAncestorContainingTokenIndex_nullContext_returnsNull()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.GetAncestorContainingTokenIndex(6);

            //Assert
            Assert.IsNull(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void ContainsTokenIndex_nullContext_returnsFalse()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.ContainsTokenIndex(6);

            //Assert
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetDescendent_nullContext_returnsNull()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.GetDescendent<ParserRuleContext>();

            //Assert
            Assert.IsNull(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetDescendents_nullContext_returnsEmptyList()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.GetDescendents<ParserRuleContext>().Count();

            //Assert
            Assert.AreEqual(0, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void TryGetChildContext_nullContext_returnsFalse()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = nullContext.TryGetChildContext<ParserRuleContext>(out _);

            //Assert
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetFirstEndOfLine_nullContext_returnsNull()
        {
            //Arrange
            VBAParser.EndOfStatementContext nullContext = null;

            //Act
            var actual = nullContext.GetFirstEndOfLine();

            //Assert
            Assert.IsNull(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void IsOptionCompareBinary_nullContext_throwsArgumentException()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            //Assert
            Assert.Throws<ArgumentException>(() => nullContext.IsOptionCompareBinary());
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetWidestDescendentContainingTokenIndex_nullContext_returnsNull()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = 
                nullContext.GetWidestDescendentContainingTokenIndex<ParserRuleContext>(6);

            //Assert
            Assert.IsNull(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetSmallestDescendentContainingTokenIndex_nullContext_returnsNull()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = 
                nullContext.GetSmallestDescendentContainingTokenIndex<ParserRuleContext>(6);

            //Assert
            Assert.IsNull(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetDescendentsContainingTokenIndex_nullContext_returnsEmptyList()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = 
                nullContext.GetDescendentsContainingTokenIndex<ParserRuleContext>(6)
                .Count();

            //Assert
            Assert.AreEqual(0, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetWidestDescendentContainingSelection_nullContext_returnsNull()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = 
                nullContext.GetWidestDescendentContainingSelection<ParserRuleContext>(new Selection());

            //Assert
            Assert.IsNull(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetSmallestDescendentContainingSelection_nullContext_returnsNull()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual = 
                nullContext.GetSmallestDescendentContainingSelection<ParserRuleContext>(new Selection());

            //Assert
            Assert.IsNull(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void GetDescendentsContainingSelection_nullContext_returnsEmptyList()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual =
                nullContext.GetDescendentsContainingSelection<ParserRuleContext>(new Selection())
                .Count();

            //Assert
            Assert.AreEqual(0, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void TryGetPrecedingContext_nullContext_returnsFalse()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual =
                nullContext.TryGetPrecedingContext<ParserRuleContext>(out _);

            //Assert
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void TryGetFollowingContext_nullContext_returnsFalse()
        {
            //Arrange
            ParserRuleContext nullContext = null;

            //Act
            var actual =
                nullContext.TryGetFollowingContext<ParserRuleContext>(out _);

            //Assert
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(ParserRuleContextExtensions))]
        public void ContainsExecutableStatements_nullContext_returnsFalse()
        {
            //Arrange
            VBAParser.BlockContext nullContext = null;

            //Act
            var actual =
                nullContext.ContainsExecutableStatements();

            //Assert
            Assert.IsFalse(actual);
        }

        private Declaration GetUserDeclaration(string inputCode, string identifier)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                return state.DeclarationFinder.MatchName(identifier).FirstOrDefault();
            }
        }
    }
}
