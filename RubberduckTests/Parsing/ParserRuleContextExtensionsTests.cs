using Antlr4.Runtime;
using Moq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace RubberduckTests.Parsing
{
    [TestFixture]
    public class ParserRuleContextExtensionsTests
    {

        private List<Declaration> _allTestingDeclarations;
        private List<Declaration> AllTestingDeclarations
        {
            get
            {
                if (_allTestingDeclarations != null)
                {
                    return _allTestingDeclarations;
                }

                const string inputCode =
@"Public Function Foo(selectCase1 As Long, selectCase2 As Long, selectCase3 As Long) As Long
    Dim firstArg As Long
    firstArg = 5
    Select Case selectCase1
        Case 8
            firstArg = selectCase1 * 2
        Case 10
            firstArg = selectCase1 / 2
        Case Else
            Select Case selectCase2
                Case 8
                    firstArg = selectCase2 * 2
                Case 10
                    firstArg = selectCase2 / 2
                Case Else
                    Select Case selectCase3
                        Case 8
                            firstArg = selectCase3 * 2
                        Case 10
                            Dim selectCase3Arg As Long
                            selectCase3Arg = selectCase3 / 2
                    End Select
             End Select
     End Select
    Foo = firstArg
End Function";
                _allTestingDeclarations = GetAllUserDeclarations(inputCode).ToList();
                return _allTestingDeclarations;
            }
        }


        [Test]
        [Category("Inspections")]
        public void ParserRuleContextExtension_HasParent_Null()
        {
            var result = AllTestingDeclarations.First().Context.HasParent(null);
            Assert.AreEqual(false, result);
        }

        [Test]
        [Category("Inspections")]
        public void ParserRuleContextExtension_HasParent_ByType_True()
        {
            var selectCase3Arg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase3Arg")).First();
            var result = selectCase3Arg.Context.HasParent<VBAParser.FunctionStmtContext>();
            Assert.AreEqual(true, result);
        }

        [Test]
        [Category("Inspections")]
        public void ParserRuleContextExtension_HasParent_ByType_False()
        {
            var selectCase3Arg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase3Arg")).First();
            var result = selectCase3Arg.Context.HasParent<VBAParser.SubStmtContext>();
            Assert.AreEqual(false, result);
        }

        [Test]
        [Category("Inspections")]
        public void ParserRuleContextExtension_HasParent_ByContext_True()
        {
            var selectCase3Arg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase3Arg")).First();
            var functContext = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("Foo")).First().Context;
            var result = selectCase3Arg.Context.HasParent(functContext);
            Assert.AreEqual(true, result);
        }

        [Test]
        [Category("Inspections")]
        public void ParserRuleContextExtension_HasParent_ByContext_False()
        {
            var selectCase3Arg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase3Arg")).First();
            var functContext = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase1")).First().Context;
            var result = selectCase3Arg.Context.HasParent(functContext);
            Assert.AreEqual(false, result);
        }

        [Test]
        [Category("Inspections")]
        public void ParserRuleContextExtension_GetChild_NullResult()
        {
            var selectCase3Arg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase3Arg")).First();
            var result = selectCase3Arg.Context.GetChild<VBAParser.SelectCaseStmtContext>();
            Assert.AreEqual(null, result);
        }

        private IEnumerable<Declaration> GetAllUserDeclarations(string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                return state.AllUserDeclarations;
            }
        }
    }
}
