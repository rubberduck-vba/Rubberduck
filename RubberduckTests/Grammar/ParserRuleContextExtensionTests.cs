using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Grammar
{
    [TestFixture]
    public class ParserRuleContextExtensionTests
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


        [TestCase("selectCase3Arg", ExpectedResult = true)]
        [TestCase("firstArg", ExpectedResult = false)]
        [Category("Inspections")]
        public bool ParserRuleCtxtExt_HasParentType(string identifer)
        {
            var testArg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals(identifer)).First();
            var result = testArg.Context.IsDescendentOf<VBAParser.SelectCaseStmtContext>();
            return result;
        }

        [TestCase("selectCase3", ExpectedResult = true)]
        [TestCase("selectCase1", ExpectedResult = false)]
        [Category("Inspections")]
        public bool ParserRuleCtxtExt_HasParentOfSameType(string contextID)
        {
            bool result = false;

            var testIdDecs = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals(contextID));
            if (testIdDecs.Any())
            {
                var refs = testIdDecs.First().References;
                var testCtxt = (ParserRuleContext)refs.Where(rf => rf.Context.Parent.Parent.Parent is VBAParser.SelectCaseStmtContext).First().Context.Parent.Parent.Parent;
                result = testCtxt.IsDescendentOf<VBAParser.SelectCaseStmtContext>();
            }
            return result;
        }

        [TestCase("selectCase3", "selectCase1", ExpectedResult = true)]
        [TestCase("selectCase1", "selectCase3", ExpectedResult = false)]
        [TestCase("selectCase3", "selectCase3", ExpectedResult = false)]
        [Category("Inspections")]
        public bool ParserRuleCtxtExt_HasParentContext(string contextID, string parentContextID)
        {
            bool result = false;
            var parentIdDec = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals(parentContextID)).First();
            var parentCtxt = (VBAParser.SelectCaseStmtContext)parentIdDec.References.Where(rf => rf.Context.Parent.Parent.Parent is VBAParser.SelectCaseStmtContext).First().Context.Parent.Parent.Parent;

            var testIdDecs = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals(contextID));
            if (testIdDecs.Any())
            {
                var refs = testIdDecs.First().References;
                var testCtxt = (ParserRuleContext)refs.Where(rf => rf.Context.Parent.Parent.Parent is VBAParser.SelectCaseStmtContext).First().Context.Parent.Parent.Parent;
                result = testCtxt.IsDescendentOf(parentCtxt);
            }
            return result;
        }


        [Test]
        [Category("Inspections")]
        public void ParserRuleCtxtExt_HasParent_ByType_False()
        {
            var selectCase3Arg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase3Arg")).First();
            var result = selectCase3Arg.Context.IsDescendentOf<VBAParser.SubStmtContext>();
            Assert.AreEqual(false, result);
        }

        [Test]
        [Category("Inspections")]
        public void ParserRuleCtxtExt_HasParent_ByContext_True()
        {
            var selectCase3Arg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase3Arg")).First();
            var functContext = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("Foo")).First().Context;
            var result = selectCase3Arg.Context.IsDescendentOf(functContext);
            Assert.AreEqual(true, result);
        }

        [Test]
        [Category("Inspections")]
        public void ParserRuleCtxtExt_HasParent_ByContext_False()
        {
            var selectCase3Arg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase3Arg")).First();
            var functContext = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase1")).First().Context;
            var result = selectCase3Arg.Context.IsDescendentOf(functContext);
            Assert.AreEqual(false, result);
        }

        [Test]
        [Category("Inspections")]
        public void ParserRuleCtxtExt_GetChild_NullResult()
        {
            var selectCase3Arg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase3Arg")).First();
            var result = selectCase3Arg.Context.GetChild<VBAParser.SelectCaseStmtContext>();
            Assert.AreEqual(null, result);
        }

        [Test]
        [Category("Inspections")]
        [Category("Grammar")]
        public void ParserRuleCtxtExt_Evil_Code_Selection_Not_Evil()
        {
            const string inputCode =
                @" _
 _
 Function _
 _
 Foo _
 _
 ( _
 _
 fizz _
 _
 ) As Boolean

 End _
 _
 Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var declaration = state.AllDeclarations.Single(d => d.IdentifierName.Equals("Foo"));

                var actual = ((VBAParser.FunctionStmtContext)declaration.Context).GetProcedureSelection();
                var expected = new Selection(3, 2, 11, 14);

                Assert.AreEqual(actual, expected);
            }
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
