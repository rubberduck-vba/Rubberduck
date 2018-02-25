using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
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
        public bool ParserRuleCtxtExt_IsDescendentOf_ByContext(string contextID, string parentContextID)
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
        public void ParserRuleCtxtExt_IsDescendentOf_ByType_False()
        {
            var selectCase3Arg = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("selectCase3Arg")).First();
            var result = selectCase3Arg.Context.IsDescendentOf<VBAParser.SubStmtContext>();
            Assert.AreEqual(false, result);
        }

        [TestCase("Foo", PRCExtensionTestContextTypes.SelectStmtCtxt, ExpectedResult = 3)]
        [TestCase("Foo", PRCExtensionTestContextTypes.PowOpCtxt, ExpectedResult = 0)]
        [Category("Inspections")]
        public int ParserRuleCtxtExt_GetDescendents(string parentName, PRCExtensionTestContextTypes descendentType)
        {
            var parentContext = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals("Foo")).First().Context;
            var descendents = new List<ParserRuleContext>();
            if (descendentType == PRCExtensionTestContextTypes.SelectStmtCtxt)
            {
                descendents = parentContext.GetDescendents<VBAParser.SelectCaseStmtContext>().Select(desc => (ParserRuleContext)desc).ToList();
            }
            if (descendentType == PRCExtensionTestContextTypes.PowOpCtxt)
            {
                descendents = parentContext.GetDescendents<VBAParser.PowOpContext>().Select(desc => (ParserRuleContext)desc).ToList();
            }
            return descendents.Count();
        }

        [TestCase("Foo", PRCExtensionTestContextTypes.SelectStmtCtxt, ExpectedResult = true)]
        [TestCase("Foo", PRCExtensionTestContextTypes.PowOpCtxt, ExpectedResult = false)]
        [Category("Inspections")]
        public bool ParserRuleCtxtExt_GetDescendent(string parentName, PRCExtensionTestContextTypes descendentType)
        {
            var parentContext = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals(parentName)).First().Context;
            ParserRuleContext descendent = null;
            if (descendentType == PRCExtensionTestContextTypes.SelectStmtCtxt)
            {
                descendent = parentContext.GetDescendent<VBAParser.SelectCaseStmtContext>();
            }
            if (descendentType == PRCExtensionTestContextTypes.PowOpCtxt)
            {
                descendent = parentContext.GetDescendent<VBAParser.PowOpContext>();
            }
            return descendent != null;
        }

        [TestCase("selectCase3Arg", PRCExtensionTestContextTypes.SelectStmtCtxt, ExpectedResult = true)]
        [TestCase("selectCase3Arg", PRCExtensionTestContextTypes.PowOpCtxt, ExpectedResult = false)]
        [Category("Inspections")]
        public bool ParserRuleCtxtExt_GetAncestor(string name, PRCExtensionTestContextTypes ancestorType)
        {
            var context = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals(name)).First().Context;
            ParserRuleContext ancestor = null;
            if (ancestorType == PRCExtensionTestContextTypes.SelectStmtCtxt)
            {
                ancestor = context.GetAncestor<VBAParser.SelectCaseStmtContext>();
            }
            if (ancestorType == PRCExtensionTestContextTypes.PowOpCtxt)
            {
                ancestor = context.GetAncestor<VBAParser.PowOpContext>();
            }
            return ancestor != null;
        }

        [TestCase("selectCase3Arg", "Foo", ExpectedResult = true)]
        [TestCase("selectCase3Arg", "selectCase1", ExpectedResult = false)]
        [Category("Inspections")]
        public bool ParserRuleCtxtExt_IsDescendentOf_ByContext2(string contextName, string parentName)
        {
            var descendentCandidate = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals(contextName)).First().Context;
            var parentCandidate = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals(parentName)).First().Context;
            var result = descendentCandidate.IsDescendentOf(parentCandidate);
            return result;
        }

        public enum PRCExtensionTestContextTypes {SelectStmtCtxt, AsTypeCtxt, PowOpCtxt };

        [TestCase("selectCase3Arg", PRCExtensionTestContextTypes.SelectStmtCtxt, ExpectedResult = false)]
        [TestCase("selectCase3Arg", PRCExtensionTestContextTypes.AsTypeCtxt, ExpectedResult = true)]
        [Category("Inspections")]
        public bool ParserRuleCtxtExt_GetChild(string parentContextName, PRCExtensionTestContextTypes ctxtType)
        {
            ParserRuleContext result = null;
            var parentContext = AllTestingDeclarations.Where(dc => dc.IdentifierName.Equals(parentContextName)).First().Context;
            if (ctxtType == PRCExtensionTestContextTypes.SelectStmtCtxt)
            {
                result = parentContext.GetChild<VBAParser.SelectCaseStmtContext>();
            }
            else if (ctxtType == PRCExtensionTestContextTypes.AsTypeCtxt)
            {
                result = parentContext.GetChild<VBAParser.AsTypeClauseContext>();
            }
            return result != null;
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
