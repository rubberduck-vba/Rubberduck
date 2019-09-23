using Antlr4.Runtime.Tree;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.CodePathAnalysis.Execution;
using Rubberduck.CodeAnalysis.CodePathAnalysis.Execution.ExtendedNodeVisitor;
using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.CodeAnalysis.CodePaths
{
    [TestFixture]
    public class ExecutionEngineTests
    {
        [Test][Category("CodePathAnalysis")]
        public void SingleCodePathInputYieldsSingleCodePath()
        {
            const string inputCode = @"Option Explicit
Public Sub DoSomething()
    MsgBox ""hello""
End Sub
";
            var paths = GetCodePaths(inputCode);
            Assert.AreEqual(1, paths.Count());
        }

        [Test][Category("CodePathAnalysis")]
        public void BranchingCodePathsInputYieldsTwoCodePaths()
        {
            const string inputCode = @"Option Explicit
Public Sub DoSomething()
    Debug.Print ""hi from path 1""
    If True Then
        MsgBox ""hello from path 2""
    End If
    Debug.Print ""still in path 1""
End Sub
";
            var paths = GetCodePaths(inputCode);
            Assert.AreEqual(2, paths.Count());
        }

        [Test][Category("CodePathAnalysis")]
        public void BranchingCodePathsInput_CodeAfterBranchIsStillInFirstPath()
        {
            const string inputCode = @"Option Explicit
Public Sub DoSomething()
    Debug.Print ""hi from path 1""
    If True Then
        MsgBox ""hello from path 2""
    End If
    Debug.Print ""still in path 1""
End Sub
";
            var path = GetCodePaths(inputCode).FirstOrDefault();
            if (path?.Count == 0) { Assert.Inconclusive("CodePath is empty"); }

            Assert.IsTrue(typeof(IExecutableNode).IsAssignableFrom(path[0].GetType()), "Expecting IExecutableNode at index 0.");
            Assert.IsTrue(typeof(IBranchNode).IsAssignableFrom(path[1].GetType()), "Expecting IBranchNode at index 1.");
            Assert.IsTrue(typeof(IEvaluatableNode).IsAssignableFrom(path[2].GetType()), "Expecting IBranchNode at index 2.");
            Assert.IsTrue(typeof(IExecutableNode).IsAssignableFrom(path[3].GetType()), "Expecting IExecutableNode at index 3.");
        }

        private IEnumerable<CodePath> GetCodePaths(string inputCode)
        {
            using (var state = MockParser.ParseString(inputCode, out var qmn))
            {
                var result = new List<CodePath>();
                foreach (var member in state.DeclarationFinder.Members(qmn))
                {
                    if (member is ModuleBodyElementDeclaration element)
                    {
                        var visitor = new ExtendedNodeVisitor(element);
                        result.AddRange(visitor.GetAllCodePaths());
                    }
                }
                return result;
            }
        }
    }
}
