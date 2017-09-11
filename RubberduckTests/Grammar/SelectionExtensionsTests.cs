using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using static Rubberduck.Parsing.Grammar.VBAParser;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using Antlr4.Runtime.Misc;
using System.Collections.Generic;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class SelectionExtensionsTests
    {
        public class CollectorVBAParserBaseVisitor<Result> : VBAParserBaseVisitor<IList<Result>>
        {
            private List<Result> defaultResult = new List<Result>();

            protected override IList<Result> DefaultResult => defaultResult;
            /*
            protected override IList<Result> AggregateResult(IList<Result> firstResult, IList<Result> secondResult)
            {
                if (firstResult != null && secondResult != null)
                    return firstResult.Concat(secondResult) as IList<Result>;

                if (secondResult == null)
                    return firstResult;

                return secondResult;
            }*/
        }

        public class SubStmtContextElementCollectorVisitor : CollectorVBAParserBaseVisitor<SubStmtContext>
        {
            public override IList<SubStmtContext> VisitSubStmt([NotNull] SubStmtContext context)
            {
                DefaultResult.Add(context);
                return base.VisitSubStmt(context);
            }
        }

        public class IfStmtContextElementCollectorVisitor : CollectorVBAParserBaseVisitor<IfStmtContext>
        {
            public override IList<IfStmtContext> VisitIfStmt([NotNull] IfStmtContext context)
            {
                DefaultResult.Add(context);
                return base.VisitIfStmt(context);
            }
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_Not_In_Selection_ZeroBased_EvilCode()
        {
            const string inputCode = @"
Option Explicit

Public _
    Sub _
foo()

Debug.Print ""foo""

    End _
  Sub : 'Lame comment!
";
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new SubStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).First();
            var selection = new Selection(3, 0, 10, 5);

            Assert.IsFalse(context.Contains(selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_In_Selection_OneBased_EvilCode()
        {
            const string inputCode = @"
Option Explicit

Public _
    Sub _
foo()

Debug.Print ""foo""

    End _
  Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new SubStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).First();
            var selection = new Selection(4, 1, 11, 8);
            
            Assert.IsTrue(context.Contains(selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_In_GetSelection_OneBased_EvilCode()
        {
            const string inputCode = @"
Option Explicit

Public _
    Sub _
foo()

Debug.Print ""foo""

    End _
  Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new SubStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).First();
            pane.Selection = new Selection(4, 1, 11, 8);
            
            Assert.IsTrue(context.Contains(pane.Selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_Not_In_GetSelection_ZeroBased()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo()

Debug.Print ""foo""

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new SubStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).First();
            pane.Selection = new Selection(3, 0, 7, 7);

            Assert.IsFalse(context.Contains(pane.Selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_In_GetSelection_OneBased()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo()

Debug.Print ""foo""

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new SubStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).First();
            pane.Selection = new Selection(4, 1, 8, 8);

            Assert.IsTrue(context.Contains(pane.Selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_In_Selection_OneBased()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo()

Debug.Print ""foo""

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new SubStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).First();
            var selection = new Selection(4, 1, 8, 8);

            Assert.IsTrue(context.Contains(selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_NotIn_Selection_StartTooSoon_OneBased()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo()

Debug.Print ""foo""

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new SubStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).First();
            var selection = new Selection(4, 2, 8, 8);

            Assert.IsFalse(context.Contains(selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_NotIn_Selection_EndsTooSoon_OneBased()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo()

Debug.Print ""foo""

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new SubStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).First();
            var selection = new Selection(4, 1, 8, 7);

            Assert.IsFalse(context.Contains(selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_In_Selection_FirstBlock_OneBased()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo(Bar As Long, Baz As Long)

If Bar > Baz Then
  Debug.Print ""Yeah!""
Else
  Debug.Print ""Boo!""
End If

If Baz > Bar Then
  Debug.Print ""Boo!""
Else
  Debug.Print ""Yeah!""
End If

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new IfStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).First();
            var selection = new Selection(6, 1, 10, 7);

            Assert.IsTrue(context.Contains(selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_Not_In_Selection_SecondBlock_OneBased()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo(Bar As Long, Baz As Long)

If Bar > Baz Then
  Debug.Print ""Yeah!""
Else
  Debug.Print ""Boo!""
End If

If Baz > Bar Then
  Debug.Print ""Boo!""
Else
  Debug.Print ""Yeah!""
End If

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new IfStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).Last();
            var selection = new Selection(6, 1, 10, 7);

            Assert.IsFalse(context.Contains(selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Context_In_Selection_SecondBlock_OneBased()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo(Bar As Long, Baz As Long)

If Bar > Baz Then
  Debug.Print ""Yeah!""
Else
  Debug.Print ""Boo!""
End If

If Baz > Bar Then
  Debug.Print ""Boo!""
Else
  Debug.Print ""Yeah!""
End If

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new IfStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).Last();
            var selection = new Selection(12, 1, 16, 7);

            Assert.IsTrue(context.Contains(selection));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Selection_Contains_LastToken()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo(Bar As Long, Baz As Long)

If Bar > Baz Then
  Debug.Print ""Yeah!""
Else
  Debug.Print ""Boo!""
End If

If Baz > Bar Then
  Debug.Print ""Boo!""
Else
  Debug.Print ""Yeah!""
End If

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new IfStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).Last();
            var token = context.Stop;
            var selection = new Selection(12, 1, 16, 7);

            Assert.IsTrue(selection.Contains(token));
        }

        [TestMethod]
        [TestCategory("Grammar")]
        public void Selection_Not_Contains_LastToken()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo(Bar As Long, Baz As Long)

If Bar > Baz Then
  Debug.Print ""Yeah!""
Else
  Debug.Print ""Boo!""
End If

If Baz > Bar Then
  Debug.Print ""Boo!""
Else
  Debug.Print ""Yeah!""
End If

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            var state = MockParser.CreateAndParse(vbe.Object);
            var tree = state.GetParseTree(new QualifiedModuleName(component));
            var visitor = new IfStmtContextElementCollectorVisitor();
            var context = visitor.Visit(tree).Last();
            var token = context.Stop;
            var selection = new Selection(12, 1, 14, 1);

            Assert.IsFalse(selection.Contains(token));
        }
    }
}
