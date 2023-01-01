using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using static Rubberduck.Parsing.Grammar.VBAParser;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime;
using System.Collections.Generic;

namespace RubberduckTests.Grammar
{
    [TestFixture]
    public class SelectionExtensionsTests
    {
        public class CollectorVBAParserBaseVisitor<Result> : VBAParserBaseVisitor<IEnumerable<Result>>
        {
            protected override IEnumerable<Result> DefaultResult => new List<Result>();

            protected override IEnumerable<Result> AggregateResult(IEnumerable<Result> firstResult, IEnumerable<Result> secondResult)
            {
                return firstResult.Concat(secondResult);
            }
        }

        public class SubStmtContextElementCollectorVisitor : CollectorVBAParserBaseVisitor<SubStmtContext>
        {
            public override IEnumerable<SubStmtContext> VisitSubStmt([NotNull] SubStmtContext context)
            {
                return new List<SubStmtContext> { context };
            }
        }

        public class IfStmtContextElementCollectorVisitor : CollectorVBAParserBaseVisitor<IfStmtContext>
        {
            public override IEnumerable<IfStmtContext> VisitIfStmt([NotNull] IfStmtContext context)
            {
                return base.VisitIfStmt(context).Concat(new List<IfStmtContext> { context });
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new SubStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).First();
                var selection = new Selection(3, 0, 10, 5);

                Assert.IsFalse(selection.Contains(context));
                Assert.IsFalse(selection.IsContainedIn(context));
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new SubStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).First();
                var selection = new Selection(4, 1, 11, 8);

                Assert.IsTrue(selection.Contains(context));
                Assert.IsFalse(selection.IsContainedIn(context));
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void Context_Not_In_Selection_Start_OneBased_EvilCode()
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new SubStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).First();
                var selection = new Selection(5, 1, 11, 8);

                Assert.IsFalse(selection.Contains(context));
                Assert.IsFalse(selection.IsContainedIn(context));
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void Context_Not_In_Selection_End_OneBased_EvilCode()
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new SubStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).First();
                var selection = new Selection(4, 1, 10, 8);

                Assert.IsFalse(selection.Contains(context));
                Assert.IsTrue(selection.IsContainedIn(context));
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new SubStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).First();
                pane.Selection = new Selection(4, 1, 11, 6);

                Assert.IsTrue(context.GetSelection().Contains(pane.Selection));
                Assert.IsTrue(pane.Selection.IsContainedIn(context));
                Assert.IsTrue(pane.Selection.Contains(context));
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new SubStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).First();
                pane.Selection = new Selection(3, 0, 7, 7);

                Assert.IsFalse(context.GetSelection().Contains(pane.Selection));
                Assert.IsFalse(pane.Selection.IsContainedIn(context));
                Assert.IsFalse(pane.Selection.Contains(context));
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new SubStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).First();
                pane.Selection = new Selection(4, 1, 8, 8);

                Assert.IsTrue(context.GetSelection().Contains(pane.Selection));
                Assert.IsTrue(pane.Selection.IsContainedIn(context));
                Assert.IsTrue(pane.Selection.Contains(context));
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new SubStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).First();
                var selection = new Selection(4, 1, 8, 8);

                Assert.IsTrue(selection.Contains(context));
                Assert.IsTrue(selection.IsContainedIn(context));
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void Context_NotIn_Selection_StartTooLate_OneBased()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo()

Debug.Print ""foo""

End Sub : 'Lame comment!
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new SubStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).First();
                var selection = new Selection(4, 2, 8, 8);

                Assert.IsFalse(selection.Contains(context));
                Assert.IsTrue(selection.IsContainedIn(context));
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new SubStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).First();
                var selection = new Selection(4, 1, 8, 7);

                Assert.IsFalse(selection.Contains(context));
                Assert.IsTrue(selection.IsContainedIn(context));
            }
        }

        [Test]
        [Category("Grammar")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new IfStmtContextElementCollectorVisitor();
                var contexts = visitor.Visit(tree);
                var selection = new Selection(6, 1, 10, 7);

                Assert.IsTrue(selection.Contains(contexts.ElementAt(0)));   // first If block
                Assert.IsFalse(selection.Contains(contexts.ElementAt(1)));  // second If block
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new IfStmtContextElementCollectorVisitor();
                var contexts = visitor.Visit(tree);
                var selection = new Selection(6, 1, 10, 7);

                Assert.IsTrue(selection.Contains(contexts.ElementAt(0)));   // first If block
                Assert.IsFalse(selection.Contains(contexts.ElementAt(1)));  // second If block
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new IfStmtContextElementCollectorVisitor();
                var contexts = visitor.Visit(tree);
                var selection = new Selection(12, 1, 16, 7);

                Assert.IsFalse(selection.Contains(contexts.ElementAt(0)));  // first If block
                Assert.IsTrue(selection.Contains(contexts.ElementAt(1)));   // second If block
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new IfStmtContextElementCollectorVisitor();
                var contexts = visitor.Visit(tree);
                var token = contexts.ElementAt(1).Stop;
                var selection = new Selection(12, 1, 16, 7);

                Assert.IsTrue(selection.Contains(token));                   // last token in second If block
                Assert.IsFalse(selection.Contains(contexts.ElementAt(0)));  // first If block
                Assert.IsTrue(selection.Contains(contexts.ElementAt(1)));   // second If block
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new IfStmtContextElementCollectorVisitor();
                var context = visitor.Visit(tree).Last();
                var token = context.Stop;
                var selection = new Selection(12, 1, 14, 1);

                Assert.IsFalse(selection.Contains(token));
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void Selection_Contains_Only_Innermost_Nested_Context()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo(Bar As Long, Baz As Long, FooBar As Long)

If Bar > Baz Then
  Debug.Print ""Yeah!""
  If FooBar Then
     Debug.Print ""Foo bar!""
  End If
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new IfStmtContextElementCollectorVisitor();
                var contexts = visitor.Visit(tree);
                var token = contexts.ElementAt(0).Stop;
                var selection = new Selection(8, 1, 10, 9);

                Assert.IsTrue(selection.Contains(token));                   // last token in innermost If block
                Assert.IsTrue(selection.Contains(contexts.ElementAt(0)));   // innermost If block
                Assert.IsFalse(selection.Contains(contexts.ElementAt(1)));  // first outer If block
                Assert.IsFalse(selection.Contains(contexts.ElementAt(2)));  // second outer If block
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void Selection_Contains_Both_Nested_Context()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo(Bar As Long, Baz As Long, FooBar As Long)

If Bar > Baz Then
  Debug.Print ""Yeah!""
  If FooBar Then
     Debug.Print ""Foo bar!""
  End If
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new IfStmtContextElementCollectorVisitor();
                var contexts = visitor.Visit(tree); //returns innermost statement first then topmost consecutively
                var token = contexts.ElementAt(0).Stop;
                var selection = new Selection(6, 1, 13, 7);

                Assert.IsTrue(selection.Contains(token));                   // last token in innermost If block
                Assert.IsTrue(selection.Contains(contexts.ElementAt(0)));   // innermost If block
                Assert.IsTrue(selection.Contains(contexts.ElementAt(1)));   // first outer If block
                Assert.IsFalse(selection.Contains(contexts.ElementAt(2)));  // second outer If block
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void Selection_Not_Contained_In_Neither_Nested_Context()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo(Bar As Long, Baz As Long, FooBar As Long)

If Bar > Baz Then
  Debug.Print ""Yeah!""
  If FooBar Then
     Debug.Print ""Foo bar!""
  End If
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var tree = state.GetParseTree(new QualifiedModuleName(component));
                var visitor = new IfStmtContextElementCollectorVisitor();
                var contexts = visitor.Visit(tree); //returns innermost statement first then topmost consecutively
                var token = contexts.ElementAt(0).Stop;
                var selection = new Selection(15, 1, 19, 7);

                Assert.IsFalse(selection.Contains(token));                  // last token in innermost If block
                Assert.IsFalse(selection.Contains(contexts.ElementAt(0)));  // innermost If block
                Assert.IsFalse(selection.Contains(contexts.ElementAt(1)));  // first outer if block
                Assert.IsTrue(selection.Contains(contexts.ElementAt(2)));   // second outer If block
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void GivenOnlyBlankLines_EndColumn_Works()
        {
            const string inputCode = @"



";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {


                var tree = (ParserRuleContext)state.GetParseTree(new QualifiedModuleName(component));
                var startToken = tree.Start;
                var endToken = tree.Stop;

                // Reminder: token columns are zero-based but lines are one-based
                Assert.IsTrue(startToken.EndColumn() == 0);
                Assert.IsTrue(endToken.EndColumn() == 0);
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void GivenOnlyBlankLines_EndLine_Works()
        {
            const string inputCode = @"


";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {


                var tree = (ParserRuleContext)state.GetParseTree(new QualifiedModuleName(component));
                var startToken = tree.Start;
                var endToken = tree.Stop;

                // Reminder: token columns are zero-based but lines are one-based
                Assert.IsTrue(startToken.EndLine() == 1);
                Assert.IsTrue(endToken.EndLine() == 4);
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void GivenBlankLinesWithLeadingSpaces_EndColumn_Works()
        {
            const string inputCode = @"

   ";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {


                var tree = (ParserRuleContext)state.GetParseTree(new QualifiedModuleName(component));
                var startToken = tree.Start;
                var endToken = tree.Stop;

                // Reminder: token columns are zero-based but lines are one-based
                Assert.IsTrue(startToken.EndColumn() == 0);
                Assert.IsTrue(endToken.EndColumn() == 3);
            }
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void GivenBlankLinesWithLeadingSpaces_EndLine_Works()
        {
            const string inputCode = @"

   ";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var pane = component.CodeModule.CodePane;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {


                var tree = (ParserRuleContext)state.GetParseTree(new QualifiedModuleName(component));
                var startToken = tree.Start;
                var endToken = tree.Stop;

                // Reminder: token columns are zero-based but lines are one-based
                Assert.IsTrue(startToken.EndLine() == 1);
                Assert.IsTrue(endToken.EndLine() == 3);
            }
        }
        
        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        public void Selection_Overlaps_Other_Selection()
        {
            const string inputCode = @"
Option Explicit

Public Sub foo(Bar As Long, Baz As Long, FooBar As Long)

If Bar > Baz Then
  Debug.Print ""Yeah!""
  If FooBar Then
     Debug.Print ""Foo bar!""
  End If
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
            var contexts = visitor.Visit(tree); //returns innermost statement first then topmost consecutively
            var token = contexts.ElementAt(0).Stop;
            var selection = new Selection(10, 2, 15, 7);
            
            Assert.IsTrue(selection.Overlaps(contexts.ElementAt(0).GetSelection()));  // innermost If block
            Assert.IsTrue(selection.Overlaps(contexts.ElementAt(1).GetSelection()));  // first outer if block
            Assert.IsTrue(selection.Overlaps(contexts.ElementAt(2).GetSelection()));  // second outer If block

            selection = new Selection(2, 1, 4, 57);
            Assert.IsFalse(selection.Overlaps(contexts.ElementAt(0).GetSelection()));  // innermost If block
            Assert.IsFalse(selection.Overlaps(contexts.ElementAt(1).GetSelection()));  // first outer if block
            Assert.IsFalse(selection.Overlaps(contexts.ElementAt(2).GetSelection()));  // second outer If block

            selection = new Selection(8, 1, 10, 9);
            Assert.IsTrue(selection.Overlaps(contexts.ElementAt(0).GetSelection()));  // innermost If block
            Assert.IsTrue(selection.Overlaps(contexts.ElementAt(1).GetSelection()));  // first outer if block
            Assert.IsFalse(selection.Overlaps(contexts.ElementAt(2).GetSelection()));  // second outer If block

            selection = new Selection(8, 2, 10, 9);
            Assert.IsTrue(selection.Overlaps(contexts.ElementAt(0).GetSelection()));  // innermost If block

            selection = new Selection(8, 1, 10, 8);
            Assert.IsTrue(selection.Overlaps(contexts.ElementAt(0).GetSelection()));  // innermost If block

            selection = new Selection(8, 1, 8, 1);
            Assert.IsFalse(selection.Overlaps(contexts.ElementAt(0).GetSelection()));  // innermost If block

            selection = new Selection(8, 1, 8, 2);
            Assert.IsFalse(selection.Overlaps(contexts.ElementAt(0).GetSelection()));  // innermost If block
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        [TestCase(1, 1, 1, 10, Description = "Partial overlap at boundary")]
        [TestCase(1, 20, 1, 30, Description = "Partial overlap at boundary")]
        [TestCase(1, 1, 1, 30, Description = "Selection contained in new selection")]
        [TestCase(1, 15, 1, 15, Description = "New selection contained in selection")]
        public void Single_Line_Selection_Overlaps_Other_Selection(int startLine, int startColumn, int endLine, int endColumn)
        {
            var selection = new Selection(1, 10, 1, 20);
            var newSelection = new Selection(startLine, startColumn, endLine, endColumn);
            Assert.IsTrue(selection.Overlaps(newSelection));
        }

        [Test]
        [Category("Grammar")]
        [Category("Selection")]
        [TestCase(1, 1, 1, 9, Description = "New selection up to edge of current selection")]
        [TestCase(1, 21, 1, 30, Description = "New selection immediately after current selection")]
        public void Single_Line_Selection_Doesnt_Overlap_Other_Selection(int startLine, int startColumn, int endLine, int endColumn)
        {
            var selection = new Selection(1, 10, 1, 20);
            var newSelection = new Selection(startLine, startColumn, endLine, endColumn);
            Assert.IsFalse(selection.Overlaps(newSelection));
        }
    }
}
