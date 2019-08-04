using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using static Rubberduck.Parsing.Grammar.VBAParser;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Resources.Experimentals;

namespace Rubberduck.Inspections.Concrete
{

    /// <summary>
    /// Identifies empty module member blocks.
    /// </summary>
    /// <why>
    /// Methods containing no executable statements are misleading as they appear to be doing something which they actually don't.
    /// This might be the result of delaying the actual implementation for a later stage of development, and then forgetting all about that.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Sub Foo()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Sub Foo()
    ///     MsgBox "?"
    /// End Sub
    /// ]]>
    /// </example>
    [Experimental(nameof(ExperimentalNames.EmptyBlockInspections))]
    internal class EmptyMethodInspection : ParseTreeInspectionBase
    {
        public EmptyMethodInspection(RubberduckParserState state)
            : base(state) { }

        public override IInspectionListener Listener { get; } =
            new EmptyMethodListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            // Exclude empty members in user interfaces, as long as all members of the interface are empty,
            // since some VB users might use concrete user defined classes as interfaces,
            // while RD marks them as interfaces all the same.
            
            var results = Listener.Contexts
                .Where(result => !result.IsIgnoringInspectionResultFor(State.DeclarationFinder, AnnotationName))
                .GroupBy(result => result.ModuleName.ComponentName)
                // Exclude results from module
                .Where(resultsInModule => !State.DeclarationFinder.FindAllUserInterfaces()
                                          // where all members of that module contain no executables
                                          .Where(interfaceModule => interfaceModule.ComponentName == resultsInModule.Key
                                                                    && interfaceModule.Members.Count == resultsInModule.Count())
                                          .Any())
                .SelectMany(resultsInModule => resultsInModule)
                .Select(result => new { actual = result, method = result.Context as IMethodStmtContext });

            return results.Select(result => new QualifiedContextInspectionResult(this,
                                                                                 string.Format(InspectionResults.EmptyMethodInspection,
                                                                                              result.method.MethodKind,
                                                                                              result.method.MethodName),
                                                                                result.actual));
        }
    }

    internal class EmptyMethodListener : EmptyBlockInspectionListenerBase
    {
        public override void EnterFunctionStmt([NotNull] FunctionStmtContext context)
        {
            InspectBlockForExecutableStatements(context.block(), context);
        }

        public override void EnterPropertyGetStmt([NotNull] PropertyGetStmtContext context)
        { 
            InspectBlockForExecutableStatements(context.block(), context);
        }

        public override void EnterPropertyLetStmt([NotNull] PropertyLetStmtContext context)
        {
            InspectBlockForExecutableStatements(context.block(), context);
        }

        public override void EnterPropertySetStmt([NotNull] PropertySetStmtContext context)
        {
            InspectBlockForExecutableStatements(context.block(), context);
        }

        public override void EnterSubStmt([NotNull] SubStmtContext context)
        {
            InspectBlockForExecutableStatements(context.block(), context);
        }

    }
}
