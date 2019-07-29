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

namespace Rubberduck.Inspections.Concrete
{
    [Experimental]
    internal class EmptyMethodInspection : ParseTreeInspectionBase
    {
        public EmptyMethodInspection(RubberduckParserState state)
            : base(state) { }

        public override IInspectionListener Listener { get; } =
            new EmptyMethodListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this, "Test string", result));
        }
    }

    internal class EmptyMethodListener : EmptyBlockInspectionListenerBase
    {
        public override void EnterFunctionStmt([NotNull] VBAParser.FunctionStmtContext context)
        {
            InspectBlockForExecutableStatements(context.block(), context);
        }

        public override void EnterPropertyGetStmt([NotNull] VBAParser.PropertyGetStmtContext context)
        {
            InspectBlockForExecutableStatements(context.block(), context);
        }

        public override void EnterPropertyLetStmt([NotNull] VBAParser.PropertyLetStmtContext context)
        {
            InspectBlockForExecutableStatements(context.block(), context);
        }

        public override void EnterPropertySetStmt([NotNull] VBAParser.PropertySetStmtContext context)
        {
            InspectBlockForExecutableStatements(context.block(), context);
        }

        public override void EnterSubStmt([NotNull] VBAParser.SubStmtContext context)
        {
            InspectBlockForExecutableStatements(context.block(), context);
        }
    }
}
