using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Resources.Experimentals;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty 'Else' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Empty code blocks are redundant, dead code that should be removed. They can also be misleading about their intent:
    /// an empty block may be signalling an unfinished thought or an oversight.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Boolean)
    ///     If foo Then
    ///         ' ...
    ///     Else
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Boolean)
    ///     If foo Then
    ///         ' ...
    ///     End If
    /// End Sub
    /// ]]>
    /// </example>
    [Experimental(nameof(ExperimentalNames.EmptyBlockInspections))]
    internal class EmptyElseBlockInspection : ParseTreeInspectionBase
    {
        public EmptyElseBlockInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionResults.EmptyElseBlockInspection,
                                                        result));
        }

        public override IInspectionListener Listener { get; } 
            = new EmptyElseBlockListener();
        
        public class EmptyElseBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterElseBlock([NotNull] VBAParser.ElseBlockContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
