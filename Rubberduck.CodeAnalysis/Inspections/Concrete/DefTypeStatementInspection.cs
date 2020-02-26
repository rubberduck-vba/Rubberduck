using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about Def[Type] statements.
    /// </summary>
    /// <why>
    /// These declarative statements make the first letter of identifiers determine the data type.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// DefBool B
    /// DefDbl D
    ///
    /// Public Sub DoSomething() 
    ///     Dim bar ' implicit Boolean
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class DefTypeStatementInspection : ParseTreeInspectionBase
    {
        public DefTypeStatementInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Listener = new DefTypeStatementInspectionListener();
        }
        
        public override IInspectionListener Listener { get; }
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            var typeName = GetTypeOfDefType(context.Context.start.Text);
            var defStmtText = context.Context.start.Text;

            return string.Format(
                InspectionResults.DefTypeStatementInspection,
                typeName,
                defStmtText);
        }

        private string GetTypeOfDefType(string defType)
        {
            _defTypes.TryGetValue(defType, out var value);
            return value;
        }

        private readonly Dictionary<string, string> _defTypes = new Dictionary<string, string>
        {
            { "DefBool", "Boolean" },
            { "DefByte", "Byte" },
            { "DefInt", "Integer" },
            { "DefLng", "Long" },
            { "DefCur", "Currency" },
            { "DefSng", "Single" },
            { "DefDbl", "Double" },
            { "DefDate", "Date" },
            { "DefStr", "String" },
            { "DefObj", "Object" },
            { "DefVar", "Variant" }
        };

        public class DefTypeStatementInspectionListener : InspectionListenerBase
        {
            public override void ExitDefType([NotNull] VBAParser.DefTypeContext context)
            {
                SaveContext(context);
            }
        }
    }
}