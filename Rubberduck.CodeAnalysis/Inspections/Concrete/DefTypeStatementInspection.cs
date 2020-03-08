using System.Collections.Generic;
using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about Def[Type] statements.
    /// </summary>
    /// <why>
    /// These declarative statements make the first letter of identifiers determine the data type.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// DefBool B
    /// DefDbl D
    ///
    /// Public Sub DoSomething() 
    ///     Dim bar ' implicit Boolean
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class DefTypeStatementInspection : ParseTreeInspectionBase<VBAParser.DefTypeContext>
    {
        public DefTypeStatementInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new DefTypeStatementInspectionListener();
        }
        
        protected override IInspectionListener<VBAParser.DefTypeContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.DefTypeContext> context)
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

        private class DefTypeStatementInspectionListener : InspectionListenerBase<VBAParser.DefTypeContext>
        {
            public override void ExitDefType([NotNull] VBAParser.DefTypeContext context)
            {
                SaveContext(context);
            }
        }
    }
}