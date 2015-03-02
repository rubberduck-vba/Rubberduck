using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    public class ImplicitByRefParameterInspection : IInspection
    {
        public ImplicitByRefParameterInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ImplicitByRef_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult) 
        {
            foreach (var module in parseResult.ComponentParseResults)
            {
                var procedures = module.ParseTree.GetContexts<ProcedureListener, ParserRuleContext>(new ProcedureListener(module.QualifiedName));
                foreach (var procedure in procedures)
                {
                    var args = GetArguments(procedure);
                    foreach (var arg in args.Where(arg => arg.BYREF() == null && arg.BYVAL() == null && arg.PARAMARRAY() == null))
                    {
                        var context = new QualifiedContext<VBParser.ArgContext>(module.QualifiedName, arg);
                        yield return new ImplicitByRefParameterInspectionResult(string.Format(Name, arg.AmbiguousIdentifier().GetText()), Severity, context);
                    }
                }
            }
        }

        private static readonly IEnumerable<Func<ParserRuleContext, VBParser.ArgListContext>> Converters =
            new List<Func<ParserRuleContext, VBParser.ArgListContext>>
            {
                GetSubArgsList,
                GetFunctionArgsList,
                GetPropertyGetArgsList,
                GetPropertyLetArgsList,
                GetPropertySetArgsList
            };

        private IEnumerable<VBParser.ArgContext> GetArguments(QualifiedContext<ParserRuleContext> procedureContext)
        {
            var argsList = Converters.Select(converter => converter(procedureContext.Context)).FirstOrDefault(args => args != null);
            if (argsList == null)
            {
                return new List<VBParser.ArgContext>();
            }

            return argsList.Arg();
        }

        private static VBParser.ArgListContext GetSubArgsList(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VBParser.SubStmtContext;
            return context == null ? null : context.ArgList();
        }

        private static VBParser.ArgListContext GetFunctionArgsList(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VBParser.FunctionStmtContext;
            return context == null ? null : context.ArgList();
        }

        private static VBParser.ArgListContext GetPropertyGetArgsList(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VBParser.PropertyGetStmtContext;
            return context == null ? null : context.ArgList();
        }

        private static VBParser.ArgListContext GetPropertyLetArgsList(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VBParser.PropertyLetStmtContext;
            return context == null ? null : context.ArgList();
        }

        private static VBParser.ArgListContext GetPropertySetArgsList(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VBParser.PropertySetStmtContext;
            return context == null ? null : context.ArgList();
        }
    }
}