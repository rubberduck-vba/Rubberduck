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
    public class ImplicitByRefSubParameterInspection : IInspection
    {
        public ImplicitByRefSubParameterInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ImplicitByRef; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult) 
        {
            foreach (var module in parseResult)
            {
                var procedures = module.ParseTree.GetContexts<ProcedureListener, ParserRuleContext>(new ProcedureListener());
                foreach (var procedure in procedures)
                {
                    var args = GetArguments(procedure);
                    foreach (var arg in args.Where(arg => arg.BYREF() == null && arg.BYVAL() == null))
                    {
                        var context = new QualifiedContext<VisualBasic6Parser.ArgContext>(module.QualifiedName, arg);
                        yield return new ImplicitByRefParameterInspectionResult(Name, Severity, context);
                    }
                }
            }
        }

        private static readonly IEnumerable<Func<ParserRuleContext, VisualBasic6Parser.ArgListContext>> Converters =
            new List<Func<ParserRuleContext, VisualBasic6Parser.ArgListContext>>
            {
                GetSubArgsList,
                GetFunctionArgsList,
                GetPropertyGetArgsList,
                GetPropertyLetArgsList,
                GetPropertySetArgsList
            };

        private IEnumerable<VisualBasic6Parser.ArgContext> GetArguments(ParserRuleContext procedureContext)
        {
            var argsList = Converters.Select(converter => converter(procedureContext)).FirstOrDefault(args => args != null);
            if (argsList == null)
            {
                return new List<VisualBasic6Parser.ArgContext>();
            }

            return argsList.arg();
        }

        private static VisualBasic6Parser.ArgListContext GetSubArgsList(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VisualBasic6Parser.SubStmtContext;
            return context == null ? null : context.argList();
        }

        private static VisualBasic6Parser.ArgListContext GetFunctionArgsList(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VisualBasic6Parser.FunctionStmtContext;
            return context == null ? null : context.argList();
        }

        private static VisualBasic6Parser.ArgListContext GetPropertyGetArgsList(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VisualBasic6Parser.PropertyGetStmtContext;
            return context == null ? null : context.argList();
        }

        private static VisualBasic6Parser.ArgListContext GetPropertyLetArgsList(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VisualBasic6Parser.PropertyLetStmtContext;
            return context == null ? null : context.argList();
        }

        private static VisualBasic6Parser.ArgListContext GetPropertySetArgsList(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VisualBasic6Parser.PropertySetStmtContext;
            return context == null ? null : context.argList();
        }
    }
}