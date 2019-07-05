using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Inspections.Concrete.ThunderCode
{
    /// <summary hidden="true">
    /// A ThunderCode inspection that locates certain specific instances of line continuations in places we'd never think to put them.
    /// </summary>
    /// <why>
    /// This inpection is flagging code we dubbed "ThunderCode", 
    /// code our friend Andrew Jackson would have written to confuse Rubberduck's parser and/or resolver. 
    /// While perfectly legal, these line continuations serve no purpose and should be removed.
    /// </why>
    /// <remarks>
    /// Note that the inspection only checks a subset of possible "evil" line continatuions 
    /// for both simplicity and performance reasons. Exhaustive inspection would likely take too much effort. 
    /// </remarks>
    public class LineContinuationBetweenKeywordsInspection : ParseTreeInspectionBase
    {
        public LineContinuationBetweenKeywordsInspection(RubberduckParserState state) : base(state)
        {
            Listener = new LineContinuationBetweenKeywordsListener();
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts.Select(c => new QualifiedContextInspectionResult(
                this, InspectionResults.LineContinuationBetweenKeywordsInspection.ThunderCodeFormat(), c));
        }

        public override IInspectionListener Listener { get; }

        public class LineContinuationBetweenKeywordsListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();

            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public void ClearContexts() => _contexts.Clear();

            public QualifiedModuleName CurrentModuleName { get; set; }

            public override void EnterSubStmt(VBAParser.SubStmtContext context)
            {
                CheckContext(context, context.END_SUB());
                base.EnterSubStmt(context);
            }

            public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
            {
                CheckContext(context, context.END_FUNCTION());
                base.EnterFunctionStmt(context);
            }

            public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
            {
                CheckContext(context, context.PROPERTY_GET());
                CheckContext(context, context.END_PROPERTY());
                base.EnterPropertyGetStmt(context);
            }

            public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
            {
                CheckContext(context, context.PROPERTY_LET());
                CheckContext(context, context.END_PROPERTY());
                base.EnterPropertyLetStmt(context);
            }

            public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
            {
                CheckContext(context, context.PROPERTY_SET());
                CheckContext(context, context.END_PROPERTY());
                base.EnterPropertySetStmt(context);
            }

            public override void EnterSelectCaseStmt(VBAParser.SelectCaseStmtContext context)
            {
                CheckContext(context, context.END_SELECT());
                base.EnterSelectCaseStmt(context);
            }

            public override void EnterWithStmt(VBAParser.WithStmtContext context)
            {
                CheckContext(context, context.END_WITH());
                base.EnterWithStmt(context);
            }

            public override void EnterExitStmt(VBAParser.ExitStmtContext context)
            {
                CheckContext(context, context.EXIT_DO());
                CheckContext(context, context.EXIT_FOR());
                CheckContext(context, context.EXIT_FUNCTION());
                CheckContext(context, context.EXIT_PROPERTY());
                CheckContext(context, context.EXIT_SUB());
                base.EnterExitStmt(context);
            }

            public override void EnterOnErrorStmt(VBAParser.OnErrorStmtContext context)
            {
                CheckContext(context, context.ON_ERROR());
                CheckContext(context, context.ON_LOCAL_ERROR());
                base.EnterOnErrorStmt(context);
            }

            public override void EnterOptionBaseStmt(VBAParser.OptionBaseStmtContext context)
            {
                CheckContext(context, context.OPTION_BASE());
                base.EnterOptionBaseStmt(context);
            }

            public override void EnterOptionCompareStmt(VBAParser.OptionCompareStmtContext context)
            {
                CheckContext(context, context.OPTION_COMPARE());
                base.EnterOptionCompareStmt(context);
            }

            public override void EnterOptionExplicitStmt(VBAParser.OptionExplicitStmtContext context)
            {
                CheckContext(context, context.OPTION_EXPLICIT());
                base.EnterOptionExplicitStmt(context);
            }

            public override void EnterOptionPrivateModuleStmt(VBAParser.OptionPrivateModuleStmtContext context)
            {
                CheckContext(context, context.OPTION_PRIVATE_MODULE());
                base.EnterOptionPrivateModuleStmt(context);
            }

            public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
            {
                CheckContext(context, context.END_ENUM());
                base.EnterEnumerationStmt(context);
            }

            public override void EnterUdtDeclaration(VBAParser.UdtDeclarationContext context)
            {
                CheckContext(context, context.END_TYPE());
                base.EnterUdtDeclaration(context);
            }



            private void CheckContext(ParserRuleContext context, IParseTree subTreeToExamine)
            {
                if (subTreeToExamine?.GetText().Contains("_") ?? false)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
