using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Microsoft.Office.Interop.Outlook;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    /// <summary>
    /// A listener that gets all module members that are visible outside the module.
    /// </summary>
    public class ExposedMemberListener : VBListenerBase, IExtensionListener<ParserRuleContext>
    {
        private readonly QualifiedModuleName _qualifiedName;

        private readonly List<QualifiedContext<ParserRuleContext>> _members = 
            new List<QualifiedContext<ParserRuleContext>>();

        private static readonly string[] PublicTokens = 
            new[] {Tokens.Public, Tokens.Global, Tokens.Friend};

        public ExposedMemberListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<ParserRuleContext>> Members
        {
            get { return _members; }
        }

        private void AddIfExposed(VBParser.VisibilityContext context)
        {
            if (context == null)
            {
                return;
            }

            var visibility = context.GetText();
            if (PublicTokens.Contains(visibility))
            {
                _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context.Parent as ParserRuleContext));
            }
        }

        public override void EnterVariableStmt(VBParser.VariableStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }

        public override void EnterEnumerationStmt(VBParser.EnumerationStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }

        public override void EnterConstStmt(VBParser.ConstStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }

        public override void EnterTypeStmt(VBParser.TypeStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }

        public override void EnterDeclareStmt(VBParser.DeclareStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }

        public override void EnterEventStmt(VBParser.EventStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }

        public override void EnterFunctionStmt(VBParser.FunctionStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }

        public override void EnterPropertyGetStmt(VBParser.PropertyGetStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }

        public override void EnterPropertyLetStmt(VBParser.PropertyLetStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }

        public override void ExitPropertySetStmt(VBParser.PropertySetStmtContext context)
        {
            AddIfExposed(context.Visibility());
        }
    }
}
