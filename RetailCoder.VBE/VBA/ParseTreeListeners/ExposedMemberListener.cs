using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections;
using Rubberduck.Parsing;

namespace Rubberduck.VBA.ParseTreeListeners
{
    /// <summary>
    /// A listener that gets all module members that are visible outside the module.
    /// </summary>
    public class ExposedMemberListener : VBABaseListener, IExtensionListener<ParserRuleContext>
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

        private void AddIfExposed(VBAParser.VisibilityContext context)
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

        public override void EnterVariableStmt(VBAParser.VariableStmtContext context)
        {
            AddIfExposed(context.visibility());
        }

        public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            AddIfExposed(context.visibility());
        }

        public override void EnterConstStmt(VBAParser.ConstStmtContext context)
        {
            AddIfExposed(context.visibility());
        }

        public override void EnterTypeStmt(VBAParser.TypeStmtContext context)
        {
            AddIfExposed(context.visibility());
        }

        public override void EnterDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            AddIfExposed(context.visibility());
        }

        public override void EnterEventStmt(VBAParser.EventStmtContext context)
        {
            AddIfExposed(context.visibility());
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            AddIfExposed(context.visibility());
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            AddIfExposed(context.visibility());
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            AddIfExposed(context.visibility());
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            AddIfExposed(context.visibility());
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            AddIfExposed(context.visibility());
        }
    }
}
