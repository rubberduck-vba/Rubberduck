using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Listeners
{
    public class VariableReferencesListener : VBABaseListener, IExtensionListener<VBAParser.AmbiguousIdentifierContext>
    {
        private readonly QualifiedModuleName _qualifiedName;
        private readonly IList<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _members = 
            new List<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();

        private QualifiedMemberName _currentMember;

        public VariableReferencesListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
            _currentMember = new QualifiedMemberName(_qualifiedName, string.Empty);
        }

        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> Members { get { return _members; } }

        public override void EnterAmbiguousIdentifier(VBAParser.AmbiguousIdentifierContext context)
        {
            var ignored = new[]
            {
                typeof (VBAParser.AsTypeClauseContext),
                typeof (VBAParser.VariableSubStmtContext),
                typeof (VBAParser.ConstSubStmtContext),
                typeof (VBAParser.SubStmtContext), 
                typeof (VBAParser.FunctionStmtContext),
                typeof (VBAParser.PropertyGetStmtContext),
                typeof (VBAParser.PropertyLetStmtContext),
                typeof (VBAParser.PropertySetStmtContext)
            };

            if (ignored.All(type => type != context.Parent.GetType()))
            {
                _members.Add(new QualifiedContext<VBAParser.AmbiguousIdentifierContext>(_currentMember, context));
            }
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _currentMember = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }
    }
}