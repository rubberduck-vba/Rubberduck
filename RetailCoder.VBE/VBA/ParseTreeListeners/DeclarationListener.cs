using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections;
using Rubberduck.Parsing;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class ParameterListener : VBABaseListener, IExtensionListener<VBAParser.AmbiguousIdentifierContext>
    {
        private readonly QualifiedModuleName _qualifiedName;
        private readonly IList<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _members =
            new List<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();

        private QualifiedMemberName _memberName;

        public ParameterListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> Members { get { return _members; } }

        public override void EnterArg(VBAParser.ArgContext context)
        {
            if (context.Parent.Parent.GetType() != typeof (VBAParser.EventStmtContext)
                && context.Parent.Parent.GetType() != typeof(VBAParser.DeclareStmtContext))
            {
                _members.Add(new QualifiedContext<VBAParser.AmbiguousIdentifierContext>(_memberName, context.ambiguousIdentifier()));
            }
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _memberName = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _memberName = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _memberName = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _memberName = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _memberName = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }
    }

    public class LocalDeclarationListener : VBABaseListener, IExtensionListener<VBAParser.AmbiguousIdentifierContext>
    {
        private readonly QualifiedModuleName _qualifiedName;
        private readonly IList<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> _members =
            new List<QualifiedContext<VBAParser.AmbiguousIdentifierContext>>();

        private QualifiedMemberName _memberName;

        public LocalDeclarationListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<VBAParser.AmbiguousIdentifierContext>> Members { get { return _members; } }

        public override void EnterArg(VBAParser.ArgContext context)
        {
            //note: args must be handled separately than variables; inspection fixes would destroy method signatures.
            //_members.Add(new QualifiedContext<VBAParser.AmbiguousIdentifierContext>(_memberName, context.ambiguousIdentifier()));
            return;
        }

        public override void EnterVariableSubStmt(VBAParser.VariableSubStmtContext context)
        {
            if (string.IsNullOrEmpty(_memberName.Name))
            {
                // ignore fields
                return;
            }

            _members.Add(new QualifiedContext<VBAParser.AmbiguousIdentifierContext>(_memberName, context.ambiguousIdentifier()));
        }

        public override void EnterConstSubStmt(VBAParser.ConstSubStmtContext context)
        {
            if (string.IsNullOrEmpty(_memberName.Name))
            {
                // ignore fields
                return;
            }

            _members.Add(new QualifiedContext<VBAParser.AmbiguousIdentifierContext>(_memberName, context.ambiguousIdentifier()));
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            _memberName = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _memberName = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _memberName = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _memberName = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _memberName = new QualifiedMemberName(_qualifiedName, context.ambiguousIdentifier().GetText());
        }
    }

    public class DeclarationListener : VBABaseListener, IExtensionListener<ParserRuleContext>
    {
        private readonly QualifiedModuleName _qualifiedName;

        private readonly IList<QualifiedContext<ParserRuleContext>> _members = 
            new List<QualifiedContext<ParserRuleContext>>();

        public DeclarationListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<ParserRuleContext>> Members { get { return _members; } }

        public override void EnterVariableStmt(VBAParser.VariableStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
            foreach (var child in context.variableListStmt().variableSubStmt())
            {
                _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, child));
            }
        }

        public override void EnterVisibility(VBAParser.VisibilityContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterEnumerationStmt(VBAParser.EnumerationStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterConstStmt(VBAParser.ConstStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
            foreach (var child in context.constSubStmt())
            {
                _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, child));
            }
        }

        public override void EnterTypeStmt(VBAParser.TypeStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterDeclareStmt(VBAParser.DeclareStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterEventStmt(VBAParser.EventStmtContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }

        public override void EnterArg(VBAParser.ArgContext context)
        {
            _members.Add(new QualifiedContext<ParserRuleContext>(_qualifiedName, context));
        }
    }

    public class DeclarationSectionListener : DeclarationListener
    {
        public DeclarationSectionListener(QualifiedModuleName qualifiedName)
            : base(qualifiedName)
        {
        }

        public override void EnterArg(VBAParser.ArgContext context)
        {
            return;
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            throw new WalkerCancelledException();
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            throw new WalkerCancelledException();
        }
    }
}
