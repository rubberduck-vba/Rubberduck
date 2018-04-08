using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.ParseTreeListeners
{
    public abstract class AttributeAnnotationListener : VBAParserBaseListener, IInspectionListener
    {
        protected AttributeAnnotationListener(RubberduckParserState state)
        {
            State = state;
        }

        public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;
        public QualifiedModuleName CurrentModuleName { get; set; }
        public void ClearContexts() => _contexts.Clear();

        protected RubberduckParserState State { get; }
        protected Lazy<Declaration> Module { get; private set; }
        protected Lazy<IDictionary<string, Declaration>> Members { get; private set; }

        protected Declaration FirstMember
        {
            get
            {
                return CurrentScopeDeclaration = State.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .Where(declaration => declaration.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName))
                    .OrderBy(declaration => declaration.Selection)
                    .FirstOrDefault();
            }
        }

        protected bool HasMembers { get; private set; }
        protected Declaration CurrentScopeDeclaration { get; set; }

        private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();

        protected void AddContext(QualifiedContext<ParserRuleContext> context)
        {
            _contexts.Add(context);
        }

        private void SetCurrentScope(string memberName = null)
        {
            HasMembers = !string.IsNullOrEmpty(memberName);
            CurrentScopeDeclaration = HasMembers ? Members.Value[memberName] : Module.Value;
        }

        public override void EnterModule(VBAParser.ModuleContext context)
        {
            Module = new Lazy<Declaration>(() => State.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .SingleOrDefault(m => m.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName)));

            Members = new Lazy<IDictionary<string, Declaration>>(() => State.DeclarationFinder
                .Members(CurrentModuleName)
                .Where(m => !m.DeclarationType.HasFlag(DeclarationType.Module))
                .GroupBy(m => m.IdentifierName)
                .ToDictionary(m => m.Key, m => m.FirstOrDefault()));

            SetCurrentScope();
        }

        public override void ExitModule(VBAParser.ModuleContext context)
        {
            CurrentScopeDeclaration = null;
        }

        public override void ExitModuleDeclarations(VBAParser.ModuleDeclarationsContext context)
        {
            var firstMember = Members.Value.Values.OrderBy(d => d.Selection).FirstOrDefault();
            if (firstMember != null)
            {
                CurrentScopeDeclaration = firstMember;
            }
            else
            {
                CurrentScopeDeclaration = State.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                    .SingleOrDefault(d => d.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName));
            }
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            SetCurrentScope(Identifier.GetName(context.subroutineName()));
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            SetCurrentScope(Identifier.GetName(context.functionName()));
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            SetCurrentScope(Identifier.GetName(context.functionName()));
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            SetCurrentScope(Identifier.GetName(context.subroutineName()));
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            SetCurrentScope(Identifier.GetName(context.subroutineName()));
        }
    }
}