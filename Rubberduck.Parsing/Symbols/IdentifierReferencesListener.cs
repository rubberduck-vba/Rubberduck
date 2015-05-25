using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// a work-in-progress replacement for <see cref="IdentifierReferenceListener"/>.
    /// </summary>
    public class IdentifierReferencesListener : VBABaseListener
    {
        private readonly Declarations _declarations;
        private readonly QualifiedModuleName _qualifiedName;

        public IdentifierReferencesListener(QualifiedModuleName qualifiedName, Declarations declarations)
        {
            _qualifiedName = qualifiedName;
            _declarations = declarations;

            _fields = new HashSet<Declaration>(FindScopedDeclarations(_qualifiedName.ToString()));
        }

        private HashSet<Declaration> _locals;
        private HashSet<Declaration> _fields;

        private static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        private IEnumerable<Declaration> FindScopedDeclarations(string scope)
        {
            return _declarations.Items.Where(item =>
                item.QualifiedName.QualifiedModuleName.Project == _qualifiedName.Project
                && item.ComponentName == _qualifiedName.ComponentName
                && item.ParentScope == scope
                && !ProcedureTypes.Contains(item.DeclarationType));
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            var name = context.ambiguousIdentifier().GetText();
            _locals = new HashSet<Declaration>(FindScopedDeclarations(_qualifiedName + "." + name));
        }

        public override void ExitSubStmt(VBAParser.SubStmtContext context)
        {
            _locals = null;
        }

        public override void EnterFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            var name = context.ambiguousIdentifier().GetText();
            _locals = new HashSet<Declaration>(FindScopedDeclarations(_qualifiedName + "." + name));
        }

        public override void ExitFunctionStmt(VBAParser.FunctionStmtContext context)
        {
            _locals = null;
        }

        public override void EnterPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            var name = context.ambiguousIdentifier().GetText();
            _locals = new HashSet<Declaration>(FindScopedDeclarations(_qualifiedName + "." + name));
        }

        public override void ExitPropertyGetStmt(VBAParser.PropertyGetStmtContext context)
        {
            _locals = null;
        }

        public override void EnterPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            var name = context.ambiguousIdentifier().GetText();
            _locals = new HashSet<Declaration>(FindScopedDeclarations(_qualifiedName + "." + name));
        }

        public override void ExitPropertyLetStmt(VBAParser.PropertyLetStmtContext context)
        {
            _locals = null;
        }

        public override void EnterPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            var name = context.ambiguousIdentifier().GetText();
            _locals = new HashSet<Declaration>(FindScopedDeclarations(_qualifiedName + "." + name));
        }

        public override void ExitPropertySetStmt(VBAParser.PropertySetStmtContext context)
        {
            _locals = null;
        }


    }
}