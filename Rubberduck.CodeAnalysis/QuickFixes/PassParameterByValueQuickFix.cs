using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class PassParameterByValueQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public PassParameterByValueQuickFix(RubberduckParserState state)
            : base(typeof(ParameterCanBeByValInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            if (result.Target.ParentDeclaration.DeclarationType == DeclarationType.Event ||
                _state.DeclarationFinder.FindAllInterfaceMembers().Contains(result.Target.ParentDeclaration))
            {
                FixMethods(result.Target);
            }
            else
            {
                FixMethod((VBAParser.ArgContext)result.Target.Context, result.QualifiedSelection);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.PassParameterByValueQuickFix;

        private void FixMethods(Declaration target)
        {
            var declarationParameters =
                _state.AllUserDeclarations.Where(declaration => declaration.DeclarationType == DeclarationType.Parameter &&
                                                                Equals(declaration.ParentDeclaration, target.ParentDeclaration))
                    .OrderBy(o => o.Selection.StartLine)
                    .ThenBy(t => t.Selection.StartColumn)
                    .ToList();

            var parameterIndex = declarationParameters.IndexOf(target);
            if (parameterIndex == -1)
            {
                return; // should only happen if the parse results are stale; prevents a crash in that case
            }

            var members = target.ParentDeclaration.DeclarationType == DeclarationType.Event
                ? _state.AllUserDeclarations.FindHandlersForEvent(target.ParentDeclaration)
                    .Select(s => s.Item2)
                    .ToList()
                : _state.DeclarationFinder.FindInterfaceImplementationMembers(target.ParentDeclaration).Cast<Declaration>().ToList();

            foreach (var member in members)
            {
                var parameters =
                    _state.AllUserDeclarations.Where(declaration => declaration.DeclarationType == DeclarationType.Parameter &&
                                                                    Equals(declaration.ParentDeclaration, member))
                        .OrderBy(o => o.Selection.StartLine)
                        .ThenBy(t => t.Selection.StartColumn)
                        .ToList();

                FixMethod((VBAParser.ArgContext)parameters[parameterIndex].Context,
                    parameters[parameterIndex].QualifiedSelection);
            }

            FixMethod((VBAParser.ArgContext)declarationParameters[parameterIndex].Context,
                declarationParameters[parameterIndex].QualifiedSelection);
        }

        private void FixMethod(VBAParser.ArgContext context, QualifiedSelection qualifiedSelection)
        {
            var rewriter = _state.GetRewriter(qualifiedSelection.QualifiedName);
            if (context.BYREF() != null)
            {
                rewriter.Replace(context.BYREF(), Tokens.ByVal);
            }
            else
            {
                rewriter.InsertBefore(context.unrestrictedIdentifier().Start.TokenIndex, "ByVal ");
            }
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}