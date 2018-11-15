using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class PassParameterByValueQuickFix : QuickFixBase
    {
        //TODO: Change this to IDeclarationFinderProvider once the FIXME below is handled.
        private readonly RubberduckParserState _state;

        public PassParameterByValueQuickFix(RubberduckParserState state)
            : base(typeof(ParameterCanBeByValInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (result.Target.ParentDeclaration.DeclarationType == DeclarationType.Event ||
                _state.DeclarationFinder.FindAllInterfaceMembers().Contains(result.Target.ParentDeclaration))
            {
                FixMethods(result.Target, rewriteSession);
            }
            else
            {
                FixMethod((VBAParser.ArgContext)result.Target.Context, result.QualifiedSelection, rewriteSession);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.PassParameterByValueQuickFix;

        private void FixMethods(Declaration target, IRewriteSession rewriteSession)
        {
            var declarationParameters =
                _state.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                    .Where(declaration => Equals(declaration.ParentDeclaration, target.ParentDeclaration))
                    .OrderBy(o => o.Selection.StartLine)
                    .ThenBy(t => t.Selection.StartColumn)
                    .ToList();

            var parameterIndex = declarationParameters.IndexOf(target);
            if (parameterIndex == -1)
            {
                return; // should only happen if the parse results are stale; prevents a crash in that case
            }

            //FIXME: Make this use the DeclarationFinder.
            var members = target.ParentDeclaration.DeclarationType == DeclarationType.Event
                ? _state.AllUserDeclarations.FindHandlersForEvent(target.ParentDeclaration)
                    .Select(s => s.Item2)
                    .ToList()
                : _state.DeclarationFinder.FindInterfaceImplementationMembers(target.ParentDeclaration).Cast<Declaration>().ToList();

            foreach (var member in members)
            {
                var parameters =
                    _state.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                        .Where(declaration => Equals(declaration.ParentDeclaration, member))
                        .OrderBy(o => o.Selection.StartLine)
                        .ThenBy(t => t.Selection.StartColumn)
                        .ToList();

                FixMethod((VBAParser.ArgContext)parameters[parameterIndex].Context,
                    parameters[parameterIndex].QualifiedSelection, rewriteSession);
            }

            FixMethod((VBAParser.ArgContext)declarationParameters[parameterIndex].Context,
                declarationParameters[parameterIndex].QualifiedSelection, rewriteSession);
        }

        private void FixMethod(VBAParser.ArgContext context, QualifiedSelection qualifiedSelection, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(qualifiedSelection.QualifiedName);
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