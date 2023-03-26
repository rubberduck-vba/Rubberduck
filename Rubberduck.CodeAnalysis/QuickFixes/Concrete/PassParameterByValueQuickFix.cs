using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Modifies a parameter to be passed by value.
    /// </summary>
    /// <inspections>
    /// <inspection name="ParameterCanBeByValInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething(ByRef value As Long)
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal value As Long)
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class PassParameterByValueQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public PassParameterByValueQuickFix(IDeclarationFinderProvider declarationFinderProvider)
            : base(typeof(ParameterCanBeByValInspection), typeof(MisleadingByRefParameterInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (result.Target.ParentDeclaration.DeclarationType == DeclarationType.Event ||
                _declarationFinderProvider.DeclarationFinder.FindAllInterfaceMembers().Contains(result.Target.ParentDeclaration))
            {
                FixMethods(result.Target, rewriteSession);
            }
            else
            {
                FixMethod((VBAParser.ArgContext)result.Target.Context, result.QualifiedSelection, rewriteSession);
            }
        }

        private void FixMethods(Declaration target, IRewriteSession rewriteSession)
        {
            var declarationParameters =
                _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                    .Where(declaration => Equals(declaration.ParentDeclaration, target.ParentDeclaration))
                    .OrderBy(o => o.Selection.StartLine)
                    .ThenBy(t => t.Selection.StartColumn)
                    .ToList();

            var parameterIndex = declarationParameters.IndexOf(target);
            if (parameterIndex == -1)
            {
                return; // should only happen if the parse results are stale; prevents a crash in that case
            }

            var members = target.ParentDeclaration.DeclarationType == DeclarationType.Event
                ? _declarationFinderProvider.DeclarationFinder
                    .FindEventHandlers(target.ParentDeclaration)
                    .ToList()
                : _declarationFinderProvider.DeclarationFinder
                    .FindInterfaceImplementationMembers(target.ParentDeclaration)
                    .ToList();

            foreach (var member in members)
            {
                var parameters =
                    _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
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

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.PassParameterByValueQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}