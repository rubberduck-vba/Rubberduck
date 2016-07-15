using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ParameterCanBeByValInspectionResult(IInspection inspection, RubberduckParserState state, Declaration target, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, qualifiedName.QualifiedModuleName, context, target)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new PassParameterByValueQuickFix(state, Target, Context, QualifiedSelection),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, inspection.AnnotationName)
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterCanBeByValInspectionResultFormat, Target.IdentifierName); }
        }
    }

    public class PassParameterByValueQuickFix : CodeInspectionQuickFix
    {
        private readonly RubberduckParserState _state;
        private readonly Declaration _target;

        public PassParameterByValueQuickFix(RubberduckParserState state, Declaration target, ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.PassParameterByValueQuickFix)
        {
            _state = state;
            _target = target;
        }

        public override void Fix()
        {
            if (_target.ParentDeclaration.DeclarationType == DeclarationType.Event ||
                _state.AllUserDeclarations.FindInterfaceMembers().Contains(_target.ParentDeclaration))
            {
                FixMethods();
            }
            else
            {
                FixMethod((VBAParser.ArgContext)Context, Selection);
            }
        }

        private void FixMethods()
        {
            var declarationParameters =
                    _state.AllUserDeclarations.Where(declaration => declaration.DeclarationType == DeclarationType.Parameter &&
                                                      declaration.ParentDeclaration == _target.ParentDeclaration)
                                .OrderBy(o => o.Selection.StartLine)
                                .ThenBy(t => t.Selection.StartColumn)
                                .ToList();

            var parameterIndex = declarationParameters.IndexOf(_target);
            if (parameterIndex == -1)
            {
                return; // should only happen if the parse results are stale; prevents a crash in that case
            }

            var members = _target.ParentDeclaration.DeclarationType == DeclarationType.Event
                ? _state.AllUserDeclarations.FindHandlersForEvent(_target.ParentDeclaration)
                    .Select(s => s.Item2)
                    .ToList()
                : _state.AllUserDeclarations.FindInterfaceImplementationMembers(_target.ParentDeclaration).ToList();

            foreach (var member in members)
            {
                var parameters =
                    _state.AllUserDeclarations.Where(declaration => declaration.DeclarationType == DeclarationType.Parameter &&
                                                      declaration.ParentDeclaration == member)
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
            var selectionLength = context.BYREF() == null ? 0 : 6;

            var module = qualifiedSelection.QualifiedName.Component.CodeModule;
            var lines = module.Lines[context.Start.Line, 1];

            var result = lines.Remove(context.Start.Column, selectionLength).Insert(context.Start.Column, Tokens.ByVal + ' ');
            module.ReplaceLine(context.Start.Line, result);
        }
    }
}
