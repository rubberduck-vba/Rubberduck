using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodModel
    {
        public ExtractMethodModel(IActiveCodePaneEditor editor, IEnumerable<Declaration> declarations, QualifiedSelection selection)
        {
            var items = declarations.ToList();

            _sourceMember = items.FindSelectedDeclaration(selection, DeclarationExtensions.ProcedureTypes, d => ((ParserRuleContext)d.Context.Parent).GetSelection());
            if (_sourceMember == null)
            {
                throw new InvalidOperationException("Invalid selection.");
            }

            _extractedMethod = new ExtractedMethod();

            _selection = selection;
            _selectedCode = editor.GetLines(selection.Selection);

            var inScopeDeclarations = items.Where(item => item.ParentScope == _sourceMember.Scope).ToList();

            var inSelection = inScopeDeclarations.SelectMany(item => item.References)
                .Where(item => selection.Selection.Contains(item.Selection))
                .ToList();

            var usedInSelection = new HashSet<Declaration>(inScopeDeclarations.Where(item =>
                selection.Selection.Contains(item.Selection) &&
                item.References.Any(reference => inSelection.Contains(reference))));

            var usedBeforeSelection = new HashSet<Declaration>(inScopeDeclarations.Where(item =>
                item.Selection.StartLine < selection.Selection.StartLine ||
                item.References.Any(reference => reference.Selection.StartLine < selection.Selection.StartLine)));

            var usedAfterSelection = new HashSet<Declaration>(inScopeDeclarations.Where(item =>
                item.Selection.StartLine > selection.Selection.StartLine ||
                item.References.Any(reference => reference.Selection.StartLine > selection.Selection.EndLine)));

            // identifiers used inside selection and before selection (or if it's a parameter) are candidates for parameters:
            var input = inScopeDeclarations.Where(item =>
                usedInSelection.Contains(item) && (usedBeforeSelection.Contains(item) || item.DeclarationType == DeclarationType.Parameter)).ToList();

            // identifiers used inside selection and after selection are candidates for return values:
            var output = inScopeDeclarations.Where(item =>
                usedInSelection.Contains(item) && usedAfterSelection.Contains(item))
                .ToList();

            // identifiers used only inside and/or after selection are candidates for locals:
            _locals = inScopeDeclarations.Where(item => item.DeclarationType != DeclarationType.Parameter && (
                item.References.All(reference => inSelection.Contains(reference))
                || (usedAfterSelection.Contains(item) && (!usedBeforeSelection.Contains(item)))))
                .ToList();

            // locals that are only used in selection are candidates for being moved into the new method:
            _declarationsToMove = _locals.Where(item => !usedAfterSelection.Contains(item)).ToList();

            _output = output.Select(declaration =>
                new ExtractedParameter(declaration.AsTypeName, ExtractedParameter.PassedBy.ByRef, declaration.IdentifierName));

            _input = input.Where(declaration => !output.Contains(declaration))
                .Select(declaration =>
                    new ExtractedParameter(declaration.AsTypeName, ExtractedParameter.PassedBy.ByVal, declaration.IdentifierName));
        }

        private readonly Declaration _sourceMember;
        public Declaration SourceMember { get { return _sourceMember; } }

        private readonly QualifiedSelection _selection;
        public QualifiedSelection Selection { get { return _selection; } }

        private readonly string _selectedCode;
        public string SelectedCode { get { return _selectedCode; } }

        private readonly List<Declaration> _locals;
        public IEnumerable<Declaration> Locals { get {return _locals;} }

        private readonly IEnumerable<ExtractedParameter> _input;
        public IEnumerable<ExtractedParameter> Inputs { get { return _input; } }

        private readonly IEnumerable<ExtractedParameter> _output;
        public IEnumerable<ExtractedParameter> Outputs { get {return _output; } }

        private readonly List<Declaration> _declarationsToMove;
        public IEnumerable<Declaration> DeclarationsToMove { get { return _declarationsToMove; } }

        private readonly ExtractedMethod _extractedMethod;
        public ExtractedMethod Method { get { return _extractedMethod; } }

        public class ExtractedMethod
        {
            public string MethodName { get; set; }
            public Accessibility Accessibility { get; set; }
            public bool SetReturnValue { get; set; }
            public ExtractedParameter ReturnValue { get; set; }
            public IEnumerable<ExtractedParameter> Parameters { get; set; }
        }
    }
}