using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.ExtractMethod
{

    public class ExtractMethodModel : IExtractMethodModel
    {
        private const string NEW_METHOD = "NewMethod";

        public ExtractMethodModel(List<IExtractMethodRule> emRules, IExtractedMethod extractedMethod)
        {
            _rules = emRules;
            _extractedMethod = extractedMethod;
        }


        public void extract(IEnumerable<Declaration> declarations, QualifiedSelection selection, string selectedCode)
        {
            var items = declarations.ToList();
            var sourceMember = items.FindSelectedDeclaration(selection, DeclarationExtensions.ProcedureTypes, d => ((ParserRuleContext)d.Context.Parent).GetSelection());
            if (sourceMember == null)
            {
                throw new InvalidOperationException("Invalid selection.");
            }

            var inScopeDeclarations = items.Where(item => item.ParentScope == sourceMember.Scope).ToList();

            _byref = new List<Declaration>();
            _byval = new List<Declaration>();
            _declarationsToMove = new List<Declaration>();

            _extractedMethod = new ExtractedMethod();
            

            var selectionToRemove = new List<Selection>();
            var selectionStartLine = selection.Selection.StartLine;
            var selectionEndLine = selection.Selection.EndLine;

            var methodInsertLine = sourceMember.Context.Stop.Line + 1;
            _positionForNewMethod = new Selection(methodInsertLine, 1, methodInsertLine, 1);

            // https://github.com/rubberduck-vba/Rubberduck/wiki/Extract-Method-Refactoring-%3A-Workings---Determining-what-params-to-move
            foreach (var item in inScopeDeclarations)
            {
                var flags = new Byte();

                foreach (var oRef in item.References)
                {
                    foreach (var rule in _rules)
                    {
                        rule.setValidFlag(ref flags, oRef, selection.Selection);
                    }
                }

                //TODO: extract this to seperate class.
                if (flags < 4) { /*ignore the variable*/ }
                else if (flags < 12)
                    _byref.Add(item);
                else if (flags == 12)
                    _declarationsToMove.Add(item);
                else if (flags > 12)
                    _byval.Add(item);

            }

            selectionToRemove.Add(selection.Selection);
            _declarationsToMove.ForEach(d => selectionToRemove.Add(d.Selection));

            var methodCallPositionStartLine = selectionStartLine - selectionToRemove.Count(s => s.StartLine < selectionStartLine);
            _positionForMethodCall = new Selection(methodCallPositionStartLine, 1, methodCallPositionStartLine, 1);

            var methodParams = _byref.Select(dec => new ExtractedParameter(dec.AsTypeName, ExtractedParameter.PassedBy.ByRef, dec.IdentifierName))
                                .Union(_byval.Select(dec => new ExtractedParameter(dec.AsTypeName, ExtractedParameter.PassedBy.ByVal, dec.IdentifierName)));

            // iterate until we have a non-clashing method name.
            var newMethodName = NEW_METHOD;

            var newMethodInc = 0;
            while (declarations.FirstOrDefault(d =>
                DeclarationExtensions.ProcedureTypes.Contains(d.DeclarationType)
                && d.IdentifierName.Equals(newMethodName)) != null)
            {
                newMethodInc++;
                newMethodName = NEW_METHOD + newMethodInc;
            }

            _extractedMethod.MethodName = newMethodName;
            _extractedMethod.ReturnValue = null;
            _extractedMethod.Accessibility = Accessibility.Private;
            _extractedMethod.SetReturnValue = false;
            _extractedMethod.Parameters = methodParams.ToList();

            _selection = selection;
            _selectedCode = selectedCode;
            _selectionToRemove = selectionToRemove.ToList();

        }

        private List<Declaration> _byref;
        private List<Declaration> _byval;
        private List<Declaration> _moveIn;

        private Declaration _sourceMember;
        public Declaration SourceMember { get { return _sourceMember; } }

        private QualifiedSelection _selection;
        public QualifiedSelection Selection { get { return _selection; } }

        private string _selectedCode;
        public string SelectedCode { get { return _selectedCode; } }

        private List<Declaration> _locals;
        public IEnumerable<Declaration> Locals { get { return _locals; } }

        private IEnumerable<ExtractedParameter> _input;
        public IEnumerable<ExtractedParameter> Inputs { get { return _input; } }

        private IEnumerable<ExtractedParameter> _output;
        public IEnumerable<ExtractedParameter> Outputs { get { return _output; } }

        private List<Declaration> _declarationsToMove;
        public IEnumerable<Declaration> DeclarationsToMove { get { return _declarationsToMove; } }

        private IExtractedMethod _extractedMethod;

        private IEnumerable<IExtractMethodRule> _rules;

        public IExtractedMethod Method { get { return _extractedMethod; } }


        private Selection _positionForMethodCall;
        public Selection PositionForMethodCall { get { return _positionForMethodCall; } }

        public string NewMethodCall { get { return _extractedMethod.NewMethodCall(); } }

        private Selection _positionForNewMethod;
        public Selection PositionForNewMethod { get { return _positionForNewMethod; } } 
        IEnumerable<Selection> _selectionToRemove;
        private List<IExtractMethodRule> emRules;

        public IEnumerable<Selection> SelectionToRemove { get { return _selectionToRemove; } }


    }
}