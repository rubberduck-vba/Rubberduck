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

    public interface IExtractMethodRule
    {
        void setValidFlag(ref byte flags, IdentifierReference reference, Selection selection);
    }
    public class ExtractMethodRuleUsedBefore : IExtractMethodRule
    {
        public void setValidFlag(ref byte flags, IdentifierReference reference, Selection selection)
        {
            if (reference.Selection.StartLine < selection.StartLine)
                flags = (byte)(flags | 1);
        }
    }
    public class ExtractMethodRuleUsedAfter : IExtractMethodRule
    {
        public void setValidFlag(ref byte flags, IdentifierReference reference, Selection selection)
        {
            if (reference.Selection.StartLine > selection.EndLine)
                flags = (byte)(flags | 2);
        }
    }
    public class ExtractMethodRuleIsAssignedInSelection : IExtractMethodRule
    {
        public void setValidFlag(ref byte flags, IdentifierReference reference, Selection selection)
        {
            if (selection.StartLine <= reference.Selection.StartLine && reference.Selection.StartLine <= selection.EndLine)
            {
                if (reference.IsAssignment)
                    flags = (byte)(flags | 4);
            }
        }
    }
    public class ExtractMethodRuleInSelection : IExtractMethodRule
    {
        public void setValidFlag(ref byte flags, IdentifierReference reference, Selection selection)
        {
            if (selection.StartLine <= reference.Selection.StartLine && reference.Selection.StartLine <= selection.EndLine)
                flags = (byte)(flags | 8);
        }
    }
    public class ExtractMethodModel : IExtractMethodModel
    {
        private const string NEW_METHOD = "NewMethod";

        public ExtractMethodModel(IEnumerable<Declaration> declarations, QualifiedSelection selection, string selectedCode)
        {
            var items = declarations.ToList();
            _sourceMember = items.FindSelectedDeclaration(selection, DeclarationExtensions.ProcedureTypes, d => ((ParserRuleContext)d.Context.Parent).GetSelection());
            if (_sourceMember == null)
            {
                throw new InvalidOperationException("Invalid selection.");
            }

            _extractedMethod = new ExtractedMethod();

            _selection = selection;
            _selectedCode = selectedCode;

            var selectionStartLine = selection.Selection.StartLine;
            var selectionEndLine = selection.Selection.EndLine;

            var inScopeDeclarations = items.Where(item => item.ParentScope == _sourceMember.Scope).ToList();

            var rules = new List<IExtractMethodRule>(){
                new ExtractMethodRuleUsedBefore(),
                new ExtractMethodRuleUsedAfter(),
                new ExtractMethodRuleInSelection(),
                new ExtractMethodRuleIsAssignedInSelection()};
            
            _byref = new List<Declaration>();
            _byval = new List<Declaration>();
            _moveIn = new List<Declaration>();
            _declarationsToMove = new List<Declaration>();

            // https://github.com/rubberduck-vba/Rubberduck/wiki/Extract-Method-Refactoring-%3A-Workings---Determining-what-params-to-move
            foreach (var item in inScopeDeclarations)
            {
                var flags = new Byte();

                foreach (var oRef in item.References)
                {
                    foreach (var rule in rules)
                    {
                        rule.setValidFlag(ref flags, oRef, _selection.Selection);
                    }
                }

                if (flags < 4) { }
                else if (flags < 12)
                    _byref.Add(item);
                else if (flags == 12)
                    _declarationsToMove.Add(item);
                else if (flags > 12)
                    _byval.Add(item);

            }


            var methodParams = _byref.Select(dec => new ExtractedParameter(dec.AsTypeName, ExtractedParameter.PassedBy.ByRef,dec.IdentifierName))
                                .Union(_byval.Select(dec => new ExtractedParameter(dec.AsTypeName, ExtractedParameter.PassedBy.ByVal ,dec.IdentifierName)));

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
        }

        private readonly List<Declaration> _byref;
        private readonly List<Declaration> _byval;
        private readonly List<Declaration> _moveIn;

        private readonly Declaration _sourceMember;
        public Declaration SourceMember { get { return _sourceMember; } }

        private readonly QualifiedSelection _selection;
        public QualifiedSelection Selection { get { return _selection; } }

        private readonly string _selectedCode;
        public string SelectedCode { get { return _selectedCode; } }

        private readonly List<Declaration> _locals;
        public IEnumerable<Declaration> Locals { get { return _locals; } }

        private readonly IEnumerable<ExtractedParameter> _input;
        public IEnumerable<ExtractedParameter> Inputs { get { return _input; } }

        private readonly IEnumerable<ExtractedParameter> _output;
        public IEnumerable<ExtractedParameter> Outputs { get { return _output; } }

        private readonly List<Declaration> _declarationsToMove;
        public IEnumerable<Declaration> DeclarationsToMove { get { return _declarationsToMove; } }

        private readonly IExtractedMethod _extractedMethod;
        public IExtractedMethod Method { get { return _extractedMethod; } }

    }
}