using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.QuickFixes
{
    public class WriteOnlyPropertyQuickFix : QuickFixBase 
    {
        private readonly Declaration _target;

        public WriteOnlyPropertyQuickFix(ParserRuleContext context, Declaration target)
            : base(context, target.QualifiedSelection, InspectionsUI.WriteOnlyPropertyQuickFix)
        {
            _target = target;
        }

        public override void Fix()
        {
            var parameters = ((IDeclarationWithParameter) _target).Parameters.Cast<ParameterDeclaration>().ToList();

            var signatureParams = parameters.Except(new[] {parameters.Last()}).Select(GetParamText);
            var propertyGet = "Public Property Get " + _target.IdentifierName + "(" + string.Join(", ", signatureParams) +
                              ") As " + parameters.Last().AsTypeName + Environment.NewLine + "End Property";

            var module = _target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            module.InsertLines(_target.Selection.StartLine, propertyGet);
        }

        private string GetParamText(ParameterDeclaration param)
        {
            return (((VBAParser.ArgContext)param.Context).BYVAL() == null ? "ByRef " : "ByVal ") + param.IdentifierName + " As " + param.AsTypeName;
        }
    }
}
