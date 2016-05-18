using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodProc : IExtractMethodProc
    {
        public string createProc(IExtractMethodModel model)
        {
            var newLine = Environment.NewLine;

            var method = model.Method;

            var access = method.Accessibility.ToString();
            var keyword = Tokens.Sub;
            var asTypeClause = string.Empty;

            /* implement simple sub(byref) style : doesn't require a gui.
             * can go back to implement more complicated version once testing passes
            if (isFunction)
            {
                keyword = Tokens.Function;
                asTypeClause = Tokens.As + ' ' + model.Method.ReturnValue.TypeName;
            }
            */

            var extractedParams = method.Parameters.Select(p => ExtractedParameter.PassedBy.ByRef + " " + p.Name + " " + Tokens.As + " " + p.TypeName);
            var parameters = "(" + string.Join(", ", extractedParams) + ")";

            var result = access + ' ' + keyword + ' ' + method.MethodName + parameters + ' ' + asTypeClause + newLine;
            
            
            var localConsts = model.Locals.Where(e => e.DeclarationType == DeclarationType.Constant)
                .Cast<ValuedDeclaration>()
                .Select(e => "    " + Tokens.Const + ' ' + e.IdentifierName + ' ' + Tokens.As + ' ' + e.AsTypeName + " = " + e.Value);

            var localVariables = model.Locals.Where(e => e.DeclarationType == DeclarationType.Variable)
                .Where(e => model.Method.Parameters.All(param => param.Name != e.IdentifierName))
                .Select(e => e.Context)
                .Cast<VBAParser.VariableSubStmtContext>()
                .Select(e => "    " + Tokens.Dim + ' ' + e.identifier().GetText() +
                    (e.LPAREN() == null
                        ? string.Empty
                        : e.LPAREN().GetText() + (e.subscripts() == null ? string.Empty : e.subscripts().GetText()) + e.RPAREN().GetText()) + ' ' +
                        (e.asTypeClause() == null ? string.Empty : e.asTypeClause().GetText()));
            var locals = string.Join(newLine, localConsts.Union(localVariables)
                            .Where(local => !model.SelectedCode.Contains(local)).ToArray()) + newLine;

            result += locals + model.SelectedCode + newLine;

            /* implement simple sub(byref) style : doesn't require a gui.
            if (isFunction)
            {
                // return value by assigning the method itself:
                var setter = model.Method.SetReturnValue ? Tokens.Set + ' ' : string.Empty;
                result += "    " + setter + model.Method.MethodName + " = " + model.Method.ReturnValue.Name + newLine;
            }
             */

            result += Tokens.End + ' ' + keyword + newLine;

            return newLine + result + newLine;
        }
    }
}
