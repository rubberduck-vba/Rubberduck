using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    /// <summary>
    /// Describes usages of a declared identifier.
    /// </summary>
    [ComVisible(false)]
    public enum ExtractedDeclarationUsage
    {
        /// <summary>
        /// A variable that isn't used in selection, 
        /// will not be extracted.
        /// </summary>
        NotUsed,

        /// <summary>
        /// A variable that is only used in selection, 
        /// will be moved to the extracted method.
        /// </summary>
        UsedOnlyInSelection,
        
        /// <summary>
        /// A variable that is used before selection,
        /// will be extracted as a parameter.
        /// </summary>
        UsedBeforeSelection,
        
        /// <summary>
        /// A variable that is used after selection,
        /// will be extracted as a <c>ByRef</c> parameter 
        /// or become the extracted method's return value.
        /// </summary>
        UsedAfterSelection
    }

    [ComVisible(true)]
    public class ExtractedParameter
    {
        public enum PassedBy
        {
            ByRef,
            ByVal
        }

        public ExtractedParameter(string name, string typeName, PassedBy passed)
        {
            Name = name;
            TypeName = typeName;
            Passed = passed;
        }

        public string Name { get; set; }
        public string TypeName { get; set; }
        public PassedBy Passed { get; set; }

        public override string ToString()
        {
            return Passed.ToString() + ' ' + Name + ' ' + Tokens.As + ' ' + TypeName;
        }
    }

    [ComVisible(false)]
    public class ExtractMethodPresenter
    {
        private readonly IExtractMethodDialog _view;

        private readonly IParseTree _parentMethodTree;
        private IDictionary<VisualBasic6Parser.AmbiguousIdentifierContext, ExtractedDeclarationUsage> _parentMethodDeclarations;

        private readonly IEnumerable<ExtractedParameter> _input;
        private readonly IEnumerable<ExtractedParameter> _output;
        private readonly IEnumerable<VisualBasic6Parser.AmbiguousIdentifierContext> _locals; 

        private readonly string _selectedCode;
        private readonly VBE _vbe;
        private readonly Selection _selection;

        public ExtractMethodPresenter(VBE vbe, IExtractMethodDialog dialog, IParseTree parentMethod, Selection selection)
        {
            _vbe = vbe;
            _selection = selection;

            _view = dialog;
            _parentMethodTree = parentMethod;
            _selectedCode = vbe.ActiveCodePane.CodeModule.get_Lines(selection.StartLine, selection.LineCount);

            _parentMethodDeclarations = ExtractMethodRefactoring.GetParentMethodDeclarations(parentMethod, selection);
            
            var input = _parentMethodDeclarations.Where(kvp => kvp.Value == ExtractedDeclarationUsage.UsedBeforeSelection).ToList();
            var output = _parentMethodDeclarations.Where(kvp => kvp.Value == ExtractedDeclarationUsage.UsedAfterSelection).ToList();
            
            _locals = _parentMethodDeclarations.Where(kvp => kvp.Value == ExtractedDeclarationUsage.UsedOnlyInSelection).Select(kvp => kvp.Key);
            _input = ExtractParameters(input);
            _output = ExtractParameters(output);
        }

        private IEnumerable<ExtractedParameter> ExtractParameters(IList<KeyValuePair<VisualBasic6Parser.AmbiguousIdentifierContext, ExtractedDeclarationUsage>> declarations)
        {
            var consts = declarations
                .Where(kvp => kvp.Key.Parent is VisualBasic6Parser.ConstSubStmtContext)
                .Select(kvp => kvp.Key.Parent)
                .Cast<VisualBasic6Parser.ConstSubStmtContext>()
                .Select(constant => new ExtractedParameter(
                    constant.ambiguousIdentifier().GetText(),
                    constant.asTypeClause() == null
                        ? Tokens.Variant
                        : constant.asTypeClause().type().GetText(),
                    ExtractedParameter.PassedBy.ByVal));

            var variables = declarations
                .Where(kvp => kvp.Key.Parent is VisualBasic6Parser.VariableSubStmtContext)
                .Select(kvp => new ExtractedParameter(
                    kvp.Key.GetText(),
                    ((VisualBasic6Parser.VariableSubStmtContext) kvp.Key.Parent).asTypeClause() == null
                        ? Tokens.Variant
                        : ((VisualBasic6Parser.VariableSubStmtContext) kvp.Key.Parent).asTypeClause().type().GetText(),
                    ExtractedParameter.PassedBy.ByVal));

            var arguments = declarations
                .Where(kvp => kvp.Key.Parent is VisualBasic6Parser.ArgContext)
                .Select(kvp => new ExtractedParameter(
                    kvp.Key.GetText(),
                    ((VisualBasic6Parser.ArgContext)kvp.Key.Parent).asTypeClause() == null
                        ? Tokens.Variant
                        : ((VisualBasic6Parser.ArgContext)kvp.Key.Parent).asTypeClause().type().GetText(),
                    ExtractedParameter.PassedBy.ByVal));

            return consts.Union(variables.Union(arguments));
        }

        public void Show()
        {
            _view.MethodName = "Method1";
            _view.Inputs = _input.ToList();
            _view.Outputs = _output.Select(output => new ExtractedParameter(output.Name, output.TypeName, ExtractedParameter.PassedBy.ByRef)).ToList();
            _view.Locals = _locals.Select(variable => new ExtractedParameter(variable.GetText(), string.Empty, ExtractedParameter.PassedBy.ByVal)).ToList();

            var returnValues = new[] { new ExtractedParameter("(none)", string.Empty, ExtractedParameter.PassedBy.ByVal) }
                .Union(_view.Outputs)
                .Union(_view.Inputs)
                .Union(_view.Locals)
                .ToList();

            _view.ReturnValues = returnValues;
            if (_output.Count() == 1)
            {
                _view.ReturnValue = _output.Single();
            }

            _view.RefreshPreview += _view_RefreshPreview;
            _view.OnRefreshPreview();

            var result = _view.ShowDialog();
            if (result != DialogResult.OK)
            {
                return;
            }

            _vbe.ActiveCodePane.CodeModule.DeleteLines(_selection.StartLine, _selection.LineCount - 1);
            _vbe.ActiveCodePane.CodeModule.ReplaceLine(_selection.StartLine, GetMethodCall());

            _vbe.ActiveCodePane.CodeModule.AddFromString(GetExtractedMethod());
        }

        private void _view_RefreshPreview(object sender, EventArgs e)
        {
            var hasReturnValue = _view.ReturnValue != null && _view.ReturnValue.Name != "(none)";
            _view.CanSetReturnValue = 
                hasReturnValue && !IsValueType(_view.ReturnValue.TypeName);

            Preview();
        }

        private void Preview()
        {
            _view.Preview = GetExtractedMethod();
        }

        private string GetMethodCall()
        {
            string result;
            var returnValueName = _view.ReturnValue.Name;
            var argsList = string.Join(", ", _view.Parameters.Select(p => p.Name));
            if (returnValueName != "(none)")
            {
                var setter = _view.SetReturnValue ? Tokens.Set + ' ' : string.Empty;
                result = setter + returnValueName + " = " + _view.MethodName + '(' + argsList + ')';
            }
            else
            {
                result = _view.MethodName + ' ' + argsList;
            }
            
            return "    " + result; // todo: smarter indentation
        }

        private static readonly IEnumerable<string> _valueTypes = new[]
        {
            Tokens.Boolean,
            Tokens.Byte,
            Tokens.Currency,
            Tokens.Date,
            Tokens.Decimal,
            Tokens.Double,
            Tokens.Integer,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.Single,
            Tokens.String
        };

        public static bool IsValueType(string typeName)
        {
            return _valueTypes.Contains(typeName);
        }

        [ComVisible(false)]
        private string GetExtractedMethod()
        {
            const string newLine = "\r\n";

            var access = _view.Accessibility.ToString();
            var keyword = Tokens.Sub;
            var returnType = string.Empty;

            var isFunction = _view.ReturnValue != null && _view.ReturnValue.Name != "(none)";
            if (isFunction)
            {
                keyword = Tokens.Function;
                returnType = Tokens.As + ' ' + _view.ReturnValue.TypeName;
            }

            var parameters = "(" + string.Join(", ", _view.Parameters) + ")";

            var result = access + ' ' + keyword + ' ' + _view.MethodName + parameters + ' ' + returnType + newLine;

            result += newLine + _selectedCode + newLine;

            if (isFunction)
            {
                // return value by assigning the method itself:
                var setter = _view.SetReturnValue ? Tokens.Set + ' ' : string.Empty;
                result += "    " + setter + _view.MethodName + " = " + _view.ReturnValue.Name + newLine;
            }

            result += Tokens.End + ' ' + keyword + newLine;

            return newLine + result + newLine;
        }
    }
}
