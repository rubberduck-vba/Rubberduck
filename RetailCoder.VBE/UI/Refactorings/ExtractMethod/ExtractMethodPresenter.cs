using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Antlr4.Runtime.Tree;
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

        public ExtractedParameter(string name, string typeName, PassedBy by)
        {
            Name = name;
            TypeName = typeName;
            By = by;
        }

        public string Name { get; set; }
        public string TypeName { get; set; }
        public PassedBy By { get; set; }
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

        private readonly Selection _selection;

        public ExtractMethodPresenter(IExtractMethodDialog dialog, IParseTree parentMethod, Selection selection)
        {
            _view = dialog;
            _parentMethodTree = parentMethod;
            _selection = selection;

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
                        : ((VisualBasic6Parser.VariableSubStmtContext) kvp.Key.Parent).asTypeClause().GetText(),
                    ExtractedParameter.PassedBy.ByVal));

            return consts.Union(variables);
        }

        public void Show()
        {
            var result = _view.ShowDialog();
            if (result != DialogResult.OK)
            {
                return;
            }

            // todo: proceed with method extraction refactoring.
        }

        public string MethodName { get; private set; }
        
        public VBAccessibility MethodAccessibility { get; private set; }
        public IEnumerable<VBAccessibility> AvailableAccessibilities
        {
            get
            {
                return new[]
                {
                    VBAccessibility.Private,
                    VBAccessibility.Public,
                    VBAccessibility.Friend
                };
            } 
        }

        public IdentifierNode MethodReturnValue { get; private set; }

        public string NewMethodPreview { get; private set; }
    }
}
