using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ResultVariableDeclaration : Declaration
    {
        public ResultVariableDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration function,
            string asTypeName,
            string typeHint,
            bool isArray)
            : base(
                  qualifiedName,
                  function,
                  function,
                  asTypeName,
                  typeHint,
                  isSelfAssigned:false,
                  isWithEvents:false,
                  accessibility:Accessibility.Implicit,
                  declarationType:DeclarationType.ResultVariable,
                  context:null,
                  attributesPassContext:null,
                  selection:Selection.Home,
                  isArray,
                  asTypeContext:null,
                  isUserDefined:false)
        {
            Function = function;
        }

        public Declaration Function { get; }
    }
}
