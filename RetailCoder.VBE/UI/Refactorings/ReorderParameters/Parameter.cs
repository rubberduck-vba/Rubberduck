namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class Parameter
    {
        public string FullDeclaration { get; private set; }
        public int Index { get; private set; }
        public bool IsOptional { get; private set; }
        public bool IsParamArray { get; private set; }

        public Parameter(string fullDeclaration, int index)
        {
            FullDeclaration = fullDeclaration;
            Index = index;
            IsOptional = FullDeclaration.Contains("Optional");
            IsParamArray = FullDeclaration.Contains("ParamArray");
        }
    }
}
