namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class Parameter
    {
        public string IdentifierName { get; private set; }
        public string FullDeclaration { get; private set; }
        public int Index { get; private set; }
        public bool IsOptional { get; private set; }

        public Parameter(string identifierName, string fullDeclaration, int index)
        {
            IdentifierName = identifierName;
            FullDeclaration = fullDeclaration;
            Index = index;
            IsOptional = FullDeclaration.Contains("Optional");
        }
    }
}
