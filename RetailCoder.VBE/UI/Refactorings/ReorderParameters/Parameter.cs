namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    /// <summary>
    /// Contains data about a method parameter for the Reorder Parameters refactoring.
    /// </summary>
    public class Parameter
    {
        public string FullDeclaration { get; private set; }
        public int Index { get; private set; }
        public bool IsOptional { get; private set; }
        public bool IsParamArray { get; private set; }

        /// <summary>
        /// Creates a Parameter.
        /// </summary>
        /// <param name="fullDeclaration">The full declaration of the parameter, such as "ByVal param As Integer".</param>
        /// <param name="index">The index of the parameter in the list of method parameters.</param>
        public Parameter(string fullDeclaration, int index)
        {
            FullDeclaration = fullDeclaration;
            Index = index;
            IsOptional = FullDeclaration.Contains("Optional");
            IsParamArray = FullDeclaration.Contains("ParamArray");
        }
    }
}
