namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class Parameter
    {
        public string Name { get; private set; }
        public string Variable { get; private set; }
        public int Index { get; private set; }
        public bool IsOptional { get; private set; }

        public Parameter(string name, string variable, int index)
        {
            Name = name;
            Variable = variable;
            Index = index;
            IsOptional = Variable.Contains("Optional");
        }
    }
}
