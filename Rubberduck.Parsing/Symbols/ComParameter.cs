namespace Rubberduck.Parsing.Symbols
{
    public class ComParameter
    {
        public bool IsArray { get; set; }
        public bool IsByRef { get; set;}
        public string Name { get; set;}

        public ComParameter(string name, bool byRef)
        {
            Name = name;
            IsByRef = byRef;
        }
    }
}
