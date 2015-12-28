namespace Rubberduck.Refactorings.ImplementInterface
{
    public class Parameter
    {
        public string ParamAccessibility { get; set; }
        public string ParamName { get; set; }
        public string ParamType { get; set; }

        public override string ToString()
        {
            return ParamAccessibility + " " + ParamName + " As " + ParamType;
        }
    }
}