namespace Rubberduck.Refactorings.ImplementInterface
{
    public class Parameter
    {
        public string Accessibility { get; set; }
        public string Name { get; set; }
        public string AsTypeName { get; set; }

        public override string ToString()
        {
            return Accessibility + " " + Name + " As " + AsTypeName;
        }
    }
}
