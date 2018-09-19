namespace Rubberduck.RegexAssistant
{
    public interface IAtom : IDescribable
    {
        Quantifier Quantifier { get; }
        string Specifier { get; }
    }
}
