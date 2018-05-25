namespace Rubberduck.AutoComplete
{
    public class AutoCompleteClosingBrace : AutoCompleteBase
    {
        public override string InputToken => "{";
        public override string OutputToken => "}";
    }
}
