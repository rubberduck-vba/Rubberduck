namespace Rubberduck.AutoComplete
{
    public class AutoCompleteClosingParenthese : AutoCompleteBase
    {
        public override string InputToken => "(";
        public override string OutputToken => ")";
    }
}
