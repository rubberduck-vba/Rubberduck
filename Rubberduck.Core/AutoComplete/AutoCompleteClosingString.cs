namespace Rubberduck.AutoComplete
{
    public class AutoCompleteClosingString : AutoCompleteBase
    {
        public override string InputToken => "\"";
        public override string OutputToken => "\"";
    }
}
