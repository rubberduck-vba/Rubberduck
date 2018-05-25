namespace Rubberduck.AutoComplete
{
    public class AutoCompleteClosingBracket : AutoCompleteBase
    {
        public override string InputToken => "[";
        public override string OutputToken => "]";
    }
}
