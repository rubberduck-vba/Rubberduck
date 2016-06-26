using Rubberduck.RegexAssistant;

namespace Rubberduck.UI.RegexAssistant
{
    class RegexAssistantViewModel : ViewModelBase
    {
        public RegexAssistantViewModel()
        {
            RecalculateDescription();
        }

        public bool GlobalFlag {
            get
            {
                return _globalFlag;
            }
            set
            {
                _globalFlag = value;
                RecalculateDescription();
            }
        }
        public bool IgnoreCaseFlag
        {
            get
            {
                return _ignoreCaseFlag;
            }
            set
            {
                _ignoreCaseFlag = value;
                RecalculateDescription();
            }
        }
        public string Pattern
        {
            get
            {
                return _pattern;
            }
            set
            {
                _pattern = value;
                RecalculateDescription();
            }
        }

        private string _description;
        private bool _globalFlag;
        private bool _ignoreCaseFlag;
        private string _pattern;

        private void RecalculateDescription()
        {
            if (_pattern.Equals(string.Empty))
            {
                _description = "No Pattern given";
                return;
            }
            _description = new Pattern(_pattern, _ignoreCaseFlag, _ignoreCaseFlag).Description;
        }

        public string DescriptionResults
        {
            get
            {
                return _description;
            }
        }
    }
}
