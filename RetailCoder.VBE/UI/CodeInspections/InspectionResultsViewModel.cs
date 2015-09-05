using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Rubberduck.Inspections;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.CodeInspections
{
    public class InspectionResultsViewModel : ViewModelBase
    {
        private readonly IInspector _inspector;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public InspectionResultsViewModel(IInspector inspector, ICodePaneWrapperFactory wrapperFactory)
        {
            _inspector = inspector;
            _wrapperFactory = wrapperFactory;
        }

        private readonly ICommand _runAllInspections;
    }
}
