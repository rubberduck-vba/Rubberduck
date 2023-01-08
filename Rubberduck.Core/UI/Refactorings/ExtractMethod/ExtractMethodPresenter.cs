using NLog;
using Rubberduck.Interaction;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.Refactorings.ExtractMethod
{
    internal class ExtractMethodPresenter : RefactoringPresenterBase<ExtractMethodModel>, IExtractMethodPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RefactoringsUI.ExtractMethod_Caption, 400, 800); //TODO - get appropriate size

        private readonly IMessageBox _messageBox;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ExtractMethodPresenter(ExtractMethodModel model,
            IRefactoringDialogFactory dialogFactory, IMessageBox messageBox) : 
            base(DialogData, model, dialogFactory)
        {
            _messageBox = messageBox;
        }

        override public ExtractMethodModel Show()
        {
            return base.Show();
        }
    }
}
