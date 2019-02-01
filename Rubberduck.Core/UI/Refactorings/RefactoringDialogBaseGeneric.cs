using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public readonly struct DialogData
    {
        public string Caption { get; }
        public int MinimumHeight { get; }
        public int MinimumWidth { get; }

        public DialogData(string caption, int minimumHeight, int minimumWidth)
        {
            Caption = caption;
            MinimumHeight = minimumHeight;
            MinimumWidth = minimumWidth;
        }

        public static DialogData Create(string caption, int minimumHeight, int minimumWidth)
        {
            return new DialogData(caption, minimumHeight, minimumWidth);
        }
    }

    public class RefactoringDialogBase<TModel, TView, TViewModel> : RefactoringDialogBase, IRefactoringDialog<TModel, TView, TViewModel>
        where TModel : class
        where TView : class, IRefactoringView<TModel>
        where TViewModel : class, IRefactoringViewModel<TModel>
    {
        public RefactoringDialogBase(DialogData dialogData, TModel model, TView view, TViewModel viewModel) 
        {
            Model = model;
            ViewModel = viewModel;

            View = view;
            View.DataContext = ViewModel;
            ViewModel.OnWindowClosed += ViewModel_OnWindowClosed;

            MinHeight = dialogData.MinimumHeight;
            MinWidth = dialogData.MinimumWidth;

            // ReSharper disable once RedundantBaseQualifier
            // We don't want virtual calls here so we need to explicitly call base.
            base.Text = dialogData.Caption;

            // Note that user control must be set after dialog data has been consumed to ensure
            // correct sizing of the dialog
            System.Diagnostics.Debug.Assert(View is System.Windows.Controls.UserControl);
            UserControl = View as System.Windows.Controls.UserControl;
        }

        public TModel Model { get; }
        public TView View { get; }
        public TViewModel ViewModel { get; }
        
        public new RefactoringDialogResult DialogResult { get; protected set; }
        public new virtual RefactoringDialogResult ShowDialog()
        {
            // The return of ShowDialog is meaningless; we use the DialogResult which the commands set.
            var result = base.ShowDialog();
            return DialogResult;
        }

        protected virtual void ViewModel_OnWindowClosed(object sender, RefactoringDialogResult result)
        {
            DialogResult = result;
            Close();
        }

        public object DataContext => UserControl.DataContext;
    }
}
