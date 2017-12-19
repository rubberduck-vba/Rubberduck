﻿namespace Rubberduck.UI.SourceControl
{
    /// <summary>
    /// Interaction logic for BranchesView.xaml
    /// </summary>
    public partial class BranchesView : IControlView
    {
        public BranchesView()
        {
            InitializeComponent();
        }

        public BranchesView(IControlViewModel vm) : this()
        {
            DataContext = vm;
        }

        public IControlViewModel ViewModel => (IControlViewModel)DataContext;
    }
}
