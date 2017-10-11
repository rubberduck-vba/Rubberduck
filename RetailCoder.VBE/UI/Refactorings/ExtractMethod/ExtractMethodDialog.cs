using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;

namespace Rubberduck.UI.Refactorings
{
    public partial class ExtractMethodDialog : Form, IRefactoringDialog<ExtractMethodViewModel>
    {
        public ExtractMethodViewModel ViewModel { get; }

        public ExtractMethodDialog()
        {
            InitializeComponent();
        }

        public ExtractMethodDialog(ExtractMethodViewModel vm) : this()
        {
            ViewModel = vm;
            ExtractMethodViewElement.DataContext = vm;
            vm.OnWindowClosed += ViewModel_OnWindowClosed;
        }

        void ViewModel_OnWindowClosed(object sender, DialogResult result)
        {
            DialogResult = result;
            Close();
        }
    }
}
