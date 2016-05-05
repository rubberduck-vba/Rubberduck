using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerPrintCommand : CommandBase
    {
        public override bool CanExecute(object parameter)
        {
            var node = parameter as CodeExplorerComponentViewModel;
            if (node == null)
            {
                return false;
            }

            return node.Declaration.QualifiedName.QualifiedModuleName.Component.CodeModule.CountOfLines != 0;
        }

        public override void Execute(object parameter)
        {
            var node = (CodeExplorerComponentViewModel)parameter;
            var component = node.Declaration.QualifiedName.QualifiedModuleName.Component;

            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck",
                component.Name + ".txt");

            var text = component.CodeModule.Lines[1, component.CodeModule.CountOfLines];

            var printDoc = new PrintDocument { DocumentName = path };
            var pd = new PrintDialog
            {
                Document = printDoc,
                AllowSelection = true,
                AllowSomePages = true
            };

            if (pd.ShowDialog() == DialogResult.OK)
            {
                printDoc.PrintPage += (sender, printPageArgs) =>
                {
                    var font = new Font(new FontFamily("Consolas"), 10, FontStyle.Regular);
                    printPageArgs.Graphics.DrawString(text, font, Brushes.Black, 0, 0, new StringFormat());
                };
                printDoc.Print();
            }
        }
    }
}