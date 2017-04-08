using System;
using System.ComponentModel;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using System.Xml;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;

namespace Rubberduck.UI.Controls
{
    //see http://stackoverflow.com/a/20823917/4088852
    public class BindableTextEditor : TextEditor, INotifyPropertyChanged
    {
        public BindableTextEditor()
        {
            WordWrap = false;

            var highlighter = LoadHighlighter("Rubberduck.UI.Controls.vba.xshd");
            SyntaxHighlighting = highlighter;

            //Style hyperlinks so they look like comments. Note - this needs to move if used for user code.
            TextArea.TextView.LinkTextUnderline = false;
            TextArea.TextView.LinkTextForegroundBrush = new SolidColorBrush(Colors.Green);
            Options.RequireControlModifierForHyperlinkClick = false;
            //This needs some work if hyperlinks need to open in an external browser.
            Options.EnableHyperlinks = false;
            Options.EnableEmailHyperlinks = true;
        }

        public new string Text
        {
            get { return base.Text; }
            set { base.Text = value; }
        }

        public static readonly DependencyProperty TextProperty =
            DependencyProperty.Register("Text", typeof(string), typeof(BindableTextEditor), new PropertyMetadata((obj, args) =>
            {
                var target = (BindableTextEditor)obj;
                target.Text = (string)args.NewValue;
            }));

        protected override void OnTextChanged(EventArgs e)
        {
            RaisePropertyChanged("Text");
            base.OnTextChanged(e);
        }

        public void RaisePropertyChanged(string property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private static IHighlightingDefinition LoadHighlighter(string resource)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resource))
            {
                if (stream == null)
                {
                    return null;
                }
                using (var reader = new XmlTextReader(stream))
                {
                    return HighlightingLoader.Load(reader, HighlightingManager.Instance);
                }
            }
        }
    }
}
