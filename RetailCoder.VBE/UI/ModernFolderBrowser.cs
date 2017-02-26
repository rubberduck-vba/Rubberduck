using System;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FileDialog = System.Windows.Forms.FileDialog;

namespace Rubberduck.UI
{
    public class ModernFolderBrowser : IFolderBrowser
    {
        // ReSharper disable InconsistentNaming
        private const BindingFlags DefaultBindingFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

        private static readonly AssemblyName FormsAssemblyName =
            Assembly.GetExecutingAssembly()
                .GetReferencedAssemblies()
                .FirstOrDefault(a => a.FullName.StartsWith("System.Windows.Forms"));

        private static readonly Assembly FormsAssembly = Assembly.Load(FormsAssemblyName);
        private static readonly Type IFileDialog = FormsAssembly.GetTypes().SingleOrDefault(t => t.Name.Equals("IFileDialog"));
        private static readonly MethodInfo OpenFileDialogCreateVistaDialog = typeof(System.Windows.Forms.OpenFileDialog).GetMethod("CreateVistaDialog", DefaultBindingFlags);
        private static readonly MethodInfo OpenFileDialogOnBeforeVistaDialog = typeof(System.Windows.Forms.OpenFileDialog).GetMethod("OnBeforeVistaDialog", DefaultBindingFlags);
        private static readonly MethodInfo FileDialogGetOptions = typeof(FileDialog).GetMethod("GetOptions", DefaultBindingFlags);
        private static readonly MethodInfo IFileDialogSetOptions = IFileDialog.GetMethod("SetOptions", DefaultBindingFlags);
        private static readonly MethodInfo IFileDialogShow = IFileDialog.GetMethod("Show", DefaultBindingFlags);
        private static readonly Type FOS = FormsAssembly.GetTypes().SingleOrDefault(t => t.Name.Equals("FOS"));
        private static readonly uint FOS_PICKFOLDERS = (uint)FOS.GetField("FOS_PICKFOLDERS").GetValue(null);
        private static readonly Type VistaDialogEvents = FormsAssembly.GetTypes().SingleOrDefault(t => t.Name.Equals("VistaDialogEvents"));
        private static readonly ConstructorInfo VistaDialogEventsCtor = VistaDialogEvents.GetConstructors().SingleOrDefault();
        private static readonly MethodInfo IFileDialogAdvise = IFileDialog.GetMethod("Advise", DefaultBindingFlags);
        private static readonly MethodInfo IFileDialogUnadvise = IFileDialog.GetMethod("Unadvise", DefaultBindingFlags);
        // ReSharper restore InconsistentNaming

        private readonly System.Windows.Forms.OpenFileDialog _dialog;
        private readonly object _newDialog;
        private readonly IEnvironmentProvider _environment;

        // ReSharper disable once UnusedParameter.Local - new folder button suppression isn't supported in this dialog.
        public ModernFolderBrowser(IEnvironmentProvider environment, string description, bool showNewFolderButton, string rootFolder)
        {
            _environment = environment;
            _root = rootFolder;
            _dialog = new System.Windows.Forms.OpenFileDialog
            {
                Title = description,
                InitialDirectory = _root,
                // ReSharper disable once LocalizableElement - This is an API keyword.
                Filter = "Folders|\n",
                AddExtension = false,
                CheckFileExists = false,
                DereferenceLinks = true,
                Multiselect = false
            };
            _newDialog = OpenFileDialogCreateVistaDialog.Invoke(_dialog, new object[] { });
            OpenFileDialogOnBeforeVistaDialog.Invoke(_dialog, new[] { _newDialog });
            var options = (uint)FileDialogGetOptions.Invoke(_dialog, new object[] { }) | FOS_PICKFOLDERS;
            IFileDialogSetOptions.Invoke(_newDialog, new object[] { options });
        }

        public ModernFolderBrowser(IEnvironmentProvider environment, string description, bool showNewFolderButton)
            : this(environment, description, showNewFolderButton, environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
        { }

        public ModernFolderBrowser(IEnvironmentProvider environment, string description) : this(environment, description, true) { }

        public string Description
        {
            get { return _dialog.Title; }
            set { _dialog.Title = value; }
        }

        //Does nothing - new folder button suppression isn't supported in this dialog.
        public bool ShowNewFolderButton
        {
            get { return true; }
            // ReSharper disable once ValueParameterNotUsed
            set { }
        }

        private string _root;
        public string RootFolder
        {
            get { return _root; }
            set
            {
                _root = value;
                _dialog.InitialDirectory = _root;
            }
        }

        public string SelectedPath
        {
            get { return _dialog.FileName; }
            set { _dialog.FileName = value; }
        }

        public DialogResult ShowDialog()
        {
            var sink = VistaDialogEventsCtor.Invoke(new object[] { _dialog });
            var cookie = 0u;
            var parameters = new[] { sink, cookie };
            IFileDialogAdvise.Invoke(_newDialog, parameters);
            //This is the cookie returned as a ref parameter in the call above.
            cookie = (uint)parameters[1];
            int returnValue;
            try
            {
                returnValue = (int)IFileDialogShow.Invoke(_newDialog, new object[] { IntPtr.Zero });
            }
            finally 
            {
                IFileDialogUnadvise.Invoke(_newDialog, new object[] { cookie });
                GC.KeepAlive(sink);                
            }
            return returnValue == 0 ? DialogResult.OK : DialogResult.Cancel;
        }

        public void Dispose()
        {
            _dialog.Dispose();
        }
    }
}
