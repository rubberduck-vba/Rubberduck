using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonDispatcher {
    [ComVisible(true)][CLSCompliant(true)]
    public class CustomAppRibbonViewModel : AbstractRibbon, IRibbonExtensibility {

        public CustomAppRibbonViewModel() : base() {;}

        public string GetCustomUI(string RibbonID) => GetResourceText("RubberDuck.RibbonSupport.CustomAppRibbon.xml");

        [SuppressMessage("Microsoft.Design", "CA1061:DoNotHideBaseClassMethods", 
            Justification = "This is simply how the Ribbon design works to share the Callback methods between multiple Office products & documents.")]
        [CLSCompliant(false)]
        public override void OnRibbonLoad(IRibbonUI ribbonUI) {
            base.OnRibbonLoad(ribbonUI);

            _standardButtonsViewModel = new StandardButtonsViewModel(RibbonFactory, (s,e) => _customButtonsViewModel.SetVisible(e.IsPressed));
            _customButtonsViewModel   = new CustomButtonsViewModel(RibbonFactory);
        }

        StandardButtonsViewModel     _standardButtonsViewModel;
        CustomButtonsViewModel       _customButtonsViewModel;

        #region Helpers

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for(int i = 0; i < resourceNames.Length; ++i) {
                if(string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using(StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if(resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
        #endregion
    }
}
