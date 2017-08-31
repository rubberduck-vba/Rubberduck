// ReSharper disable InconsistentNaming

namespace Rubberduck
{
    public static class RubberduckGuid
    {
        // Addin Guids:
        public const string ExtensionGuid = "8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66";                 // shipped
        public const string DockableWindowHostGuid = "9CF1392A-2DC9-48A6-AC0B-E601A9802608";        // shipped

        // Rubberduck API Guids:
        public const string DeclarationClassGuid = "67940D0B-081A-45BE-B0B9-CAEAFE034BC0";          // shipped prior to 2.0.14
        public const string IdentifierReferenceClassGuid = "57F78E64-8ADF-4D81-A467-A0139B877D14";  // shipped prior to 2.0.14
        public const string ParserStateClassGuid = "28754D11-10CC-45FD-9F6A-525A65412B7A";          // shipped prior to 2.0.14
        public const string IParserStateEventsGuid = "3D8EAA28-8983-44D5-83AF-2EEC4C363079";        // shipped prior to 2.0.14

        // Unit testing Guids:
        private const string UnitTestingGuidspace = "-43F0-3B33-B105-9B8188A6F040";
        public const string AssertClassGuid = "69E194DA" + UnitTestingGuidspace;                    // shipped prior to 2.0.14
        public const string IAssertGuid = "69E194DB" + UnitTestingGuidspace;                        // added for 2.0.14
        public const string PermissiveAssertClassGuid = "40F71F29-D63F-4481-8A7D-E04A4B054501";     // shipped prior to 2.0.14
        public const string FakesProviderClassGuid = "69E194DD" + UnitTestingGuidspace;             // added for 2.0.14
        public const string IFakesProviderGuid = "69E194DE" + UnitTestingGuidspace;                 // added for 2.0.14
        public const string IFakeGuid = "69E194DF" + UnitTestingGuidspace;                          // added for 2.0.14
        public const string IVerifyGuid = "69E194E0" + UnitTestingGuidspace;                        // added for 2.0.14
        public const string IStubGuid = "69E194E1" + UnitTestingGuidspace;                          // added for 2.0.14

        // Ribbon Dispatcher Guids:
        private const string RibbonDispatcherGuidspace = "-43F1-3B33-B105-9B8188A6F040";
        public const string IClickedEvents              = "69E19400" + RibbonDispatcherGuidspace;
        public const string IClickedEventArgs           = "69E19401" + RibbonDispatcherGuidspace;
        public const string IToggledEvents              = "69E19402" + RibbonDispatcherGuidspace;
        public const string IToggledEventArgs           = "69E19403" + RibbonDispatcherGuidspace;
        public const string ISelectedEvents             = "69E19404" + RibbonDispatcherGuidspace;
        public const string ISelectedEventArgs          = "69E19405" + RibbonDispatcherGuidspace;

        public const string RdControlSize               = "69E19407" + RibbonDispatcherGuidspace;
        public const string Typelib                     = "69E19408" + RibbonDispatcherGuidspace;
        public const string IResourceManager            = "69E19409" + RibbonDispatcherGuidspace;
        public const string IAbstractDispatcher         = "69E1940A" + RibbonDispatcherGuidspace;
        public const string AbstractDispatcher          = "69E1940B" + RibbonDispatcherGuidspace;
        public const string IMain                       = "69E1940C" + RibbonDispatcherGuidspace;
        public const string Main                        = "69E1940D" + RibbonDispatcherGuidspace;
        public const string IRibbonFactory              = "69E1940E" + RibbonDispatcherGuidspace;
        public const string RibbonFactory               = "69E1940F" + RibbonDispatcherGuidspace;
        public const string IRibbonViewModel            = "69E19410" + RibbonDispatcherGuidspace;
        public const string RibbonViewModel             = "69E19411" + RibbonDispatcherGuidspace;
        public const string IRibbonCommon               = "69E19412" + RibbonDispatcherGuidspace;
        public const string RibbonCommon                = "69E19413" + RibbonDispatcherGuidspace;
        public const string IRibbonButton               = "69E19414" + RibbonDispatcherGuidspace;
        public const string RibbonButton                = "69E19415" + RibbonDispatcherGuidspace;
        public const string IRibbonCheckBox             = "69E19416" + RibbonDispatcherGuidspace;
        public const string RibbonCheckBox              = "69E19417" + RibbonDispatcherGuidspace;
        public const string IRibbonDropDown             = "69E19418" + RibbonDispatcherGuidspace;
        public const string RibbonDropDown              = "69E19419" + RibbonDispatcherGuidspace;
        public const string IRibbonGroup                = "69E1941A" + RibbonDispatcherGuidspace;
        public const string RibbonGroup                 = "69E1941B" + RibbonDispatcherGuidspace;
        public const string IRibbonTextLanguageControl  = "69E1941C" + RibbonDispatcherGuidspace;
        public const string RibbonTextLanguageControl   = "69E1941D" + RibbonDispatcherGuidspace;
        public const string IRibbonToggleButton         = "69E1941E" + RibbonDispatcherGuidspace;
        public const string RibbonToggleButton          = "69E1941F" + RibbonDispatcherGuidspace;
        public const string ISelectableItem             = "69E19420" + RibbonDispatcherGuidspace;
        public const string SelectableItem              = "69E19421" + RibbonDispatcherGuidspace;
    }
}