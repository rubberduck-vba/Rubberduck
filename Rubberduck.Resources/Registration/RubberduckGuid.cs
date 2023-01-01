namespace Rubberduck.Resources.Registration
{
    public static class RubberduckGuid
    {
        // Guid Suffix
        private const string GuidSuffix = "-43F0-3B33-B105-9B8188A6F040";

        // TypeLib Guid:
        private const string TypeLibGuidspace = "E07C84";
        public const string RubberduckTypeLibGuid = TypeLibGuidspace + "1C" + GuidSuffix;
        public const string RubberduckApiTypeLibGuid = TypeLibGuidspace + "1D" + GuidSuffix;
        
        // Addin Guids:
        private const string AddinGuidspace = "69E0F6";
        public const string ExtensionGuid = AddinGuidspace + "97" + GuidSuffix;
        public const string IDockableWindowHostGuid = AddinGuidspace + "98" + GuidSuffix;
        public const string DockableWindowHostGuid = AddinGuidspace + "99" + GuidSuffix;

        // Unit testing Guids:
        private const string UnitTestingGuidspace = "69E0F7";
        public const string AssertClassGuid = UnitTestingGuidspace + "DA" + GuidSuffix;
        public const string IAssertGuid = UnitTestingGuidspace + "DB" + GuidSuffix;
        public const string PermissiveAssertClassGuid = UnitTestingGuidspace + "DC" + GuidSuffix;
        public const string FakesProviderClassGuid = UnitTestingGuidspace + "DD" + GuidSuffix;
        public const string IFakesProviderGuid = UnitTestingGuidspace + "DE" + GuidSuffix;
        public const string IFakeGuid = UnitTestingGuidspace + "DF" + GuidSuffix;
        public const string IVerifyGuid = UnitTestingGuidspace + "E0" + GuidSuffix;
        public const string IStubGuid = UnitTestingGuidspace + "E1" + GuidSuffix;
        public const string IParamsGuid = UnitTestingGuidspace + "E2" + GuidSuffix;

        public const string ParamsClassGuid = UnitTestingGuidspace + "40" + GuidSuffix;
        public const string ParamsMsgBoxGuid = UnitTestingGuidspace + "41" + GuidSuffix;
        public const string ParamsInputBoxGuid = UnitTestingGuidspace + "42" + GuidSuffix;
        public const string ParamsEnvironGuid = UnitTestingGuidspace + "44" + GuidSuffix;
        public const string ParamsShellGuid = UnitTestingGuidspace + "47" + GuidSuffix;
        public const string ParamsSendKeysGuid = UnitTestingGuidspace + "48" + GuidSuffix;
        public const string ParamsKillGuid = UnitTestingGuidspace + "49" + GuidSuffix;
        public const string ParamsMkDirGuid = UnitTestingGuidspace + "4A" + GuidSuffix;
        public const string ParamsRmDirGuid = UnitTestingGuidspace + "4B" + GuidSuffix;
        public const string ParamsChDirGuid = UnitTestingGuidspace + "4C" + GuidSuffix;
        public const string ParamsChDriveGuid = UnitTestingGuidspace + "4D" + GuidSuffix;
        public const string ParamsCurDirGuid = UnitTestingGuidspace + "4E" + GuidSuffix;
        public const string ParamsRndGuid = UnitTestingGuidspace + "52" + GuidSuffix;
        public const string ParamsDeleteSettingGuid = UnitTestingGuidspace + "53" + GuidSuffix;
        public const string ParamsSaveSettingGuid = UnitTestingGuidspace + "54" + GuidSuffix;
        public const string ParamsGetSettingGuid = UnitTestingGuidspace + "55" + GuidSuffix;
        public const string ParamsRandomizeGuid = UnitTestingGuidspace + "56" + GuidSuffix;
        public const string ParamsGetAllSettingsGuid = UnitTestingGuidspace + "57" + GuidSuffix;
        public const string ParamsSetAttrGuid = UnitTestingGuidspace + "58" + GuidSuffix;
        public const string ParamsGetAttrGuid = UnitTestingGuidspace + "59" + GuidSuffix;
        public const string ParamsFileLenGuid = UnitTestingGuidspace + "5A" + GuidSuffix;
        public const string ParamsFileDateTimeGuid = UnitTestingGuidspace + "5B" + GuidSuffix;
        public const string ParamsFreeFileGuid = UnitTestingGuidspace + "5C" + GuidSuffix;
        public const string ParamsDirGuid = UnitTestingGuidspace + "5E" + GuidSuffix;
        public const string ParamsFileCopyGuid = UnitTestingGuidspace + "5F" + GuidSuffix;

        // Rubberduck API Guids:
        private const string ApiGuidspace = "69E0F7";
        public const string IDeclarationGuid = ApiGuidspace + "81" + GuidSuffix;
        public const string DeclarationClassGuid = ApiGuidspace + "82" + GuidSuffix;
        public const string IIdentifierReferenceGuid = ApiGuidspace + "83" + GuidSuffix;
        public const string IdentifierReferenceClassGuid = ApiGuidspace + "84" + GuidSuffix;
        public const string IParserGuid = ApiGuidspace + "85" + GuidSuffix;
        public const string ParserClassGuid = ApiGuidspace + "86" + GuidSuffix;
        public const string IParserEventsGuid = ApiGuidspace + "87" + GuidSuffix;
        public const string IDeclarationsGuid = ApiGuidspace + "88" + GuidSuffix;
        public const string DeclarationsClassGuid = ApiGuidspace + "89" + GuidSuffix;
        public const string IApiProviderGuid = ApiGuidspace + "8A" + GuidSuffix;
        public const string ApiProviderClassGuid = ApiGuidspace + "8B" + GuidSuffix;
        public const string IIdentifierReferencesGuid = ApiGuidspace + "8C" + GuidSuffix;
        public const string IdentifierReferencesClassGuid = ApiGuidspace + "8D" + GuidSuffix;

        // Enum Guids:
        private const string RecordGuidspace = "69E100";
        public const string DeclarationTypeGuid = RecordGuidspace + "23" + GuidSuffix;
        public const string AccessibilityGuid = RecordGuidspace + "24" + GuidSuffix;
        public const string ParserStateGuid = RecordGuidspace + "25" + GuidSuffix;

        // Debug Guids:
        private const string DebugGuidspace = "69E101";
        public const string DebugAddinObjectInterfaceGuid = DebugGuidspace + "23" + GuidSuffix;
        public const string DebugAddinObjectClassGuid = DebugGuidspace + "24" + GuidSuffix;
    }
}