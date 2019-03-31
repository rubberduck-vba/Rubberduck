using NUnit.Framework;
using Moq;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class IndenterSettingsTests
    {
        // Defaults
        private const int DefaultAlignDimColumn = 15;
        private const EndOfLineCommentStyle DefaultEndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
        private const EmptyLineHandling DefaultEmptyLineHandlingMethod = EmptyLineHandling.Ignore;
        private const int DefaultEndOfLineCommentColumnSpaceAlignment = 50;
        private const int DefaultIndentSpaces = 4;
        private const int DefaultLinesBetweenProcedures = 1;

        // Nondefaults
        private const int NondefaultAlignDimColumn = 16;
        private const EndOfLineCommentStyle NondefaultEndOfLineCommentStyle = EndOfLineCommentStyle.Absolute;
        private const EmptyLineHandling NondefaultEmptyLineHandlingMethod = EmptyLineHandling.Remove;
        private const int NondefaultEndOfLineCommentColumnSpaceAlignment = 60;
        private const int NondefaultIndentSpaces = 2;
        private const int NondefaultLinesBetweenProcedures = 2;

        public static Rubberduck.SmartIndenter.IndenterSettings GetMockIndenterSettings(bool nondefault = false)
        {
            var output = new Mock<Rubberduck.SmartIndenter.IndenterSettings>();

            output.SetupProperty(s => s.IndentEntireProcedureBody);
            output.SetupProperty(s => s.IndentFirstCommentBlock);
            output.SetupProperty(s => s.IndentFirstDeclarationBlock);
            output.SetupProperty(s => s.IgnoreEmptyLinesInFirstBlocks);
            output.SetupProperty(s => s.AlignCommentsWithCode);
            output.SetupProperty(s => s.AlignContinuations);
            output.SetupProperty(s => s.IgnoreOperatorsInContinuations);
            output.SetupProperty(s => s.IndentCase);
            output.SetupProperty(s => s.IndentEnumTypeAsProcedure);
            output.SetupProperty(s => s.ForceDebugPrintInColumn1);
            output.SetupProperty(s => s.ForceDebugAssertInColumn1);
            output.SetupProperty(s => s.ForceStopInColumn1);
            output.SetupProperty(s => s.ForceCompilerDirectivesInColumn1);
            output.SetupProperty(s => s.IndentCompilerDirectives);
            output.SetupProperty(s => s.AlignDims);
            output.SetupProperty(s => s.AlignDimColumn);
            output.SetupProperty(s => s.EndOfLineCommentStyle);
            output.SetupProperty(s => s.EndOfLineCommentColumnSpaceAlignment);
            output.SetupProperty(s => s.EmptyLineHandlingMethod);
            output.SetupProperty(s => s.IndentSpaces);
            output.SetupProperty(s => s.VerticallySpaceProcedures);
            output.SetupProperty(s => s.LinesBetweenProcedures);

            output.Object.IndentEntireProcedureBody = !nondefault;
            output.Object.IndentFirstCommentBlock = !nondefault;
            output.Object.IndentFirstDeclarationBlock = !nondefault;
            output.Object.IgnoreEmptyLinesInFirstBlocks = nondefault;
            output.Object.AlignCommentsWithCode = !nondefault;
            output.Object.AlignContinuations = !nondefault;
            output.Object.IgnoreOperatorsInContinuations = !nondefault;
            output.Object.IndentCase = nondefault;
            output.Object.IndentEnumTypeAsProcedure = nondefault;
            output.Object.ForceDebugPrintInColumn1 = nondefault;
            output.Object.ForceDebugAssertInColumn1 = nondefault;
            output.Object.ForceStopInColumn1 = nondefault;
            output.Object.ForceCompilerDirectivesInColumn1 = nondefault;
            output.Object.IndentCompilerDirectives = !nondefault;
            output.Object.AlignDims = nondefault;
            output.Object.AlignDimColumn = nondefault ? NondefaultAlignDimColumn : DefaultAlignDimColumn;
            output.Object.EndOfLineCommentStyle = nondefault ? NondefaultEndOfLineCommentStyle : DefaultEndOfLineCommentStyle;
            output.Object.EndOfLineCommentColumnSpaceAlignment = nondefault ? NondefaultEndOfLineCommentColumnSpaceAlignment : DefaultEndOfLineCommentColumnSpaceAlignment;
            output.Object.EmptyLineHandlingMethod = nondefault ? NondefaultEmptyLineHandlingMethod : DefaultEmptyLineHandlingMethod;
            output.Object.IndentSpaces = nondefault ? NondefaultIndentSpaces : DefaultIndentSpaces;
            output.Object.VerticallySpaceProcedures = !nondefault;
            output.Object.LinesBetweenProcedures = nondefault ? NondefaultLinesBetweenProcedures : DefaultLinesBetweenProcedures;

            return output.Object;
        }

        private Configuration GetDefaultConfig()
        {
            var userSettings = new UserSettings(null, null, null, null, null, null, GetMockIndenterSettings(), null);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var userSettings = new UserSettings(null, null, null, null, null, null, GetMockIndenterSettings(true), null);
            return new Configuration(userSettings);
        }

        [Category("Settings")]
        [Test]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new IndenterSettingsViewModel(customConfig, null);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(config.UserSettings.IndenterSettings.AlignCommentsWithCode, viewModel.AlignCommentsWithCode);
                Assert.AreEqual(config.UserSettings.IndenterSettings.AlignContinuations, viewModel.AlignContinuations);
                Assert.AreEqual(config.UserSettings.IndenterSettings.AlignDimColumn, viewModel.AlignDimColumn);
                Assert.AreEqual(config.UserSettings.IndenterSettings.AlignDims, viewModel.AlignDims);
                Assert.AreEqual(config.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, viewModel.EndOfLineCommentColumnSpaceAlignment);
                Assert.AreEqual(config.UserSettings.IndenterSettings.EndOfLineCommentStyle, viewModel.EndOfLineCommentStyle);
                Assert.AreEqual(config.UserSettings.IndenterSettings.EmptyLineHandlingMethod, viewModel.EmptyLineHandlingMethod);
                Assert.AreEqual(config.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, viewModel.ForceCompilerDirectivesInColumn1);
                Assert.AreEqual(config.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, viewModel.ForceDebugStatementsInColumn1);
                Assert.AreEqual(config.UserSettings.IndenterSettings.ForceDebugPrintInColumn1, viewModel.ForceDebugPrintInColumn1);
                Assert.AreEqual(config.UserSettings.IndenterSettings.ForceDebugAssertInColumn1, viewModel.ForceDebugAssertInColumn1);
                Assert.AreEqual(config.UserSettings.IndenterSettings.ForceStopInColumn1, viewModel.ForceStopInColumn1);
                Assert.AreEqual(config.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, viewModel.IgnoreOperatorsInContinuations);
                Assert.AreEqual(config.UserSettings.IndenterSettings.IndentCase, viewModel.IndentCase);
                Assert.AreEqual(config.UserSettings.IndenterSettings.IndentEnumTypeAsProcedure, viewModel.IndentEnumTypeAsProcedure);
                Assert.AreEqual(config.UserSettings.IndenterSettings.IndentCompilerDirectives, viewModel.IndentCompilerDirectives);
                Assert.AreEqual(config.UserSettings.IndenterSettings.IndentEntireProcedureBody, viewModel.IndentEntireProcedureBody);
                Assert.AreEqual(config.UserSettings.IndenterSettings.IndentFirstCommentBlock, viewModel.IndentFirstCommentBlock);
                Assert.AreEqual(config.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, viewModel.IndentFirstDeclarationBlock);
                Assert.AreEqual(config.UserSettings.IndenterSettings.IgnoreEmptyLinesInFirstBlocks, viewModel.IgnoreEmptyLinesInFirstBlocks);
                Assert.AreEqual(config.UserSettings.IndenterSettings.IndentSpaces, viewModel.IndentSpaces);
                Assert.AreEqual(config.UserSettings.IndenterSettings.VerticallySpaceProcedures, viewModel.VerticallySpaceProcedures);
                Assert.AreEqual(config.UserSettings.IndenterSettings.LinesBetweenProcedures, viewModel.LinesBetweenProcedures);
            });
        }

        [Category("Settings")]
        [Test]
        public void SetDefaultsWorks()
        {
            var viewModel = new IndenterSettingsViewModel(GetNondefaultConfig(), null);

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignCommentsWithCode, viewModel.AlignCommentsWithCode);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignContinuations, viewModel.AlignContinuations);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDimColumn, viewModel.AlignDimColumn);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDims, viewModel.AlignDims);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, viewModel.EndOfLineCommentColumnSpaceAlignment);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle, viewModel.EndOfLineCommentStyle);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EmptyLineHandlingMethod, viewModel.EmptyLineHandlingMethod);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, viewModel.ForceCompilerDirectivesInColumn1);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceDebugPrintInColumn1, viewModel.ForceDebugPrintInColumn1);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceDebugAssertInColumn1, viewModel.ForceDebugAssertInColumn1);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceStopInColumn1, viewModel.ForceStopInColumn1);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, viewModel.ForceDebugStatementsInColumn1);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, viewModel.IgnoreOperatorsInContinuations);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCase, viewModel.IndentCase);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEnumTypeAsProcedure, viewModel.IndentEnumTypeAsProcedure);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCompilerDirectives, viewModel.IndentCompilerDirectives);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody, viewModel.IndentEntireProcedureBody);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock, viewModel.IndentFirstCommentBlock);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, viewModel.IndentFirstDeclarationBlock);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IgnoreEmptyLinesInFirstBlocks, viewModel.IgnoreEmptyLinesInFirstBlocks);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentSpaces, viewModel.IndentSpaces);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.VerticallySpaceProcedures, viewModel.VerticallySpaceProcedures);
                Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.LinesBetweenProcedures, viewModel.LinesBetweenProcedures);
            });
        }

        [Category("Settings")]
        [Test]
        public void AlignCommentsWithCodeIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignCommentsWithCode, viewModel.AlignCommentsWithCode);
        }

        [Category("Settings")]
        [Test]
        public void AlignContinuationsIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignContinuations, viewModel.AlignContinuations);
        }

        [Category("Settings")]
        [Test]
        public void AlignDimColumnIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDimColumn, viewModel.AlignDimColumn);
        }

        [Category("Settings")]
        [Test]
        public void AlignDimsIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDims, viewModel.AlignDims);
        }

        [Category("Settings")]
        [Test]
        public void EndOfLineCommentColumnSpaceAlignmentIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, viewModel.EndOfLineCommentColumnSpaceAlignment);
        }

        [Category("Settings")]
        [Test]
        public void EndOfLineCommentStyleIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle, viewModel.EndOfLineCommentStyle);
        }

        [Category("Settings")]
        [Test]
        public void ForceCompilerDirectivesInColumn1IsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, viewModel.ForceCompilerDirectivesInColumn1);
        }

        [Category("Settings")]
        [Test]
        public void ForceDebugStatementsInColumn1IsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, viewModel.ForceDebugStatementsInColumn1);
        }

        [Category("Settings")]
        [Test]
        public void IgnoreOperatorsInContinuationsIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, viewModel.IgnoreOperatorsInContinuations);
        }

        [Category("Settings")]
        [Test]
        public void IndentCaseIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCase, viewModel.IndentCase);
        }

        [Category("Settings")]
        [Test]
        public void IndentEnumTypeAsProcedureIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEnumTypeAsProcedure, viewModel.IndentEnumTypeAsProcedure);
        }

        [Category("Settings")]
        [Test]
        public void IndentCompilerDirectivesIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCompilerDirectives, viewModel.IndentCompilerDirectives);
        }

        [Category("Settings")]
        [Test]
        public void IndentEntireProcedureBodyIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody, viewModel.IndentEntireProcedureBody);
        }

        [Category("Settings")]
        [Test]
        public void IndentFirstCommentBlockIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock, viewModel.IndentFirstCommentBlock);
        }

        [Category("Settings")]
        [Test]
        public void IndentFirstDeclarationBlockIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, viewModel.IndentFirstDeclarationBlock);
        }

        [Category("Settings")]
        [Test]
        public void VerticallySpaceProceduresIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.VerticallySpaceProcedures, viewModel.VerticallySpaceProcedures);
        }

        [Test]
        [Category("Settings")]
        public void LinesBetweenProceduresIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.LinesBetweenProcedures, viewModel.LinesBetweenProcedures);
        }

        [Test]
        [Category("Settings")]
        public void EmptyLineHandlingMethodIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EmptyLineHandlingMethod, viewModel.EmptyLineHandlingMethod);
        }
    }
}
