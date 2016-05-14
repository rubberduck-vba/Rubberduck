﻿using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public interface ICodeInspectionConfigProvider
    {
        //CodeInspectionConfig Create();
        CodeInspectionSettings Create(IEnumerable<IInspection> inspections);
        CodeInspectionSettings CreateDefaults();
        void Save(CodeInspectionSettings settings);
        event EventHandler LanguageChanged;
    }

    public class CodeInspectionConfigProvider : ICodeInspectionConfigProvider
    {
        private readonly IPersistanceService<CodeInspectionSettings> _persister;
        private IEnumerable<IInspection> _inspections;

        public CodeInspectionConfigProvider(IPersistanceService<CodeInspectionSettings> persister)
        {
            _persister = persister;
        }

        //public CodeInspectionConfig Create()
        //{
        //    //var prototype = new CodeInspectionConfig(_inspections);
        //    //return _persister.Load(prototype) ?? prototype;
        //    return null;
        //}

        public CodeInspectionSettings Create(IEnumerable<IInspection> inspections)
        {
            if (inspections == null) return null;

            _inspections = inspections;
            var prototype = new CodeInspectionSettings(GetDefaultCodeInspections());
            return _persister.Load(prototype) ?? prototype;
        }

        public CodeInspectionSettings CreateDefaults()
        {
            //This sucks.
            return _inspections != null ? new CodeInspectionSettings(GetDefaultCodeInspections()) : null;
        }

        public void Save(CodeInspectionSettings settings)
        {
            _persister.Save(settings);
        }

        public HashSet<CodeInspectionSetting> GetDefaultCodeInspections()
        {
            return new HashSet<CodeInspectionSetting>(_inspections.Select(x =>
                        new CodeInspectionSetting(x.Name, x.Description, x.InspectionType, x.DefaultSeverity,
                            x.DefaultSeverity)));
        }

        private void ApplicationLanguageChanged(object sender, EventArgs e)
        {
            OnLanguageChanged(e);
        }

        public event EventHandler LanguageChanged;
        protected virtual void OnLanguageChanged(EventArgs e)
        {
            var handler = LanguageChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }
}
