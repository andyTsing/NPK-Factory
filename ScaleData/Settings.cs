using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace ScalesData.Properties {
    
    
    // This class allows you to handle specific events on the settings class:
    //  The SettingChanging event is raised before a setting's value is changed.
    //  The PropertyChanged event is raised after a setting's value is changed.
    //  The SettingsLoaded event is raised after the setting values are loaded.
    //  The SettingsSaving event is raised before the setting values are saved.
    internal sealed partial class Settings {

        private List<String> mList;
        private Configuration mConfig;

        public Settings()
        {
            // // To add event handlers for saving and changing settings, uncomment the lines below:
            //
            // this.SettingChanging += this.SettingChangingEventHandler;
            //
            // this.SettingsSaving += this.SettingsSavingEventHandler;
            //
            mList = new List<string>();
            string codebase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            Uri p = new Uri(codebase);
            string localPath = p.LocalPath;
            string executingFilename = System.IO.Path.GetFileNameWithoutExtension(localPath);
            string sectionGroupName = "userSettings";
            string sectionName = executingFilename + ".Properties.Settings";
            string configName = "NBFactory.config";
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
            fileMap.ExeConfigFilename = Path.Combine(Path.GetDirectoryName(localPath), configName);
            mConfig = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);
            mList.AddRange(mConfig.AppSettings.Settings.AllKeys);

            // read section of properties
            //var sectionGroup = mConfig.GetSectionGroup(sectionGroupName);
            //var settingsSection = (ClientSettingsSection)sectionGroup.Sections[sectionName];
            //mList = settingsSection.Settings.OfType<ConfigurationElement>().ToList();

            // read section of Connectionstrings
            //var sections = config.Sections.OfType<ConfigurationSection>();
            //var connSection = (from section in sections
            //                   where section.GetType() == typeof(ConnectionStringsSection)
            //                   select section).FirstOrDefault() as ConnectionStringsSection;
            //if (connSection != null)
            //{
            //    mList.AddRange(connSection.ConnectionStrings.Cast<ConfigurationElement>());
            //}

            //config.Save();
        }

        public override object this[string propertyName]
        {
            get
            {
                //var result = (from item in mList
                //              where Convert.ToString(item.ElementInformation.Properties["name"].Value) == propertyName
                //              select item).FirstOrDefault();
                //if (result != null)
                //{
                //    if (result.ElementInformation.Type == typeof(ConnectionStringSettings))
                //    {
                //        return result.ElementInformation.Properties["connectionString"].Value;
                //    }
                //    else if (result.ElementInformation.Type == typeof(SettingElement))
                //    {
                //        return ((SettingValueElement) result.ElementInformation.Properties["value"].Value).ValueXml.InnerText;
                //    }
                //}
                if (mList.Contains(propertyName))
                {
                    return mConfig.AppSettings.Settings[propertyName].Value;
                }
                return base[propertyName];
            }

            set
            {
                if (mList.Contains(propertyName))
                {
                    mConfig.AppSettings.Settings[propertyName].Value = value.ToString();
                }
                base[propertyName] = value;
            }
        }

        public override void Save()
        {
            base.Save();
            mConfig.Save(ConfigurationSaveMode.Modified);
        }

        private void SettingChangingEventHandler(object sender, System.Configuration.SettingChangingEventArgs e) {
            // Add code to handle the SettingChangingEvent event here.
        }
        
        private void SettingsSavingEventHandler(object sender, System.ComponentModel.CancelEventArgs e) {
            // Add code to handle the SettingsSaving event here.
        }
    }
}
