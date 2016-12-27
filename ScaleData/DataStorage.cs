using ScalesData.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace ScalesData
{
    class DataStorage
    {
        private static string[] sPath = new string[(int) DataFile.Log + 1];
        private static string[] sName = new string[(int) DataFile.Log + 1];

        static DataStorage()
        {
            sName[(int)DataFile.Database] = "NBFactory.sdf";
            sName[(int)DataFile.Log] = "NB_Log.txt";

            if (String.IsNullOrEmpty(Settings.Default.Destination))
            {
                Settings.Default.Destination = Environment
                    .GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Settings.Default.Save();
            }
            if (string.IsNullOrEmpty(Settings.Default.DataPath))
            {
                Settings.Default.DataPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                Settings.Default.Save();
            }

            updatePath(Settings.Default.DataPath);
            Settings.Default.SettingChanging += settingChanged;
            Settings.Default.SettingsLoaded += Default_SettingsLoaded;
        }

        private static void Default_SettingsLoaded(object sender, System.Configuration.SettingsLoadedEventArgs e)
        {
            updatePath(Settings.Default.DataPath);
        }

        public static void settingChanged(object sender, System.Configuration.SettingChangingEventArgs e)
        {
            if (e.SettingName == "DataPath"
                && !Settings.Default.DataPath.Equals(e.NewValue))
            {
                Launcher launcher = new Launcher();

                lock (Launcher.getConnection())
                {
                    launcher.close();
                    updatePath(e.NewValue.ToString());
                    launcher.open();
                }
            }
        }

        public static bool validateDataPath(string path, out string message)
        {
            message = null;
            if (!Directory.Exists(path))
            {
                message = "File doesn't exists.";
                return false;
            }
            if ((File.GetAttributes(path) & FileAttributes.Directory) != FileAttributes.Directory)
            {
                message = "File isn't a directory.";
                return false;
            }

            string dbPath = path + "\\" + getFileName(DataFile.Database);

            if (!File.Exists(dbPath))
            {
                message = string.Format("Db file {0} doesn't exists.", getFileName(DataFile.Database));
                return false;
            }
            if (!HasWritePermissionOnFile(dbPath))
            {
                message = "Can't write to database file. Check permission.";
                return false;
            }

            return true;
        }

        public static void updatePath(string path)
        {
            Logger.Log("Data folder changed to: " + path);

            for (int i = 0; i < sName.Length; i++)
            {
                sPath[i] = path + "\\" + sName[i];
            }
        }

        public static bool HasWritePermissionOnFile(string path)
        {
            var writeAllow = false;
            var writeDeny = false;
            var accessControlList = File.GetAccessControl(path);
            if (accessControlList == null)
                return false;
            var accessRules = accessControlList.GetAccessRules(true, true,
                                        typeof(System.Security.Principal.SecurityIdentifier));
            if (accessRules == null)
                return false;

            foreach (FileSystemAccessRule rule in accessRules)
            {
                if ((FileSystemRights.Write & rule.FileSystemRights) != FileSystemRights.Write)
                    continue;

                if (rule.AccessControlType == AccessControlType.Allow)
                    writeAllow = true;
                else if (rule.AccessControlType == AccessControlType.Deny)
                    writeDeny = true;
            }

            return writeAllow && !writeDeny && !(new FileInfo(path).IsReadOnly);
        }

        public static string getFileName(DataFile file)
        {
            return sName[(int)file];
        }

        public static string getFilePath(DataFile file)
        {
            return sPath[(int)file];
        }
    }

    public enum DataFile
    {
        Database = 0,
        Log = 1
    }
}
