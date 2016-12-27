using ScalesData.Properties;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Permissions;
using System.Security.Principal;
using System.Threading.Tasks;

namespace ScalesData
{
    [RunInstaller(true)]
    public partial class LibInstaller : Installer
    {
        public LibInstaller()
        {
            InitializeComponent();
        }

        [SecurityPermission(SecurityAction.Demand)]
        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);
        }

        /// <summary>
        /// Runs regasm /tlb /codebase to create and register the TLB file.
        /// </summary>
        [SecurityPermission(SecurityAction.Demand)]
        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);

            // Get the location of regasm.
            string regasmPath = RuntimeEnvironment.GetRuntimeDirectory() + @"regasm.exe";

            // Get the location of our DLL.
            string componentPath = typeof(LibInstaller).Assembly.Location;

            // Execute regasm.
            Process p = new Process();
            p.StartInfo.FileName = regasmPath;
            p.StartInfo.Arguments = "\"" + componentPath + "\" /tlb /codebase";
            p.StartInfo.Verb = "runas";          // To run as administrator.
            p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            p.Start();
            //MessageBoxEx.Show(string.Format("Install committed with {0}, location {1}", regasmPath, componentPath));

            // Setup permission for dbfile
            DirectoryInfo dInfo = new DirectoryInfo(Path.GetDirectoryName(componentPath));
            DirectorySecurity dSecurity = dInfo.GetAccessControl();
            dSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), FileSystemRights.FullControl, InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit, PropagationFlags.NoPropagateInherit, AccessControlType.Allow));
            dInfo.SetAccessControl(dSecurity);

            p.WaitForExit();
        }

        [SecurityPermission(SecurityAction.Demand)]
        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);
        }

        /// <summary>
        /// Runs regasm /u /tlb /codebase to un-register and delete the TLB file.
        /// </summary>
        [SecurityPermission(SecurityAction.Demand)]
        public override void Uninstall(IDictionary savedState)
        {
            // Get the location of regasm.
            string regasmPath = RuntimeEnvironment.GetRuntimeDirectory() + @"regasm.exe";

            // Get the location of our DLL.
            string componentPath = typeof(LibInstaller).Assembly.Location;

            // Execute regasm.
            Process p = new Process();

            p.StartInfo.FileName = regasmPath;
            p.StartInfo.Arguments = "/u \"" + componentPath + "\" /tlb /codebase";
            p.StartInfo.Verb = "runas";          // To run as administrator.
            p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            p.Start();
            p.WaitForExit();

            // Delete the TLB file.
            File.Delete(Path.Combine(Path.GetDirectoryName(componentPath), Path.GetFileNameWithoutExtension(componentPath) + ".tlb"));

            base.Uninstall(savedState);
        }
    }
}
