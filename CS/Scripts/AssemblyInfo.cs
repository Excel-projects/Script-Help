using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using System.Deployment.Application;
using Microsoft.Win32;

namespace ScriptHelp.Scripts
{
    public static class AssemblyInfo
    {
        public static string versionFolderNumber;

        public static string Title
        {
            get
            {
                string result = string.Empty;
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

                if (assembly != null)
                {
                    object[] customAttributes = assembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                    if ((customAttributes != null) && (customAttributes.Length > 0))
                        result = ((AssemblyTitleAttribute)customAttributes[0]).Title;
                }

                return result;
            }
        }

        public static string Description
        {
            get
            {
                string result = string.Empty;
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

                if (assembly != null)
                {
                    object[] customAttributes = assembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                    if ((customAttributes != null) && (customAttributes.Length > 0))
                        result = ((AssemblyDescriptionAttribute)customAttributes[0]).Description;
                }
                return result;
            }
        }

        public static string Company
        {
            get
            {
                string result = string.Empty;
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

                if (assembly != null)
                {
                    object[] customAttributes = assembly.GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                    if ((customAttributes != null) && (customAttributes.Length > 0))
                        result = ((AssemblyCompanyAttribute)customAttributes[0]).Company;
                }

                return result;
            }
        }

        public static string Product
        {
            get
            {
                string result = string.Empty;
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

                if (assembly != null)
                {
                    object[] customAttributes = assembly.GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                    if ((customAttributes != null) && (customAttributes.Length > 0))
                        result = ((AssemblyProductAttribute)customAttributes[0]).Product;
                }
                return result;
            }
        }

        public static string Copyright
        {
            get
            {
                string result = string.Empty;
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

                if (assembly != null)
                {
                    object[] customAttributes = assembly.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                    if ((customAttributes != null) && (customAttributes.Length > 0))
                        result = ((AssemblyCopyrightAttribute)customAttributes[0]).Copyright;
                }
                return result;
            }
        }

        public static string Trademark
        {
            get
            {
                string result = string.Empty;
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

                if (assembly != null)
                {
                    object[] customAttributes = assembly.GetCustomAttributes(typeof(AssemblyTrademarkAttribute), false);
                    if ((customAttributes != null) && (customAttributes.Length > 0))
                        result = ((AssemblyTrademarkAttribute)customAttributes[0]).Trademark;
                }
                return result;
            }
        }

        public static string AssemblyVersion
        {
            get
            {
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                return assembly.GetName().Version.ToString();
            }
        }

        public static string FileVersion
        {
            get
            {
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                return fvi.FileVersion;
            }
        }

        public static string Guid
        {
            get
            {
                string result = string.Empty;
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

                if (assembly != null)
                {
                    object[] customAttributes = assembly.GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), false);
                    if ((customAttributes != null) && (customAttributes.Length > 0))
                        result = ((System.Runtime.InteropServices.GuidAttribute)customAttributes[0]).Value;
                }
                return result;
            }
        }

        public static string FileName
        {
            get
            {
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                return fvi.OriginalFilename;
            }
        }

        public static string FilePath
        {
            get
            {
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                return fvi.FileName;
            }
        }

        public static string GetCurrentFileName()
        {
            try
            {
                return Globals.ThisAddIn.Application.ActiveWorkbook.Path + @"\" + Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        public static string GetClickOnceLocation()
        {
            try
            {
                System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
                Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
                return Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());

            }
            catch (Exception)
            {
                return string.Empty;

            }

        }

        public static string GetAssemblyLocation()
        {
            try
            {
                System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
                return assemblyInfo.Location;

            }
            catch (Exception)
            {
                return string.Empty;

            }

        }

        public static void SetAssemblyFolderVersion()
        {
            try
            {
                if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
                {
                    Version ver = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
                    string versionNumber = string.Format("{0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision);
                    versionFolderNumber = "_" + versionNumber.Replace(".", "_");
                }
                else
                {
                    versionFolderNumber = "_" + FileVersion.Replace(".", "_");
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

        public static void SetAddRemoveProgramsIcon(string iconName)
        {
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed
                 && ApplicationDeployment.CurrentDeployment.IsFirstRun)
            {
                try
                {
                    Assembly code = Assembly.GetExecutingAssembly();
                    AssemblyDescriptionAttribute asdescription =
                        (AssemblyDescriptionAttribute)Attribute.GetCustomAttribute(code, typeof(AssemblyDescriptionAttribute));
                    string assemblyDescription = asdescription.Description;

                    //Get the assembly information
                    System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

                    //CodeBase is the location of the ClickOnce deployment files
                    Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
                    string clickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());

                    //the icon is included in this program
                    string iconSourcePath = Path.Combine(clickOnceLocation, @"Resources\" + iconName);
                    if (!File.Exists(iconSourcePath))
                        return;

                    RegistryKey myUninstallKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall");
                    string[] mySubKeyNames = myUninstallKey.GetSubKeyNames();
                    for (int i = 0; i < mySubKeyNames.Length; i++)
                    {
                        RegistryKey myKey = myUninstallKey.OpenSubKey(mySubKeyNames[i], true);
                        object myValue = myKey.GetValue("DisplayName");
                        if (myValue != null && myValue.ToString() == assemblyDescription)
                        {
                            myKey.SetValue("DisplayIcon", iconSourcePath);
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.DisplayMessage(ex);
                }
            }
        }

        public static void OpenFile(string filePath)
        {
            try
            {
                if (filePath == string.Empty)
                    return;
                var attributes = File.GetAttributes(filePath);
                File.SetAttributes(filePath, attributes | FileAttributes.ReadOnly);
                System.Diagnostics.Process.Start(filePath);

            }
            catch (System.ComponentModel.Win32Exception)
            {
                MessageBox.Show("No application is assicated to this file type." + Environment.NewLine + Environment.NewLine + filePath, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }

        }

    }
}