// Version 1.0
// Written by Jeremy Saunders (jeremy@jhouseconsulting.com) 29th August 2021
// Modified by Jeremy Saunders (jeremy@jhouseconsulting.com) 12th November 2021
//
using System;
// Required for a List
using System.Collections.Generic;
// Required for NameValueCollection
using System.Collections.Specialized;
// Required to get the appSettings from the App.config
using System.Configuration;
// Required to start process
using System.Diagnostics;
// Required for Linq queries
using System.Linq;
// Required for the ManagementObject LastBootUpTime method 
using System.Management;
// Requires for Stringbuilder
using System.Text;
// Required for regular expressions
using System.Text.RegularExpressions;
// Required for the Timer
using System.Timers;
// Required to Get Services
using System.ServiceProcess;
// Required for file system watcher and log file
using System.IO;
// Required for file system watcher and log file
using System.Reflection;
// Required for registry access
using Microsoft.Win32;
// Required for the Win32Exception class for interpreting Win32 errors
using System.ComponentModel;

namespace ComputerRestartService
{
    public partial class Service1 : ServiceBase
    {

        // Static class to hold global variables.
        public static class Globals
        {
            private static bool _useSettingsFromRegistry;
            public static bool UseSettingsFromRegistry { get { return _useSettingsFromRegistry; } set { _useSettingsFromRegistry = value; } }

            private static int _maxUptimeInDays;
            public static int MaxUptimeInDays { get { return _maxUptimeInDays; } set { _maxUptimeInDays = value; } }

            private static int _repeatTimerInMilliseonds;
            public static int RepeatTimerInMilliseonds { get { return _repeatTimerInMilliseonds; } set { _repeatTimerInMilliseonds = value; } }

            private static int _delayBeforeRestartingInSeconds;
            public static int DelayBeforeRestartingInSeconds { get { return _delayBeforeRestartingInSeconds; } set { _delayBeforeRestartingInSeconds = value; } }

            private static bool _forceRestart;
            public static bool ForceRestart { get { return _forceRestart; } set { _forceRestart = value; } }

            private static bool _restartAfterLogoff;
            public static bool RestartAfterLogoff { get { return _restartAfterLogoff; } set { _restartAfterLogoff = value; } }

            private static bool _checkIfServiceIsRunning;
            public static bool CheckIfServiceIsRunning { get { return _checkIfServiceIsRunning; } set { _checkIfServiceIsRunning = value; } }

            private static string _serviceNameToCheck;
            public static string ServiceNameToCheck { get { return _serviceNameToCheck; } set { _serviceNameToCheck = value; } }

            private static bool _debugToEventLog;
            public static bool DebugToEventLog { get { return _debugToEventLog; } set { _debugToEventLog = value; } }

            private static bool _debugToFile;
            public static bool DebugToFile { get { return _debugToFile; } set { _debugToFile = value; } }

            private static DateTime _fromThisTime;
            public static DateTime FromThisTime { get { return _fromThisTime; } set { _fromThisTime = value; } }

            public static string logfile = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\" + Assembly.GetEntryAssembly().GetName().Name + ".log";

        }

        // A method that reads the appSettings from the registry and updates the global variables.

        public static string GetConfigValueFromRegistry(string valueName)
        {
            StringBuilder stringBuilder = new StringBuilder();
            string regPolicyPath = @"SOFTWARE\Policies\Jeremy Saunders\ComputerRestartService";
            string regPreferencePath = @"SOFTWARE\Jeremy Saunders\ComputerRestartService";
            string output = string.Empty;
            bool valuefound = false;
            try
            {
                RegistryKey subKey1 = Registry.LocalMachine.OpenSubKey(regPolicyPath, false);
                if (subKey1 != null)
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: Checking for the '" + subKey1.ToString() + @"\" + valueName + "' under the policies key with a value of " + output);
                    output = (subKey1.GetValue(valueName).ToString());
                    if (!string.IsNullOrEmpty(output))
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: " + valueName + " found under the policies key with a value of " + output);
                        valuefound = true;
                    }
                    subKey1.Close();
                }
            }
            catch (Exception ex)
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: " + valueName + " not found under the policies key: " + ex.Message.ToString());
            }
            if (!valuefound)
            {
                try
                {
                    RegistryKey subKey2 = Registry.LocalMachine.OpenSubKey(regPreferencePath, false);
                    if (subKey2 != null)
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: Checking for the '" + subKey2.ToString() + @"\" + valueName + "' under the policies key with a value of " + output);
                        output = (string)subKey2.GetValue(valueName);
                        if (!string.IsNullOrEmpty(output))
                        {
                            stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: " + valueName + " found under the preferences key with a value of " + output);
                            valuefound = true;
                        }
                        subKey2.Close();
                    }
                }
                catch (Exception ex)
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: " + valueName + " not found under the preferences key: " + ex.Message.ToString());
                }
            }
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
            return output;
        }

        // A method that reads the appSettings from the registry and updates the global variables.
        public static void GetConfigurationSettingsFromRegistry()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: Reading from registry...");

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("MaxUptimeInDays")))
            {
                int.TryParse(GetConfigValueFromRegistry("MaxUptimeInDays"), out int MaxUptimeInDays);
                if (MaxUptimeInDays < 1) { MaxUptimeInDays = 1; }
                Globals.MaxUptimeInDays = MaxUptimeInDays;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - MaxUptimeInDays is set to: " + Globals.MaxUptimeInDays.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - MaxUptimeInDays is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
            }

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("RepeatTimerInMilliseonds")))
            {
                int.TryParse(GetConfigValueFromRegistry("RepeatTimerInMilliseonds"), out int RepeatTimerInMilliseonds);
                if (RepeatTimerInMilliseonds < 1000) { RepeatTimerInMilliseonds = 1000; }
                Globals.RepeatTimerInMilliseonds = RepeatTimerInMilliseonds;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - RepeatTimerInMilliseonds is set to: " + Globals.RepeatTimerInMilliseonds.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - RepeatTimerInMilliseonds is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
            }

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("DelayBeforeRestartingInSeconds")))
            {
                int.TryParse(GetConfigValueFromRegistry("DelayBeforeRestartingInSeconds"), out int DelayBeforeRestartingInSeconds);
                Globals.DelayBeforeRestartingInSeconds = DelayBeforeRestartingInSeconds;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - DelayBeforeRestartingInSeconds is set to: " + Globals.DelayBeforeRestartingInSeconds.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - DelayBeforeRestartingInSeconds is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
            }

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("ForceRestart")))
            {
                bool.TryParse(GetConfigValueFromRegistry("ForceRestart"), out bool ForceRestart);
                Globals.ForceRestart = ForceRestart;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - ForceRestart is set to: " + Globals.ForceRestart.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - ForceRestart is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
            }

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("RestartAfterLogoff")))
            {
                bool.TryParse(GetConfigValueFromRegistry("RestartAfterLogoff"), out bool RestartAfterLogoff);
                Globals.RestartAfterLogoff = RestartAfterLogoff;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - RestartAfterLogoff is set to: " + Globals.RestartAfterLogoff.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - RestartAfterLogoff is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
            }

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("CheckIfServiceIsRunning")))
            {
                bool.TryParse(GetConfigValueFromRegistry("CheckIfServiceIsRunning"), out bool CheckIfServiceIsRunning);
                Globals.CheckIfServiceIsRunning = CheckIfServiceIsRunning;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - CheckIfServiceIsRunning is set to: " + Globals.CheckIfServiceIsRunning.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - CheckIfServiceIsRunning is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
            }

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("ServiceNameToCheck")))
            {
                Globals.ServiceNameToCheck = GetConfigValueFromRegistry("ServiceNameToCheck");
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - ServiceNameToCheck is set to: " + Globals.ServiceNameToCheck);
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - ServiceNameToCheck is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
            }

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("DebugToEventLog")))
            {
                bool.TryParse(GetConfigValueFromRegistry("DebugToEventLog"), out bool DebugToEventLog);
                Globals.DebugToEventLog = DebugToEventLog;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - DebugToEventLog is set to: " + Globals.DebugToEventLog.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - DebugToEventLog is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
            }

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("DebugToFile")))
            {
                bool.TryParse(GetConfigValueFromRegistry("DebugToFile"), out bool DebugToFile);
                Globals.DebugToFile = DebugToFile;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - DebugToFile is set to: " + Globals.DebugToFile.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - DebugToFile is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
            }

            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }

        // A method that reads the appSettings from the App.config and updates the global variables.
        public static void GetConfigurationSettings()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: Reading appSettings...");

            // Always read from the disk to get the latest setting
            ConfigurationManager.RefreshSection("appSettings");

            NameValueCollection appSettings = ConfigurationManager.AppSettings;

            foreach (string s in appSettings.AllKeys)
            {
                if (s == "UseSettingsFromRegistry")
                {
                    bool.TryParse(appSettings.Get(s), out bool UseSettingsFromRegistry);
                    Globals.UseSettingsFromRegistry = UseSettingsFromRegistry;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - UseSettingsFromRegistry is set to: " + Globals.UseSettingsFromRegistry.ToString());
                }
                if (s == "MaxUptimeInDays")
                {
                    int.TryParse(appSettings.Get(s), out int MaxUptimeInDays);
                    if (MaxUptimeInDays < 1) { MaxUptimeInDays = 1; }
                    Globals.MaxUptimeInDays = MaxUptimeInDays;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - MaxUptimeInDays is set to: " + Globals.MaxUptimeInDays.ToString());
                }
                if (s == "RepeatTimerInMilliseonds")
                {
                    int.TryParse(appSettings.Get(s), out int RepeatTimerInMilliseonds);
                    if (RepeatTimerInMilliseonds < 1000) { RepeatTimerInMilliseonds = 1000; }
                    Globals.RepeatTimerInMilliseonds = RepeatTimerInMilliseonds;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - RepeatTimerInMilliseonds is set to: " + Globals.RepeatTimerInMilliseonds.ToString());
                }
                if (s == "DelayBeforeRestartingInSeconds")
                {
                    int.TryParse(appSettings.Get(s), out int DelayBeforeRestartingInSeconds);
                    Globals.DelayBeforeRestartingInSeconds = DelayBeforeRestartingInSeconds;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - DelayBeforeRestartingInSeconds is set to: " + Globals.DelayBeforeRestartingInSeconds.ToString());
                }
                if (s == "ForceRestart")
                {
                    bool.TryParse(appSettings.Get(s), out bool ForceRestart);
                    Globals.ForceRestart = ForceRestart;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - ForceRestart is set to: " + Globals.ForceRestart.ToString());
                }
                if (s == "RestartAfterLogoff")
                {
                    bool.TryParse(appSettings.Get(s), out bool RestartAfterLogoff);
                    Globals.RestartAfterLogoff = RestartAfterLogoff;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - RestartAfterLogoff is set to: " + Globals.RestartAfterLogoff.ToString());
                }
                if (s == "CheckIfServiceIsRunning")
                {
                    bool.TryParse(appSettings.Get(s), out bool CheckIfServiceIsRunning);
                    Globals.CheckIfServiceIsRunning = CheckIfServiceIsRunning;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - CheckIfServiceIsRunning is set to: " + Globals.CheckIfServiceIsRunning.ToString());
                }
                if (s == "ServiceNameToCheck")
                {
                    Globals.ServiceNameToCheck = appSettings.Get(s);
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - ServiceNameToCheck is set to: " + Globals.ServiceNameToCheck);
                }
                if (s == "DebugToEventLog")
                {
                    bool.TryParse(appSettings.Get(s), out bool DebugToEventLog);
                    Globals.DebugToEventLog = DebugToEventLog;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - DebugToEventLog is set to: " + Globals.DebugToEventLog.ToString());
                }
                if (s == "DebugToFile")
                {
                    bool.TryParse(appSettings.Get(s), out bool DebugToFile);
                    Globals.DebugToFile = DebugToFile;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - DebugToFile is set to: " + Globals.DebugToFile.ToString());
                }
            }
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }

        public static bool GetIdleFile(string path)
        {
            var fileIdle = false;
            const int MaximumAttemptsAllowed = 30;
            var attemptsMade = 0;

            while (!fileIdle && attemptsMade <= MaximumAttemptsAllowed)
            {
                try
                {
                    using (File.Open(path, FileMode.Open, FileAccess.ReadWrite))
                    {
                        fileIdle = true;
                    }
                }
                catch
                {
                    attemptsMade++;
                    System.Threading.Thread.Sleep(100);
                }
            }

            return fileIdle;
        }

        private static readonly FileSystemWatcher fileSystemWatcher = new FileSystemWatcher();
        public void WatchForFileChangeEvent()
        {

            // Associate event handlers with the events
            fileSystemWatcher.Changed += FileSystemWatcher_Changed;

            // tell the watcher where to look
            fileSystemWatcher.Path = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            // set the watcher filters
            fileSystemWatcher.NotifyFilter = NotifyFilters.LastWrite;

            // Watch for changes on the .config file
            fileSystemWatcher.Filter = Assembly.GetEntryAssembly().GetName().Name + ".exe.config";

            fileSystemWatcher.IncludeSubdirectories = false;

            // You must add this line - this allows events to fire.
            fileSystemWatcher.EnableRaisingEvents = true;
        }

        public void FileSystemWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed : The " + e.FullPath + " file has changed.");
            try
            {
                fileSystemWatcher.Changed -= FileSystemWatcher_Changed;
                fileSystemWatcher.EnableRaisingEvents = false;
                // Read and update the global variables from the appSettings section of the App.config
                if (GetIdleFile(e.FullPath))
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed : Updating the configuration settings.");
                    GetConfigurationSettings();
                }
                // Read and update the global variables from the registry if UseSettingsFromRegistry is set to true
                if (Globals.UseSettingsFromRegistry)
                {
                    GetConfigurationSettingsFromRegistry();
                }

                // If the actual uptime is less than the MaxUptimeInDays variable, set the timer interval to the difference plus the repeat timer, for some buffer, and ensure the timer is enabled.
                // Else
                // If the RestartAfterLogoff variable is false, set the timer interval to the RepeatTimerInMilliseond and ensure the timer is enabled.
                // else disable the timer because the logoff event will be used to restart this computer.
                if (Convert.ToInt32(GetUptime().TotalMilliseconds) < (Globals.MaxUptimeInDays * 86400000))
                {
                    int interval = ((Globals.MaxUptimeInDays) * 86400000) - Convert.ToInt32(GetUptime().TotalMilliseconds + Globals.RepeatTimerInMilliseonds);
                    _checkUptimeTimer.Interval = interval;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed : - The timer interval has now been updated to check every " + interval.ToString() + " milliseconds.");
                    _checkUptimeTimer.Enabled = true;
                }
                else
                {
                    if (!Globals.RestartAfterLogoff)
                    {
                        // Update the Timer interval. If the interval is set after the Timer has started, the counter is reset.
                        if (_checkUptimeTimer.Interval != Globals.RepeatTimerInMilliseonds)
                        {
                            _checkUptimeTimer.Interval = Globals.RepeatTimerInMilliseonds;
                            stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed : - The timer interval has now been updated to check every " + Globals.RepeatTimerInMilliseonds.ToString() + " milliseconds.");
                        }
                        _checkUptimeTimer.Enabled = true;
                    }
                    else
                    {
                        // Desctivate/Stop the Timer
                        _checkUptimeTimer.Enabled = false;
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed : - The timer has now been disabled because the RestartAfterLogoff variable is set to True. Therefore the logoff event will be used to restart this computer.");
                    }
                }
            }
            catch (Exception exception)
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed : " + exception.Message);
            }
            finally
            {
                fileSystemWatcher.Changed += FileSystemWatcher_Changed;
                fileSystemWatcher.EnableRaisingEvents = true;
            }
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }

        // Get the computer uptime.
        public static TimeSpan GetUptime()
        {
            ManagementObject mo = new ManagementObject(@"\\.\root\cimv2:Win32_OperatingSystem=@");
            DateTime lastBootUp = ManagementDateTimeConverter.ToDateTime(mo["LastBootUpTime"].ToString());
            return DateTime.Now.ToUniversalTime() - lastBootUp.ToUniversalTime();
        }

        // Check if a specified service exists.
        public bool DoesServiceExist(string serviceName)
        {
            bool result = false;
            var service = ServiceController.GetServices().FirstOrDefault(s => (s.ServiceName == serviceName || s.DisplayName == serviceName));
            if (service != null)
            {
                result = true;
            }
            return result;
        }

        // Check if a specified service is running.
        public bool IsServiceRunning(string serviceName)
        {
            bool result = false;
            var service = ServiceController.GetServices().FirstOrDefault(s => (s.ServiceName == serviceName || s.DisplayName == serviceName));
            if (service != null)
            {
                if (service.Status.Equals(ServiceControllerStatus.Running))
                {
                    result = true;
                }
            }
            return result;
        }

        // Restart the computer using unmanaged code
        public static void RestartMachineUnmanagedCode(bool force, int timeout)
        {
            StringBuilder stringBuilder = new StringBuilder();
            _eventLog.WriteEntry("The Computer Restart Service has initiated a restart.", EventLogEntryType.Information, 100);
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": RestartMachineUnmanagedCode: The Computer Restart Service has initiated a restart.");
            string filename = "shutdown.exe";
            string arguments = string.Empty;
            if (!force)
            {
                arguments = "/r /t " + timeout;
            }
            else
            {
                arguments = "/r /t " + timeout + " /f";
            }
            var proc = new ProcessStartInfo(filename, arguments);
            proc.CreateNoWindow = true;
            proc.UseShellExecute = false;
            Process.Start(proc);
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }

        // Restart the computer using managed code
        public static void RestartMachineManagedCode(bool force, int timeout)
        {
            StringBuilder stringBuilder = new StringBuilder();
            _eventLog.WriteEntry("The Computer Restart Service has initiated a restart.", EventLogEntryType.Information, 100);
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": RestartMachineManagedCode: The Computer Restart Service has initiated a restart.");
            ManagementClass managementClass = new ManagementClass("Win32_OperatingSystem");
            managementClass.Get();
            bool EnablePrivileges = managementClass.Scope.Options.EnablePrivileges;
            managementClass.Scope.Options.EnablePrivileges = true;
            ManagementBaseObject methodParameters = managementClass.GetMethodParameters("Win32ShutdownTracker");
            if (!force)
            {
                methodParameters["Flags"] = 2; //Reboot
            }
            else
            {
                methodParameters["Flags"] = 6; //Forced Reboot
            }
            methodParameters["Timeout"] = timeout;
            foreach (ManagementObject instance in managementClass.GetInstances())
            {
                var outParams = instance.InvokeMethod("Win32Shutdown", methodParameters, (InvokeMethodOptions)null);
                int.TryParse(outParams["ReturnValue"].ToString(), out int returnCode);
                if (returnCode != 0)
                {
                    var ex = new Win32Exception(returnCode);
                    _eventLog.WriteEntry("The Computer Restart Service failed to restart the computer: " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim(), EventLogEntryType.Error, 100);
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": RestartMachineManagedCode: The Computer Restart Service failed to restart the computer: " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim());
                }
            }
            managementClass.Scope.Options.EnablePrivileges = EnablePrivileges;
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }

        // The SeesionIDMapping dictionary, which keeps a mapping of the SessionID to Username.
        public static Dictionary<string, string> SeesionIDMapping = new Dictionary<string, string>();

        // The LoggedOnUser class is used by the outpit of the "query user" command.
        public class LoggedOnUsers
        {
            public string Username { get; set; }
            public string SessionName { get; set; }
            public string SessionID { get; set; }
            public string State { get; set; }
            public string IdleTime { get; set; }
            public DateTime LogonTime { get; set; }
        }

        // Run the "query user" command and process the output.
        public List<LoggedOnUsers> GetNumberLoggedOnUserSessions()
        {
            StringBuilder stringBuilder = new StringBuilder();

            String CmdText = @"/c query user";

            Process proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = CmdText,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                }
            };
            proc.Start();
            proc.WaitForExit();
            String stdOut = proc.StandardOutput.ReadToEnd();
            String stdErr = proc.StandardError.ReadToEnd();
            proc.Dispose();

            string computerName = Environment.MachineName.ToString();

            List<LoggedOnUsers> loggedonusers = new List<LoggedOnUsers> { };
            int matchCount = 0;

            string pattern = @"^\s?(?<username>[/^>+/\w\/.\/_\-]{2,64})\s\s+(?<sessionname>\w?.{0,22})\s\s+(?<sessionid>\d{1,5})\s\s+(?<state>.{0,12})\s\s+(?<idletime>.{0,20})\s\s+(?<logontime>.{0,32})$";
            Regex regex = new Regex(pattern, RegexOptions.Multiline);
            Match match = regex.Match(stdOut);
            if (match.Success)
            {
                SeesionIDMapping.Clear();
                while (match.Success)
                {
                    matchCount++;
                    String username = match.Groups["username"].Value.ToString().Trim();
                    if (username.Substring(0, 1) == ">")
                    {
                        username = username.Substring(1, username.Length - 1);
                    }
                    DateTime parsedLogonTime;
                    if (!DateTime.TryParse(match.Groups["logontime"].Value.ToString().Trim(), out parsedLogonTime))
                    {
                        parsedLogonTime = DateTime.MinValue;
                    }
                    loggedonusers.Add(new LoggedOnUsers
                    {
                        Username = username,
                        SessionName = match.Groups["sessionname"].Value.ToString().Trim(),
                        SessionID = match.Groups["sessionid"].Value.ToString().Trim(),
                        State = match.Groups["state"].Value.ToString().Trim(),
                        IdleTime = match.Groups["idletime"].Value.ToString().Trim(),
                        LogonTime = parsedLogonTime
                    });
                    SeesionIDMapping.Add(match.Groups["sessionid"].Value.ToString().Trim(), username);
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions: - Username: " + username);
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - SessionName: " + match.Groups["sessionname"].Value.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - SessionID: " + match.Groups["sessionid"].Value.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - State: " + match.Groups["state"].Value.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - IdleTime: " + match.Groups["idletime"].Value.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - LogonTime: " + parsedLogonTime);
                    match = match.NextMatch();
                }
                if (matchCount == 1)
                {
                    stringBuilder.Insert(0, DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions: Found " + matchCount + " logged on user session:" + Environment.NewLine);
                }
                else
                {
                    stringBuilder.Insert(0, DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions: Found " + matchCount + " logged on user sessions:" + Environment.NewLine);
                }
            }
            else
            {
                if (stdErr.IndexOf("No User exists for", StringComparison.CurrentCultureIgnoreCase) >= 0)
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions: No logged on user sessions were found");

                }
                else
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions: An error occured: " + stdErr);
                }
            }
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
            return loggedonusers;
        }

        public string GetUsernameFromSessionID(string SessionID)
        {
            string username = string.Empty;
            if (SeesionIDMapping.ContainsKey(SessionID))
            {
                SeesionIDMapping.TryGetValue(SessionID, out username);
            }
            return username;
        }

        public bool HasTheMaximumUpTimeBeenReached()
        {
            bool result = false;

            StringBuilder stringBuilder = new StringBuilder();
            string computerName = Environment.MachineName.ToString();

            int actualUptimeInMilliseconds = Convert.ToInt32(GetUptime().TotalMilliseconds);

            stringBuilder.AppendLine(DateTime.Now.ToString() + ": HasTheMaximumUpTimeBeenReached: The " + computerName + " computer has been up for " + actualUptimeInMilliseconds.ToString() + " milliseconds.");

            if (actualUptimeInMilliseconds >= (Globals.MaxUptimeInDays * 86400000))
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": HasTheMaximumUpTimeBeenReached: The " + computerName + " computer needs to be rebooted as the maximum uptime of " + Globals.MaxUptimeInDays + " days has been reached.");
                result = true;
            }

            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
            return result;
        }
        public void WatchForSecurityEvent()
        {
            string machineName = Environment.MachineName;
            string eventLog = "Security";
            EventLog logListener = new EventLog(eventLog, machineName);
            logListener.EnableRaisingEvents = true;
            logListener.EntryWritten += new EntryWrittenEventHandler(OnEntryWritten);
        }
        public void OnEntryWritten(object source, EntryWrittenEventArgs entryArg)
        {
            // Looking for:
            // Event ID 4624 with LogonType of either 2, 10 or 11. This tells us that an account was successfully logged on.
            // Event ID 4634 with LogonType of 3. This tells us that an account was logged off. This event is generated when a logon session is destroyed.
            StringBuilder stringBuilder = new StringBuilder();
            bool EventIDMatch = false;
            string computerName = Environment.MachineName.ToString();
            var timeWritten = entryArg.Entry.TimeWritten;

            // To ensure we are only dealing with new events, and not events from the past, we store the last read datetime
            // and compare with new events from the OnEntryWritten.
            if (timeWritten > Globals.FromThisTime)
            {
                int[] eventids = { 4624, 4634 };
                foreach (int eventid in eventids)
                {
                    if (eventid == entryArg.Entry.InstanceId)
                    {

                        Match logonType = Regex.Match(entryArg.Entry.Message, @"Logon Type:(.*)");
                        if (logonType.Success)
                        {
                            int type = Convert.ToInt32(logonType.Groups[1].Value.Trim());
                            if (type == 2 || type == 3 || type == 10 || type == 11)
                            {
                                string username = string.Empty;
                                Match accountName = Regex.Match(entryArg.Entry.Message, @"Account Name:(.*)");
                                if (accountName.Success)
                                {
                                    username = accountName.Groups[1].Value.Trim();
                                    if (entryArg.Entry.InstanceId == 4624)
                                    {
                                        accountName = accountName.NextMatch();
                                        username = accountName.Groups[1].Value.Trim();
                                    }
                                }
                                if (username.Substring(username.Length - 1, 1) != "$" && username.Substring(0, 4) != "DWM-" && username != "-")
                                {
                                    if (entryArg.Entry.InstanceId == 4624 && (type == 2 || type == 10 || type == 11))
                                    {
                                        EventIDMatch = true;
                                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnEntryWritten: " + username + " just logged on to " + computerName + ".");
                                    }
                                    if (entryArg.Entry.InstanceId == 4634 && type == 3)
                                    {
                                        EventIDMatch = true;
                                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnEntryWritten: " + username + " just logged off from " + computerName + ".");
                                    }
                                }
                            }
                            if (EventIDMatch)
                            {
                                Globals.FromThisTime = timeWritten;
                            }
                            //if (Globals.DebugToEventLog & EventIDMatch) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
                            //if (Globals.DebugToFile & EventIDMatch) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
                        }
                    }
                }
            }
        }

        // Setup the Timer object.
        private static Timer _checkUptimeTimer = new Timer();

        // Setup the Event Log object.
        private static EventLog _eventLog = new EventLog();

        public Service1()
        {
            CanHandlePowerEvent = true;
            CanHandleSessionChangeEvent = true;

            AutoLog = false;
            string source = "ComputerRestartService";
            string log = "Application";
            if (!EventLog.SourceExists(source))
            {
                EventLog.CreateEventSource(source, log);
            }
            _eventLog.Source = source;
            _eventLog.Log = log;

            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            StringBuilder stringBuilder = new StringBuilder();
            string computerName = Environment.MachineName.ToString();
            _eventLog.WriteEntry("The Computer Restart Service, written by Jeremy Saunders, has been started.", EventLogEntryType.Information, 100);
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStart: The Computer Restart Service, written by Jeremy Saunders, has been started on " + computerName + ".");
            File.AppendAllText(Globals.logfile, stringBuilder.ToString());
            stringBuilder.Clear();

            // Set the Start Time of the service, which is used to help filter the events from the EntryWrittenEventHandler.
            Globals.FromThisTime = DateTime.Now;

            // Read and update the global variables from the appSettings section of the App.config and registry
            GetConfigurationSettings();
            if (Globals.UseSettingsFromRegistry)
            {
                GetConfigurationSettingsFromRegistry();
            }

            // Create a list of logged on users at service start
            GetNumberLoggedOnUserSessions();

            // Setup and call the file change event watcher
            WatchForFileChangeEvent();

            // Setup and call the security event watcher
            WatchForSecurityEvent();

            // Call and set the computer uptime
            int actualUptimeInMilliseconds = Convert.ToInt32(GetUptime().TotalMilliseconds);
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStart: Computer uptime in milliseconds: " + actualUptimeInMilliseconds.ToString());

            // Setup and call the Timer
            _checkUptimeTimer.Elapsed += new ElapsedEventHandler(CheckUptimeTimerElapsed);
            _checkUptimeTimer.AutoReset = true;
            // If the actual uptime is less than or equal to the MaxUptimeInDays variable, set the timer interval to the difference in milliseonds.
            // We do this to avoid unnecessarily checking the MaxUptime until it has actually been reached.If it has been reached,
            // we set it to an interval as specified in the App.Config file.
            // 1 day = 86400000 milliseonds
            int interval = Globals.RepeatTimerInMilliseonds;
            if (actualUptimeInMilliseconds <= (Globals.MaxUptimeInDays * 86400000))
            {
                interval = (Globals.MaxUptimeInDays * 86400000) - actualUptimeInMilliseconds + Globals.RepeatTimerInMilliseonds;
            }
            _checkUptimeTimer.Interval = interval;
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStart: Setting the timer interval to " + interval.ToString() + " milliseconds.");

            // Activate/Start the Timer
            _checkUptimeTimer.Enabled = true;

            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }

        /// <summary>
        /// OnSessionChange(): To handle a session change event. Useful if you need to
        /// determine when a user logs on or off either remotely or on the console.
        /// </summary>
        /// <param name="changeDescription">The Session Change Event that occured.</param>
        protected override void OnSessionChange(SessionChangeDescription changeDescription)
        {
            StringBuilder stringBuilder = new StringBuilder();
            string username = string.Empty;
            string computerName = Environment.MachineName.ToString();
            switch (changeDescription.Reason)
            {
                case SessionChangeReason.SessionLogon:
                    // Refresh the list of logged on users after a logon has occurred.
                    GetNumberLoggedOnUserSessions();
                    username = GetUsernameFromSessionID(changeDescription.SessionId.ToString());
                    if (string.IsNullOrEmpty(username)) { username = "A user"; }
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: " + username + " (Session ID " + changeDescription.SessionId.ToString() + ") has logged on to a session");
                    break;
                case SessionChangeReason.SessionLogoff:
                    username = GetUsernameFromSessionID(changeDescription.SessionId.ToString());
                    if (string.IsNullOrEmpty(username)) { username = "A user"; }
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: " + username + " (Session ID " + changeDescription.SessionId.ToString() + ") has logged off from a session");
                    if (Globals.RestartAfterLogoff)
                    {
                        // Refresh the list of logged on users after a logoff has occurred. If no more users are logged in, it can be rebooted.
                        if (GetNumberLoggedOnUserSessions().Count == 0)
                        {
                            bool OkayToRestart = false;
                            if (Globals.CheckIfServiceIsRunning)
                            {
                                if (DoesServiceExist(Globals.ServiceNameToCheck))
                                {
                                    if (IsServiceRunning(Globals.ServiceNameToCheck))
                                    {
                                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: The " + Globals.ServiceNameToCheck + " service is running on " + computerName + ".");
                                        OkayToRestart = true;
                                    }
                                    else
                                    {
                                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: The " + Globals.ServiceNameToCheck + " service is not running on " + computerName + ".");
                                        OkayToRestart = false;
                                    }
                                }
                                else
                                {
                                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: The " + Globals.ServiceNameToCheck + " service does not exist on " + computerName + ".");
                                    OkayToRestart = true;
                                }
                            }
                            if (OkayToRestart)
                            {
                                stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: Okay to restart.");
                                RestartMachineManagedCode(Globals.ForceRestart, Globals.DelayBeforeRestartingInSeconds);
                            }
                            else
                            {
                                stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: It cannot be restarted at this point.");
                            }
                        }
                    }
                    break;
                case SessionChangeReason.RemoteConnect:
                    username = GetUsernameFromSessionID(changeDescription.SessionId.ToString());
                    if (string.IsNullOrEmpty(username)) { username = "A user"; }
                    //stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: " + username + " (Session ID " + changeDescription.SessionId.ToString() + ") connected to a remote session");
                    break;
                case SessionChangeReason.RemoteDisconnect:
                    username = GetUsernameFromSessionID(changeDescription.SessionId.ToString());
                    if (string.IsNullOrEmpty(username)) { username = "A user"; }
                    //stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: " + username + " (Session ID " + changeDescription.SessionId.ToString() + ") disconnected from a remote session");
                    break;
                case SessionChangeReason.SessionRemoteControl:
                    username = GetUsernameFromSessionID(changeDescription.SessionId.ToString());
                    if (string.IsNullOrEmpty(username)) { username = "A user"; }
                    //stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: " + username + " (Session ID " + changeDescription.SessionId.ToString() + ") remote control status has changed.");
                    break;
                case SessionChangeReason.ConsoleConnect:
                    username = GetUsernameFromSessionID(changeDescription.SessionId.ToString());
                    if (string.IsNullOrEmpty(username)) { username = "A user"; }
                    //stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: " + username + " (Session ID " + changeDescription.SessionId.ToString() + ") connected to the console session");
                    break;
                case SessionChangeReason.ConsoleDisconnect:
                    username = GetUsernameFromSessionID(changeDescription.SessionId.ToString());
                    if (string.IsNullOrEmpty(username)) { username = "A user"; }
                    //stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: " + username + " (Session ID " + changeDescription.SessionId.ToString() + ") disconnected from the console session");
                    break;
                case SessionChangeReason.SessionLock:
                    username = GetUsernameFromSessionID(changeDescription.SessionId.ToString());
                    if (string.IsNullOrEmpty(username)) { username = "A user"; }
                    //stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: " + username + " (Session ID " + changeDescription.SessionId.ToString() + ") locked their session");
                    break;
                case SessionChangeReason.SessionUnlock:
                    username = GetUsernameFromSessionID(changeDescription.SessionId.ToString());
                    if (string.IsNullOrEmpty(username)) { username = "A user"; }
                    //stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: " + username + " (Session ID " + changeDescription.SessionId.ToString() + ") unlocked their session");
                    break;
                default:
                    username = GetUsernameFromSessionID(changeDescription.SessionId.ToString());
                    if (string.IsNullOrEmpty(username)) { username = "A user"; }
                    //stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: " + username + " (Session ID " + changeDescription.SessionId.ToString() + ") has initiated an unknown action");
                    break;
            }
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }

        protected override void OnStop()
        {
            _eventLog.WriteEntry("The Computer Restart Service, written by Jeremy Saunders, has been stopped.", EventLogEntryType.Information, 100);

            StringBuilder stringBuilder = new StringBuilder();
            string computerName = Environment.MachineName.ToString();
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStop: The Computer Restart Service, written by Jeremy Saunders, has been stopped on " + computerName + ".");
            File.AppendAllText(Globals.logfile, stringBuilder.ToString());

            _checkUptimeTimer.Stop();
        }

        public void CheckUptimeTimerElapsed(object sender, ElapsedEventArgs e)
        {
            StringBuilder stringBuilder = new StringBuilder();
            string computerName = Environment.MachineName.ToString();

            if (HasTheMaximumUpTimeBeenReached())
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed: The maximum uptime of " + Globals.MaxUptimeInDays.ToString() + " days has been reached on " + computerName + ".");
                int CurrentLoggedOnUserCount = GetNumberLoggedOnUserSessions().Count;
                if (CurrentLoggedOnUserCount == 0)
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed: - No users are currently logged in.");
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed: - Restarting...");
                    RestartMachineManagedCode(Globals.ForceRestart, Globals.DelayBeforeRestartingInSeconds);
                }
                else
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - " + CurrentLoggedOnUserCount + " users are still logged in.");
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - It cannot be restarted at this point.");

                }
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed: " + computerName + " has not yet reached its maximum uptime of " + Globals.MaxUptimeInDays.ToString() + " days.");
            }
            if (!Globals.RestartAfterLogoff)
            {
                // Update the Timer interval. If the interval is set after the Timer has started, the counter is reset.
                if (_checkUptimeTimer.Interval != Globals.RepeatTimerInMilliseonds)
                {
                    _checkUptimeTimer.Interval = Globals.RepeatTimerInMilliseonds;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - The timer interval has now been updated to check every " + Globals.RepeatTimerInMilliseonds.ToString() + " milliseconds.");
                }
            }
            else
            {
                // Desctivate/Stop the Timer
                _checkUptimeTimer.Enabled = false;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - The timer has now been disabled because the RestartAfterLogoff variable is set to True. Therefore the logoff event will restart this computer.");
            }
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }
    }
}
