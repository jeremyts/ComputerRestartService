// Version 1.1
// Written by Jeremy Saunders (jeremy@jhouseconsulting.com) 29th August 2021
// Modified by Jeremy Saunders (jeremy@jhouseconsulting.com) 28th December 2023
//
// Note this this code contains two functions that provides the same output. GetNumberLoggedOnUserSessions() and GetNumberLoggedOnUserSessions2()
// I am using GetNumberLoggedOnUserSessions() after feedback from Remko Weijnen and Guy Leech convinced me that I should be using the WTS API for
// reliability and accuracy. The query user (or quser) output lacks granularity for the logon time down to seconds, the session name is often
// truncated on AVD. The WTS API gives way more information and is language independent. Due to the ammount of time and effort I put into the
// GetNumberLoggedOnUserSessions2() function, I have left it in the code for future reference. It was converted from F# code originally posted by
// "epicTurk" on the Code Project site here: https://www.codeproject.com/Tips/1160965/Parse-quser-exe-Results-with-Regex
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
// Required for the DllImport
using System.Runtime.InteropServices;

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

            private static int _inactivityTimerInMinutes;
            public static int InactivityTimerInMinutes { get { return _inactivityTimerInMinutes; } set { _inactivityTimerInMinutes = value; } }

            private static int _disconnectTimerInMinutes;
            public static int DisconnectTimerInMinutes { get { return _disconnectTimerInMinutes; } set { _disconnectTimerInMinutes = value; } }

            private static int _repeatTimerInMilliseonds;
            public static int RepeatTimerInMilliseonds { get { return _repeatTimerInMilliseonds; } set { _repeatTimerInMilliseonds = value; } }

            private static int _delayBeforeRestartingInSeconds;
            public static int DelayBeforeRestartingInSeconds { get { return _delayBeforeRestartingInSeconds; } set { _delayBeforeRestartingInSeconds = value; } }

            private static bool _forceRestart;
            public static bool ForceRestart { get { return _forceRestart; } set { _forceRestart = value; } }

            private static bool _restartAfterLogoff;
            public static bool RestartAfterLogoff { get { return _restartAfterLogoff; } set { _restartAfterLogoff = value; } }

            private static bool _shutdownAfterLogoff;
            public static bool ShutdownAfterLogoff { get { return _shutdownAfterLogoff; } set { _shutdownAfterLogoff = value; } }

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

            public static string regPolicyPath = @"SOFTWARE\Policies\Jeremy Saunders\ComputerRestartService";

            public static string regPreferencePath = @"SOFTWARE\Jeremy Saunders\ComputerRestartService";
        }

        // A method that reads the appSettings from the registry and updates the global variables.

        public static string GetConfigValueFromRegistry(string valueName)
        {
            StringBuilder stringBuilder = new StringBuilder();
            string output = string.Empty;
            bool valuefound = false;
            try
            {
                RegistryKey subKey1 = Registry.LocalMachine.OpenSubKey(Globals.regPolicyPath, false);
                if (subKey1 != null)
                {
                    if (subKey1.GetValue(valueName) != null)
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: Checking for the '" + subKey1.ToString() + @"\" + valueName + "' under the policies key with a value of " + output);
                        output = (subKey1.GetValue(valueName).ToString());
                        if (!string.IsNullOrEmpty(output))
                        {
                            stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: " + valueName + " found under the policies key with a value of " + output);
                            valuefound = true;
                        }
                    }
                    else
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: " + valueName + " not found under the policies key: " + ex.Message.ToString());
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
                    RegistryKey subKey2 = Registry.LocalMachine.OpenSubKey(Globals.regPreferencePath, false);
                    if (subKey2 != null)
                    {
                        if (subKey2.GetValue(valueName) != null)
                        {
                            stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: Checking for the '" + subKey2.ToString() + @"\" + valueName + "' under the policies key with a value of " + output);
                            output = (string)subKey2.GetValue(valueName);
                            if (!string.IsNullOrEmpty(output))
                            {
                                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: " + valueName + " found under the preferences key with a value of " + output);
                                valuefound = true;
                            }
                        }
                        subKey2.Close();
                    }
                    else
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigValueFromRegistry: " + valueName + " not found under the preferences key: " + ex.Message.ToString());
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

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("InactivityTimerInMinutes")))
            {
                int.TryParse(GetConfigValueFromRegistry("InactivityTimerInMinutes"), out int InactivityTimerInMinutes);
                Globals.InactivityTimerInMinutes = InactivityTimerInMinutes;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - InactivityTimerInMinutes is set to: " + Globals.InactivityTimerInMinutes.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - InactivityTimerInMinutes is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
            }

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("DisconnectTimerInMinutes")))
            {
                int.TryParse(GetConfigValueFromRegistry("DisconnectTimerInMinutes"), out int DisconnectTimerInMinutes);
                Globals.DisconnectTimerInMinutes = DisconnectTimerInMinutes;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - DisconnectTimerInMinutes is set to: " + Globals.DisconnectTimerInMinutes.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - InactivityTimerInMinutes is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
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

            if (!string.IsNullOrEmpty(GetConfigValueFromRegistry("ShutdownAfterLogoff")))
            {
                bool.TryParse(GetConfigValueFromRegistry("ShutdownAfterLogoff"), out bool ShutdownAfterLogoff);
                Globals.ShutdownAfterLogoff = ShutdownAfterLogoff;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - ShutdownAfterLogoff is set to: " + Globals.ShutdownAfterLogoff.ToString());
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettingsFromRegistry: - ShutdownAfterLogoff is not set in the registry. Will default to what is set in the " + Assembly.GetEntryAssembly().GetName().Name + ".exe.config");
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
                if (s == "InactivityTimerInMinutes")
                {
                    int.TryParse(appSettings.Get(s), out int InactivityTimerInMinutes);
                    Globals.InactivityTimerInMinutes = InactivityTimerInMinutes;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - InactivityTimerInMinutes is set to: " + Globals.InactivityTimerInMinutes.ToString());
                }
                if (s == "DisconnectTimerInMinutes")
                {
                    int.TryParse(appSettings.Get(s), out int DisconnectTimerInMinutes);
                    Globals.DisconnectTimerInMinutes = DisconnectTimerInMinutes;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - DisconnectTimerInMinutes is set to: " + Globals.InactivityTimerInMinutes.ToString());
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
                if (s == "ShutdownAfterLogoff")
                {
                    bool.TryParse(appSettings.Get(s), out bool ShutdownAfterLogoff);
                    Globals.ShutdownAfterLogoff = ShutdownAfterLogoff;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetConfigurationSettings: - ShutdownAfterLogoff is set to: " + Globals.ShutdownAfterLogoff.ToString());
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
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed: The " + e.FullPath + " file has changed.");
            try
            {
                fileSystemWatcher.Changed -= FileSystemWatcher_Changed;
                fileSystemWatcher.EnableRaisingEvents = false;
                // Read and update the global variables from the appSettings section of the App.config
                if (GetIdleFile(e.FullPath))
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed: Updating the configuration settings.");
                    if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
                    if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
                    stringBuilder.Clear();
                    GetConfigurationSettings();
                }
                // Read and update the global variables from the registry if UseSettingsFromRegistry is set to true
                if (Globals.UseSettingsFromRegistry)
                {
                    GetConfigurationSettingsFromRegistry();
                }
                if (Globals.RestartAfterLogoff && Globals.ShutdownAfterLogoff)
                {
                    Globals.ShutdownAfterLogoff = false;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed: As RestartAfterLogoff is set to true, the ShutdownAfterLogoff will be set to false");
                }

                // If the actual uptime is less than the MaxUptimeInDays variable, set the timer interval to the difference plus the repeat timer, for some buffer, and ensure the timer is enabled.
                // Else
                // If the RestartAfterLogoff variable is false, set the timer interval to the RepeatTimerInMilliseond and ensure the timer is enabled.
                // else disable the timer because the logoff event will be used to restart this computer.
                // This value can be too large for an Int32, so we use long.
                if (Convert.ToInt64(GetUptime().TotalMilliseconds) < (Globals.MaxUptimeInDays * 86400000))
                {
                    long interval = ((Globals.MaxUptimeInDays) * 86400000) - Convert.ToInt64(GetUptime().TotalMilliseconds + Globals.RepeatTimerInMilliseonds);
                    // The Timer.Interval Property is the time, in milliseconds, between Elapsed events. The value must be greater than zero, and less than or equal to Int32.MaxValue.
                    // As the default is 100 milliseconds, this is what we set it too for the lowest value to avoid issues plus the Globals.RepeatTimerInMilliseond.
                    if (interval <= 0)
                    {
                        interval = 100 + Globals.RepeatTimerInMilliseonds;
                    }
                    if (interval > Int32.MaxValue)
                    {
                        interval = Int32.MaxValue;
                    }
                    _checkUptimeTimer.Interval = interval;
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed: - The timer interval has now been updated to check every " + interval.ToString() + " milliseconds.");
                    _checkUptimeTimer.Enabled = true;
                }
                else
                {
                    if (!Globals.RestartAfterLogoff && !Globals.ShutdownAfterLogoff)
                    {
                        // Update the Timer interval. If the interval is set after the Timer has started, the counter is reset.
                        if (_checkUptimeTimer.Interval != Globals.RepeatTimerInMilliseonds)
                        {
                            _checkUptimeTimer.Interval = Globals.RepeatTimerInMilliseonds;
                            if (Globals.RepeatTimerInMilliseonds < (Globals.MaxUptimeInDays * 86400000))
                            {
                                stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed: - The timer interval has now been updated to check every " + Globals.RepeatTimerInMilliseonds.ToString() + " milliseconds.");
                            }
                            else
                            {
                                stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed: - The timer interval has now been updated to check after " + Globals.RepeatTimerInMilliseonds.ToString() + " milliseconds.");
                            }
                        }
                        _checkUptimeTimer.Enabled = true;
                    }
                    else
                    {
                        // Desctivate/Stop the Timer
                        _checkUptimeTimer.Enabled = false;
                        if (Globals.RestartAfterLogoff)
                        {
                            stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed: - The timer has now been disabled because the RestartAfterLogoff variable is set to True. Therefore the logoff event will be used to restart this computer.");
                        }
                        else
                        {
                            stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed: - The timer has now been disabled because the ShutdownAfterLogoff variable is set to True. Therefore the logoff event will be used to shutdown this computer.");
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": FileSystemWatcher_Changed: " + exception.Message);
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

        // Logoff the session using unmanaged code
        public static void LogoffSession(string username, string sessionid)
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": LogoffSession: Attempting to logoff " + username + " session ID" + sessionid + ".");
            string stdOutput = string.Empty;
            string stdError = string.Empty;
            var proc = new ProcessStartInfo();
            proc.FileName = "logoff.exe";
            proc.Arguments = sessionid + " /V";
            proc.CreateNoWindow = true;
            proc.UseShellExecute = false;
            proc.RedirectStandardOutput = true;
            proc.RedirectStandardError = true;
            var pi = new Process();
            pi.StartInfo = proc;
            try
            {
                pi.Start();
                stdOutput = pi.StandardOutput.ReadToEnd();
                stdError = pi.StandardError.ReadToEnd();
                pi.WaitForExit();
            }
            catch (Exception e)
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": LogoffSession: " + e.Message);
            }
            if (pi.ExitCode == 0)
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": LogoffSession: " + stdOutput.ToString());
            }
            else
            {
                if (!string.IsNullOrEmpty(stdError))
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": LogoffSession: " + stdError.ToString());
                }
                if (stdOutput.Length != 0)
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": LogoffSession: " + stdOutput.ToString());
                }
            }
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }
        
        // Restart the computer using unmanaged code
        public static void RestartMachineUnmanagedCode(bool restart, bool force, int timeout)
        {
            StringBuilder stringBuilder = new StringBuilder();
            if (restart)
            {
                _eventLog.WriteEntry("The Computer Restart Service has initiated a restart.", EventLogEntryType.Information, 100);
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": RestartMachineUnmanagedCode: The Computer Restart Service has initiated a restart.");
            }
            else
            {
                _eventLog.WriteEntry("The Computer Restart Service has initiated a shutdown.", EventLogEntryType.Information, 100);
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": RestartMachineUnmanagedCode: The Computer Restart Service has initiated a shutdown.");
            }
            string filename = "shutdown.exe";
            string arguments = string.Empty;
            if (restart)
            {
                arguments = "/r /t " + timeout;
            }
            else
            {
                arguments = "/s /t " + timeout;
            }
            if (force)
            {
                arguments = arguments + " /f";
            }
            var proc = new ProcessStartInfo(filename, arguments);
            proc.CreateNoWindow = true;
            proc.UseShellExecute = false;
            Process.Start(proc);
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }

        // Restart the computer using managed code
        public static void RestartMachineManagedCode(bool restart, bool force, int timeout)
        {
            StringBuilder stringBuilder = new StringBuilder();
            if (restart)
            {
                _eventLog.WriteEntry("The Computer Restart Service has initiated a restart.", EventLogEntryType.Information, 100);
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": RestartMachineManagedCode: The Computer Restart Service has initiated a restart.");
            }
            else
            {
                _eventLog.WriteEntry("The Computer Restart Service has initiated a shutdown.", EventLogEntryType.Information, 100);
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": RestartMachineManagedCode: The Computer Restart Service has initiated a shutdown.");
            }
            ManagementClass managementClass = new ManagementClass("Win32_OperatingSystem");
            managementClass.Get();
            bool EnablePrivileges = managementClass.Scope.Options.EnablePrivileges;
            managementClass.Scope.Options.EnablePrivileges = true;
            ManagementBaseObject methodParameters = managementClass.GetMethodParameters("Win32ShutdownTracker");
            if (!force)
            {
                if (restart)
                {
                    methodParameters["Flags"] = 2; //Graceful Reboot
                }
                else
                {
                    methodParameters["Flags"] = 1; //Graceful Shutdown
                }
            }
            else
            {
                if (restart)
                {
                    methodParameters["Flags"] = 6; //Forced Reboot
                }
                else
                {
                    methodParameters["Flags"] = 5; //Forced Shutdown
                }
            }
            methodParameters["Timeout"] = timeout; //Timeout
            foreach (ManagementObject instance in managementClass.GetInstances())
            {
                var outParams = instance.InvokeMethod("Win32ShutdownTracker", methodParameters, (InvokeMethodOptions)null);
                int.TryParse(outParams["ReturnValue"].ToString(), out int returnCode);
                if (returnCode != 0)
                {
                    var ex = new Win32Exception(returnCode);
                    if (restart)
                    {
                        _eventLog.WriteEntry("The Computer Restart Service failed to restart the computer: " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim(), EventLogEntryType.Error, 100);
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": RestartMachineManagedCode: The Computer Restart Service failed to restart the computer: " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim());
                    }
                    else
                    {
                        _eventLog.WriteEntry("The Computer Restart Service failed to shutdown the computer: " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim(), EventLogEntryType.Error, 100);
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": RestartMachineManagedCode: The Computer Restart Service failed to shutdown the computer: " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim());
                    }
                }
            }
            managementClass.Scope.Options.EnablePrivileges = EnablePrivileges;
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }

        // The SeesionIDMapping dictionary, which keeps a mapping of the SessionID to Username.
        public static Dictionary<string, string> SeesionIDMapping = new Dictionary<string, string>();

        // The LoggedOnUser class is used by the output of the EnumerateSessionInfo function. It's purposely been made to match the output of the "query user" command.
        public class LoggedOnUsers
        {
            public string Username { get; set; }
            public string SessionName { get; set; }
            public string SessionID { get; set; }
            public string State { get; set; }
            public string IdleTime { get; set; }
            public DateTime LogonTime { get; set; }
            public bool StaleSession { get; set; }
        }

        // Constants for session information classes
        const int WTS_CURRENT_SERVER_HANDLE = -1;

        // WTS_INFO_CLASS enum for session information
        public enum WTS_INFO_CLASS
        {
            WTSSessionId,
            WTSUserName,
            WTSWinStationName,
            WTSClientName,
            WTSClientAddress,
            WTSClientDisplay,
            WTSClientProtocolType,
            WTSIdleTime,
            WTSLogonTime,
            WTSConnectState,
        }

        // Structure for session information
        [StructLayout(LayoutKind.Sequential)]
        public struct WTS_SESSION_INFO
        {
            public int SessionID;
            public string UserName;
            public string SessionName;
            public string ClientName;
            public string ClientAddress;
            public string ClientDisplay;
            public int ClientProtocolType;
            public long IdleTime;
            public long LogonTime;
            public int ConnectState;
        }

        [DllImport("wtsapi32.dll")]
        static extern bool WTSEnumerateSessions(IntPtr hServer, int Reserved, int Version, ref IntPtr ppSessionInfo, ref int pCount);

        [DllImport("wtsapi32.dll")]
        static extern void WTSFreeMemory(IntPtr pMemory);

        [DllImport("wtsapi32.dll", SetLastError = true)]
        static extern bool WTSQuerySessionInformation(IntPtr hServer, int sessionId, WTS_INFO_CLASS wtsInfoClass, out IntPtr ppBuffer, out int pBytesReturned);

        public static List<WTS_SESSION_INFO> EnumerateSessionInfo()
        {
            List<WTS_SESSION_INFO> sessionInfoList = new List<WTS_SESSION_INFO>();
            IntPtr pSessionInfo = IntPtr.Zero;
            int sessionCount = 0;

            if (WTSEnumerateSessions(IntPtr.Zero, 0, 1, ref pSessionInfo, ref sessionCount))
            {
                IntPtr current = pSessionInfo;
                for (int i = 0; i < sessionCount; i++)
                {
                    int sessionId = Marshal.ReadInt32(current);
                    current += sizeof(int);

                    // Get session information for all required fields in a single call
                    IntPtr pInfo = IntPtr.Zero;
                    int bytesReturned;

                    if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WTS_INFO_CLASS.WTSUserName, out pInfo, out bytesReturned))
                    {
                        WTS_SESSION_INFO sessionInfo = new WTS_SESSION_INFO();
                        sessionInfo.SessionID = sessionId;
                        sessionInfo.UserName = Marshal.PtrToStringAnsi(pInfo);
                        WTSFreeMemory(pInfo);

                        if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WTS_INFO_CLASS.WTSWinStationName, out pInfo, out bytesReturned))
                        {
                            sessionInfo.SessionName = Marshal.PtrToStringAnsi(pInfo);
                            WTSFreeMemory(pInfo);
                        }

                        if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WTS_INFO_CLASS.WTSClientName, out pInfo, out bytesReturned))
                        {
                            sessionInfo.ClientName = Marshal.PtrToStringAnsi(pInfo);
                            WTSFreeMemory(pInfo);
                        }

                        if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WTS_INFO_CLASS.WTSClientAddress, out pInfo, out bytesReturned))
                        {
                            sessionInfo.ClientAddress = Marshal.PtrToStringAnsi(pInfo);
                            WTSFreeMemory(pInfo);
                        }

                        if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WTS_INFO_CLASS.WTSClientDisplay, out pInfo, out bytesReturned))
                        {
                            sessionInfo.ClientDisplay = Marshal.PtrToStringAnsi(pInfo);
                            WTSFreeMemory(pInfo);
                        }

                        if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WTS_INFO_CLASS.WTSClientProtocolType, out pInfo, out bytesReturned))
                        {
                            sessionInfo.ClientProtocolType = Marshal.ReadInt32(pInfo);
                            WTSFreeMemory(pInfo);
                        }

                        if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WTS_INFO_CLASS.WTSIdleTime, out pInfo, out bytesReturned))
                        {
                            sessionInfo.IdleTime = Marshal.ReadInt64(pInfo);
                            WTSFreeMemory(pInfo);
                        }

                        if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WTS_INFO_CLASS.WTSLogonTime, out pInfo, out bytesReturned))
                        {
                            // If LogonTime is to be represented in long format, use...
                            sessionInfo.LogonTime = Marshal.ReadInt64(pInfo);
                            // If LogonTime is to be represented in DateTime format use...
                            //long logonTimeTicks = Marshal.ReadInt64(pInfo);
                            //sessionInfo.LogonTime = DateTime.FromFileTimeUtc(logonTimeTicks);
                            WTSFreeMemory(pInfo);
                        }

                        if (WTSQuerySessionInformation(IntPtr.Zero, sessionId, WTS_INFO_CLASS.WTSConnectState, out pInfo, out bytesReturned))
                        {
                            sessionInfo.ConnectState = Marshal.ReadInt32(pInfo);
                            WTSFreeMemory(pInfo);
                        }

                        sessionInfoList.Add(sessionInfo);
                    }
                }

                WTSFreeMemory(pSessionInfo);
            }
            else
            {
                Console.WriteLine($"WTSEnumerateSessions failed with error code: {Marshal.GetLastWin32Error()}");
            }

            return sessionInfoList;
        }

        public List<LoggedOnUsers> GetNumberLoggedOnUserSessions()
        {
            StringBuilder stringBuilder = new StringBuilder();

            List<LoggedOnUsers> loggedonusers = new List<LoggedOnUsers> { };
            
            List<WTS_SESSION_INFO> sessions = EnumerateSessionInfo();

            if (sessions.Count > 0)
            {
                foreach (var session in sessions)
                {
                    SeesionIDMapping.Clear();

                    int TotalDisconnectedTimeInMinutes = 0;
                    bool IsStaleSession = false;

                    DateTime parsedLogonTime;
                    if (!DateTime.TryParse(session.LogonTime.ToString().Trim(), out parsedLogonTime))
                    {
                        parsedLogonTime = DateTime.MinValue;
                    }
                    if (session.ConnectState.ToString().Trim().IndexOf("Disc", StringComparison.CurrentCultureIgnoreCase) == 0)
                    {
                        if (session.IdleTime.ToString().Trim().IndexOf("none", StringComparison.CurrentCultureIgnoreCase) < 0 && session.IdleTime.ToString().Trim().IndexOf(".", StringComparison.CurrentCultureIgnoreCase) < 0)
                        {
                            int DaysDisconnected = 0;
                            int HoursDisconnected = 0;
                            int MinutesDisconnected = 0;
                            if (session.IdleTime.ToString().Trim().IndexOf("+") > 0)
                            {
                                int.TryParse(session.IdleTime.ToString().Trim().Split('+')[0], out DaysDisconnected);
                                int.TryParse(session.IdleTime.ToString().Trim().Split('+')[1].Split(':')[0], out HoursDisconnected);
                                int.TryParse(session.IdleTime.ToString().Trim().Split('+')[1].Split(':')[1], out MinutesDisconnected);
                            }
                            else if (session.IdleTime.ToString().Trim().IndexOf(":") > 0)
                            {
                                int.TryParse(session.IdleTime.ToString().Trim().Split(':')[0], out HoursDisconnected);
                                int.TryParse(session.IdleTime.ToString().Trim().Split(':')[1], out MinutesDisconnected);
                            }
                            else
                            {
                                int.TryParse(session.IdleTime.ToString().Trim(), out MinutesDisconnected);
                            }
                            TotalDisconnectedTimeInMinutes = (DaysDisconnected * 1440) + (HoursDisconnected * 60) + MinutesDisconnected;
                        }
                    }
                    if (TotalDisconnectedTimeInMinutes > Globals.DisconnectTimerInMinutes)
                    {
                        IsStaleSession = true;
                        // logoff the session
                    }
                    loggedonusers.Add(new LoggedOnUsers
                    {
                        Username = session.UserName.ToString().Trim(),
                        SessionName = session.SessionName.ToString().Trim(),
                        SessionID = session.SessionID.ToString().Trim(),
                        State = session.ConnectState.ToString().Trim(),
                        IdleTime = session.IdleTime.ToString().Trim(),
                        LogonTime = parsedLogonTime,
                        StaleSession = IsStaleSession
                    });
                    SeesionIDMapping.Add(session.SessionID.ToString().Trim(), session.UserName.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions: - Username: " + session.UserName.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - SessionName: " + session.SessionName.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - SessionID: " + session.SessionID.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - State: " + session.ConnectState.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - IdleTime: " + session.IdleTime.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - LogonTime: " + parsedLogonTime);
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - StaleSession: " + IsStaleSession.ToString());
                }
                if (sessions.Count == 1)
                {
                    stringBuilder.Insert(0, DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions: Found " + sessions.Count + " logged on user session:" + Environment.NewLine);
                }
                else
                {
                    stringBuilder.Insert(0, DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions: Found " + sessions.Count + " logged on user sessions:" + Environment.NewLine);
                }
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions: No logged on user sessions were found");
            }
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
            return loggedonusers;
        }

        // Run the "query user" command and process the output.
        public List<LoggedOnUsers> GetNumberLoggedOnUserSessions2()
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
                    int TotalDisconnectedTimeInMinutes = 0;
                    bool IsStaleSession = false;
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
                    if (match.Groups["state"].Value.ToString().Trim().IndexOf("Disc", StringComparison.CurrentCultureIgnoreCase) == 0)
                    {
                        if (match.Groups["idletime"].Value.ToString().Trim().IndexOf("none", StringComparison.CurrentCultureIgnoreCase) < 0 && match.Groups["idletime"].Value.ToString().Trim().IndexOf(".", StringComparison.CurrentCultureIgnoreCase) < 0)
                        {
                            int DaysDisconnected = 0;
                            int HoursDisconnected = 0;
                            int MinutesDisconnected = 0;
                            if (match.Groups["idletime"].Value.ToString().Trim().IndexOf("+") > 0)
                            {
                                int.TryParse(match.Groups["idletime"].Value.ToString().Trim().Split('+')[0], out DaysDisconnected);
                                int.TryParse(match.Groups["idletime"].Value.ToString().Trim().Split('+')[1].Split(':')[0], out HoursDisconnected);
                                int.TryParse(match.Groups["idletime"].Value.ToString().Trim().Split('+')[1].Split(':')[1], out MinutesDisconnected);
                            }
                            else if (match.Groups["idletime"].Value.ToString().Trim().IndexOf(":") > 0)
                            {
                                int.TryParse(match.Groups["idletime"].Value.ToString().Trim().Split(':')[0], out HoursDisconnected);
                                int.TryParse(match.Groups["idletime"].Value.ToString().Trim().Split(':')[1], out MinutesDisconnected);
                            }
                            else
                            {
                                int.TryParse(match.Groups["idletime"].Value.ToString().Trim(), out MinutesDisconnected);
                            }
                            TotalDisconnectedTimeInMinutes = (DaysDisconnected * 1440) + (HoursDisconnected * 60) + MinutesDisconnected;
                        }
                    }
                    if (TotalDisconnectedTimeInMinutes > Globals.DisconnectTimerInMinutes)
                    {
                        IsStaleSession = true;
                        // logoff the session
                    }
                    loggedonusers.Add(new LoggedOnUsers
                    {
                        Username = username,
                        SessionName = match.Groups["sessionname"].Value.ToString().Trim(),
                        SessionID = match.Groups["sessionid"].Value.ToString().Trim(),
                        State = match.Groups["state"].Value.ToString().Trim(),
                        IdleTime = match.Groups["idletime"].Value.ToString().Trim(),
                        LogonTime = parsedLogonTime,
                        StaleSession = IsStaleSession
                    });
                    SeesionIDMapping.Add(match.Groups["sessionid"].Value.ToString().Trim(), username);
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions: - Username: " + username);
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - SessionName: " + match.Groups["sessionname"].Value.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - SessionID: " + match.Groups["sessionid"].Value.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - State: " + match.Groups["state"].Value.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - IdleTime: " + match.Groups["idletime"].Value.ToString().Trim());
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - LogonTime: " + parsedLogonTime);
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": GetNumberLoggedOnUserSessions:   - StaleSession: " + IsStaleSession.ToString());
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

        // Setup the Timer objects.
        private static Timer _checkUptimeTimer = new Timer();
        private static Timer _checkInactivityTimer = new Timer();

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
            if (Globals.RestartAfterLogoff && Globals.ShutdownAfterLogoff) {
                Globals.ShutdownAfterLogoff = false;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStart: As RestartAfterLogoff is set to true, the ShutdownAfterLogoff will be set to false");
            }

            // Create a list of logged on users at service start
            GetNumberLoggedOnUserSessions();

            // Setup and call the file change event watcher
            WatchForFileChangeEvent();

            // Setup and call the security event watcher
            WatchForSecurityEvent();

            // Call and set the computer uptime
            int actualUptimeInDays = Convert.ToInt32(GetUptime().TotalDays);
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStart: Computer uptime in days: " + actualUptimeInDays.ToString());
            // This value when converted to milliseconds can be too large for an Int32, so we use long.
            long actualUptimeInMilliseconds = Convert.ToInt64(GetUptime().TotalMilliseconds);
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStart: Computer uptime in milliseconds: " + actualUptimeInMilliseconds.ToString());

            // Setup and call the Uptime Timer
            _checkUptimeTimer.Elapsed += new ElapsedEventHandler(CheckUptimeTimerElapsed);

            _checkUptimeTimer.AutoReset = true;
            // If the actual uptime is less than or equal to the MaxUptimeInDays variable, set the timer interval to the difference in milliseonds.
            // We do this to avoid unnecessarily checking the MaxUptime until it has actually been reached. If it has been reached,
            // we set it to an interval as specified in the App.Config file.
            // 1 day = 86400000 milliseonds
            long upTimeinterval = Globals.RepeatTimerInMilliseonds;
            if (actualUptimeInMilliseconds <= (Globals.MaxUptimeInDays * 86400000))
            {
                upTimeinterval = (Globals.MaxUptimeInDays * 86400000) - actualUptimeInMilliseconds + Globals.RepeatTimerInMilliseonds;
            }
            _checkUptimeTimer.Interval = upTimeinterval;
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStart: Setting the UpTime timer interval to " + upTimeinterval.ToString() + " milliseconds.");

            // Activate/Start the Uptime Timer
            _checkUptimeTimer.Enabled = true;

            // Setup and call the Inactivity Timer
            _checkInactivityTimer.Elapsed += new ElapsedEventHandler(CheckInactivityTimerElapsed);

            _checkInactivityTimer.AutoReset = true;
            // If the actual uptime is less than or equal to the InactivityTimerInMinutes variable, set the timer interval to the difference in milliseonds.
            // We do this to avoid unnecessarily checking the Inactivity until it has actually been reached. If it has been reached,
            // we set it to an interval as specified in the App.Config file.
            // 1 day = 86400000 milliseonds
            long InactivityInterval = Globals.InactivityTimerInMinutes;
            if (actualUptimeInMilliseconds <= (Globals.InactivityTimerInMinutes * 86400000))
            {
                InactivityInterval = (Globals.InactivityTimerInMinutes * 86400000) - actualUptimeInMilliseconds + Globals.RepeatTimerInMilliseonds;
            }

            // The Timer.Interval Property is the time, in milliseconds, between Elapsed events. The value must be greater than zero, and less than or equal to Int32.MaxValue.
            // As the default is 100 milliseconds, this is what we set it too for the lowest value to avoid issues.
            // Remember that if the Globals.InactivityTimerInMinutes is 0, the _checkInactivityTimer is disabled.
            if (InactivityInterval <= 0)
            {
                InactivityInterval = 100;
            }
            if (InactivityInterval > Int32.MaxValue)
            {
                InactivityInterval = Int32.MaxValue;
            }
            _checkInactivityTimer.Interval = InactivityInterval;
            stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStart: Setting the Inactivity timer interval to " + InactivityInterval.ToString() + " milliseconds.");

            // Activate/Start the Inactivity Timer
            if (Globals.InactivityTimerInMinutes > 0)
            {
                _checkInactivityTimer.Enabled = true;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStart: Enabling the Inactivity timer.");
            }
            else
            {
                _checkInactivityTimer.Enabled = false;
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnStart: Disabling the Inactivity timer because the InactivityTimerInMinutes setting is set to 0.");
            }

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
                    if (Globals.RestartAfterLogoff || Globals.ShutdownAfterLogoff)
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
                                if (Globals.RestartAfterLogoff)
                                {
                                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: Okay to restart.");
                                    RestartMachineManagedCode(true, Globals.ForceRestart, Globals.DelayBeforeRestartingInSeconds);
                                }
                                else
                                {
                                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: Okay to shutdown.");
                                    RestartMachineManagedCode(false, Globals.ForceRestart, Globals.DelayBeforeRestartingInSeconds);
                                }
                            }
                            else
                            {
                                if (Globals.RestartAfterLogoff)
                                {
                                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: It cannot be restarted at this point.");
                                }
                                else
                                {
                                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": OnSessionChange: It cannot be shutdown at this point.");
                                }
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

        public void CheckInactivityTimerElapsed(object sender, ElapsedEventArgs e)
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
                    if (Globals.RestartAfterLogoff)
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed: - Restarting...");
                        RestartMachineManagedCode(true, Globals.ForceRestart, Globals.DelayBeforeRestartingInSeconds);
                    }
                    else
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed: - Shutting down...");
                        RestartMachineManagedCode(false, Globals.ForceRestart, Globals.DelayBeforeRestartingInSeconds);
                    }
                }
                else
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - " + CurrentLoggedOnUserCount + " users are still logged in.");
                    if (Globals.RestartAfterLogoff)
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - It cannot be restarted at this point.");
                    }
                    else
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - It cannot be shutdown at this point.");
                    }
                }
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed: " + computerName + " has not yet reached its maximum uptime of " + Globals.MaxUptimeInDays.ToString() + " days.");
            }
            if (!Globals.RestartAfterLogoff && !Globals.ShutdownAfterLogoff)
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
                if (Globals.RestartAfterLogoff)
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - The timer has now been disabled because the RestartAfterLogoff variable is set to True. Therefore the logoff event will restart this computer.");
                }
                else
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - The timer has now been disabled because the ShutdownAfterLogoff variable is set to True. Therefore the logoff event will shutdown this computer.");
                }
            }
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
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
                    if (Globals.RestartAfterLogoff)
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed: - Restarting...");
                        RestartMachineManagedCode(true, Globals.ForceRestart, Globals.DelayBeforeRestartingInSeconds);
                    }
                    else
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed: - Shutting down...");
                        RestartMachineManagedCode(false, Globals.ForceRestart, Globals.DelayBeforeRestartingInSeconds);
                    }
                }
                else
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - " + CurrentLoggedOnUserCount + " users are still logged in.");
                    if (Globals.RestartAfterLogoff)
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - It cannot be restarted at this point.");
                    }
                    else
                    {
                        stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - It cannot be shutdown at this point.");
                    }
                }
            }
            else
            {
                stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed: " + computerName + " has not yet reached its maximum uptime of " + Globals.MaxUptimeInDays.ToString() + " days.");
            }
            if (!Globals.RestartAfterLogoff && !Globals.ShutdownAfterLogoff)
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
                if (Globals.RestartAfterLogoff)
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - The timer has now been disabled because the RestartAfterLogoff variable is set to True. Therefore the logoff event will restart this computer.");
                }
                else
                {
                    stringBuilder.AppendLine(DateTime.Now.ToString() + ": checkUptimeTimerElapsed : - The timer has now been disabled because the ShutdownAfterLogoff variable is set to True. Therefore the logoff event will shutdown this computer.");
                }
            }
            if (Globals.DebugToEventLog) { _eventLog.WriteEntry(stringBuilder.ToString(), EventLogEntryType.Information, 100); }
            if (Globals.DebugToFile) { File.AppendAllText(Globals.logfile, stringBuilder.ToString()); }
        }

    }
}
