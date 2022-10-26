using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Automation;

//This is all rather messy and was put together more as a test, but it works.
//I'm not sure if I'll tidy it up much as it's not going to need updating much.
//Though it would be good to add in more safety checks and error handling.

#nullable enable
namespace PowerAutomateRunner
{
    internal class Program
    {
        private static readonly Dictionary<string, string?> parsedArgs = new();
        
        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll", EntryPoint = "mouse_event")]
        private static extern void MouseEvent(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [Flags]
        public enum MouseEventFlags
        {
            RightDown = 0x00000008,
            RightUp = 0x00000010
        }

        static void Main(string[] args)
        {
            ProcessArguments(args);
            RunPowerAutomate(out bool appWasAlreadyRunning);
            UIAutomate(parsedArgs["--flow-name"]!, !appWasAlreadyRunning);

#if DEBUG && false
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
#endif
        }

        private static void Exit(string? message, int code = 0)
        {
            if (message != null)
            {
                Console.WriteLine(message);

                if (parsedArgs.ContainsKey("-pause-on-error"))
                {
                    Console.WriteLine("Press enter to exit...");
                    Console.ReadLine();
                }
            }

            Environment.Exit(code);
        }

        private static void ProcessArguments(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];
                string? value = null;
                if (arg.StartsWith("--"))
                {
                    if (i + 1 < args.Length)
                    {
                        value = args[i + 1];
                        i++;
                    }
                }
                parsedArgs.Add(arg, value);
            }

            //Check that the required arguments exist.
            if (!parsedArgs.ContainsKey("--flow-name") || parsedArgs["--flow-name"] == null)
                Exit("Missing required argument: --flow-name", 1);
        }

        private static int? GetIntArg(string key)
        {
            if (parsedArgs.TryGetValue(key, out string? value))
            {
                if (int.TryParse(value, out int result))
                    return result;
            }
            return null;
        }

        private static void RunPowerAutomate(out bool appWasAlreadyRunning)
        {
            //Check if Power Automate is running.
            appWasAlreadyRunning = Process.GetProcessesByName("PAD.Console.Host").Length >= 1;
            if (appWasAlreadyRunning) return;

            //Get the start command using powershell.
            //https://stackoverflow.com/questions/32074404/launching-windows-10-store-apps
            //https://stackoverflow.com/questions/33155081/list-and-launch-metro-apps-in-windows-10-with-c-sharp
            PowerShell ps = PowerShell.Create();
            ps.AddScript("Get-AppxPackage | % { $_.PackageFamilyName + \"!PAD.Console\" }");
            ps.AddCommand("Out-String");
            string[] result = ps.Invoke()
                .SelectMany(r => r.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries))
                .ToArray();
            ps.Dispose();

            string? aumid = result.FirstOrDefault(r => r.StartsWith("Microsoft.PowerAutomateDesktop"));
            if (aumid == null)
                Exit("Power Automate not found.", 2);

            Console.WriteLine("Starting Power Automate...");
            Process.Start($"shell:appsfolder\\{aumid}");

            //Wait for the app to start.
            for (int i = 0; i < (GetIntArg("--retry-attempts") ?? 10); i++)
            {
                if (AutomationElement.RootElement.FindElement(e => e.Name == "Power Automate") != null)
                    break;
                Thread.Sleep(GetIntArg("--retry-interval") ?? 250);
            }
        }

        private static void TryClearNotification()
        {
            //Fix this element never being found.
            AutomationElement? notification = AutomationElement.RootElement.FindElement(e => e.ClassName == "Windows.UI.Core.CoreWindow");
            AutomationElement? pane = notification?.FindElement(e => e.ClassName == "ScrollViewer");
            AutomationElement? toast = pane?.FindElement(e => e.ClassName == "FlexibleToastView");
            AutomationElement? closeButton = toast?.FindElement(e => e.AutomationId == "DismissButton");
            InvokePattern? invokePattern = closeButton?.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern?.Invoke();
        }

        private static void UIAutomate(string flowName, bool forceMinimizeWindow = false)
        {
            //Navigate through the power automate window.
            AutomationElement? powerAutomateWindow = AutomationElement.RootElement.FindElement(e => e.Name == "Power Automate");
            bool wasPowerAutomateWindowMinimized = powerAutomateWindow == null || forceMinimizeWindow;

            if (powerAutomateWindow == null)
            {
                TryClearNotification();

                //If we couldn't find the window, check if it is in the tray.
                AutomationElement? taskbar = AutomationElement.RootElement.FindElement(e => e.ClassName == "Shell_TrayWnd");
                AutomationElement? trayNotifyWnd = taskbar?.FindElement(e => e.ClassName == "TrayNotifyWnd");
                AutomationElement? notificationChevron = trayNotifyWnd?.FindElement(e => e.Name == "Notification Chevron");
                InvokePattern? invokePattern = notificationChevron?.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                invokePattern?.Invoke();

                AutomationElement? trayIcon = null;
                for (int i = 0; i < (GetIntArg("--retry-attempts") ?? 10); i++)
                {
                    AutomationElement? notificationOverflow = AutomationElement.RootElement.FindElement(e => e.ClassName == "NotifyIconOverflowWindow");
                    AutomationElement? overflowNotificationArea = notificationOverflow?.FindElement(e => e.Name == "Overflow Notification Area");
                    trayIcon = overflowNotificationArea?.FindElement(e => e.Name == "Power Automate");

                    if (powerAutomateWindow != null)
                        break;
                    Thread.Sleep(GetIntArg("--retry-interval") ?? 250);
                }
                if (trayIcon == null)
                    Exit("Power Automate tray icon not found.", 3);

                //The invoke method cant seem to trigger a specific event. As we want to right click, we will have to do this manually.
                Rect? bounds = trayIcon?.Current.BoundingRectangle;
                if (bounds.HasValue)
                {
                    //Move the cursor to the center of the icon.
                    SetCursorPos((int)bounds.Value.Left + (int)bounds.Value.Width / 2, (int)bounds.Value.Top + (int)bounds.Value.Height / 2);

                    //Send a right click.
                    //https://stackoverflow.com/questions/2416748/how-do-you-simulate-mouse-click-in-c
                    MouseEvent((int)MouseEventFlags.RightDown | (int)MouseEventFlags.RightUp, 0, 0, 0, 0);
                }

                InvokePattern? invokePattern_1 = null;
                for (int i = 0; i < (GetIntArg("--retry-attempts") ?? 10); i++)
                {
                    AutomationElement? toolBar = AutomationElement.RootElement.FindElement(e => e.Name == "DropDown" && e.LocalizedControlType == "tool bar");
                    AutomationElement? toolBarItem = toolBar?.FindElement(e => e.Name == "Open Power Automate console");
                    invokePattern_1 = toolBarItem?.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;

                    if (invokePattern_1 != null)
                        break;
                    Thread.Sleep(GetIntArg("--retry-interval") ?? 250);
                }
                if (invokePattern_1 == null)
                    Exit("Failed to reveal Power Automate window.", 4);

                invokePattern_1?.Invoke();
            }

            AutomationElement? flows_2 = null;
            for (int i = 0; i < (GetIntArg("--retry-attempts") ?? 10); i++)
            {
                try
                {
                    powerAutomateWindow = AutomationElement.RootElement.FindElement(e => e.Name == "Power Automate");
                    AutomationElement? modalManager = powerAutomateWindow?.FindElement(e => e.ClassName == "ModalManager");
                    AutomationElement? modalManager_1 = modalManager?.FindElement(e => e.ClassName == "ModalManager");
                    AutomationElement? modalManager_2 = modalManager_1?.FindElement(e => e.ClassName == "ModalManager");
                    AutomationElement? processesView = modalManager_2?.FindElement(e => e.ClassName == "ProcessesView");
                    AutomationElement? flows = processesView?.FindElement(e => e.Name == "Flows");
                    AutomationElement? myFlows = flows?.FindElement(e => e.Name == "My flows" && e.ClassName == "TabItem");
                    flows_2 = myFlows?.FindElement(e => e.Name == "Flows" && e.ClassName == "DataGrid");

                    if (flows_2 != null)
                        break;
                }
                catch {}

                Thread.Sleep(GetIntArg("--retry-interval") ?? 250);
            }
            if (flows_2 == null)
                Exit("Could not find the flows UI element.", 5);

            //Find the specified flow.
            AutomationElement? flow = flows_2?.FindElement(e => e.Name == flowName && e.ClassName == "DataGridRow");
            AutomationElement? dataGridCell = flow?.FindElement(e => e.Name == flowName && e.ClassName == "DataGridCell");
            AutomationElement? runButton = dataGridCell?.FindElement(e => e.Name == "Run");
            InvokePattern? invokePattern_2 = runButton?.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern_2?.Invoke();

            Console.WriteLine("Successfully triggered flow.");

            if (!wasPowerAutomateWindowMinimized)
                return;
            AutomationElement? closeButton = powerAutomateWindow?.FindElement(e => e.Name == "Close window");
            InvokePattern? invokePattern_3 = closeButton?.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern_3?.Invoke();
        }
    }
}
