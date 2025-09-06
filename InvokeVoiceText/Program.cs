using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace InvokeVoiceText
{
    internal static class Program
    {
        private const uint WM_SETTEXT = 0x000C;
        private const uint WM_COMMAND = 0x0111;
        private const uint MF_DISABLED = 0x00000002;
        private const uint MF_GRAYED = 0x00000001;
        private const uint MF_BYCOMMAND = 0x00000000;
        private const uint PlayButtonCommandId = 32771;
        private const uint StopButtonCommandId = 32772;
        private const int SynthMenuIndex = 2;

        static IntPtr mainHandle = IntPtr.Zero;

        delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern IntPtr GetMenu(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetSubMenu(IntPtr hMenu, int nPos);

        [DllImport("user32.dll")]
        static extern uint GetMenuState(IntPtr hMenu, uint uId, uint uFlags);

        [STAThread]
        private static void Main(string[] args)
        {
            Console.WriteLine("Use \"--help\" (without quotes) to speak the help guide.");
            Console.WriteLine("Use \"--stop\" (without quotes) to stop audio currently playing from VoiceText.");
            Console.WriteLine("Use \"--speech-txt\" (without quotes) to speak from \"speech.txt\".");
            //Console.WriteLine("Use \"--gui\" (without quotes) to open the furry GUI.");

            string FullSpeak = "This is the help guide of Invoke VoiceText, written by BunnyTub.\r\n" +
                "To change this text, invoke the program with command line arguments of what you want to say.\r\n" +
                "The program will return an exit code of 0 when speaking is finished, or 1 if speaking fails.";

            if (args.Length > 0)
            {
                if (args[0].ToLowerInvariant() != "--help")
                {
                    if (args[0].ToLowerInvariant() != "--stop")
                    {
                        if (args[0].ToLowerInvariant() == "--speech-txt")
                        {
                            try
                            {
                                FullSpeak = File.ReadAllText(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\speech.txt");
                            }
                            catch (Exception ex)
                            {
                                Console.Error.WriteLine(ex.Message);
                                FullSpeak = $"Cannot read from the speech.txt file. {ex.Message}";
                            }
                        }
                        else
                        {
                            FullSpeak = string.Join(" ", args);
                        }
                    }
                    else
                    {
                        FullSpeak = string.Empty;
                    }
                }
            }
            else
            {
                FullSpeak = "Invoke VoiceText has no command line arguments. Please provide command line arguments.";
            }

            Console.WriteLine("--- START OF TEXT ---\r\n" +
                $"{FullSpeak}\r\n" +
                "---  END OF TEXT  ---");

            //bool isSomething(int n)
            //{
            //    return n % 2 == 0;
            //}

            //if (isSomething((int)DateTime.UtcNow.Ticks))
            //{
            //    if (!string.IsNullOrWhiteSpace(FullSpeak))
            //    {
            //        string value = Environment.GetEnvironmentVariable("DisableInvokeVTWatermark");

            //        if (!string.IsNullOrWhiteSpace(value))
            //        {
            //            if (value.ToLowerInvariant() != "yes")
            //            {
            //                // add text watermark haha funny
            //                FullSpeak = $"This copy of Invoke Voice Text is strictly for demonstration purposes only.\r\n" +
            //                    $"{FullSpeak}";
            //            }
            //        }
            //    }
            //}

            if (SubMain(FullSpeak))
            {
                Environment.Exit(0);
            }
            else
            {
                Environment.Exit(1);
            }
            Thread.Sleep(10000);
        }

        [DllImport("shell32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsUserAnAdmin();

        static bool SubMain(string TextToSpeak)
        {
            try
            {
                if (!IsUserAnAdmin())
                {
                    Console.Error.WriteLine("Invoke VoiceText must be run with administrative permissions to interact with VoiceText.");
                    return false;
                }

                Process[] processes = Process.GetProcesses();
                Process process = null;

                foreach (var proc in processes)
                {
                    try
                    {
                        if (proc.ProcessName.ToLowerInvariant().Contains("vt"))
                        {
                            string exePath = proc.MainModule.FileName;
                            var versionInfo = FileVersionInfo.GetVersionInfo(exePath);



                            if (versionInfo.FileDescription.ToLowerInvariant().Contains("vteditor"))
                            {
                                process = proc;
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Cannot find info about {proc.ProcessName}. {ex.Message}");
                    }
                }


                //Process[] processes = Process.GetProcessesByName("VTEditor_ENG");
                if (process == null)
                {
                    Console.Error.WriteLine("No process found. Make sure the VoiceText Editor is open.");
                    return false;
                }

                mainHandle = process.MainWindowHandle;

                IntPtr richEdit = IntPtr.Zero;
                EnumChildWindows(mainHandle, (hwnd, l) =>
                {
                    var sb = new StringBuilder(256);
                    GetClassName(hwnd, sb, sb.Capacity);
                    if (sb.ToString() == "RichEdit20W")
                    {
                        richEdit = hwnd;
                        return false;
                    }
                    return true;
                }, IntPtr.Zero);

                if (!string.IsNullOrEmpty(TextToSpeak))
                {
                    if (richEdit != IntPtr.Zero)
                    {
                        SendText(richEdit, TextToSpeak);
                    }
                    else
                    {
                        Console.Error.WriteLine("No text input discovered. You may be using a version of VoiceText that is unsupported.");
                        return false;
                    }

                    TriggerPlay(mainHandle);

                    bool Reading = false;

                    void timerMethod(object state)
                    {
                        if (CheckReadMenuState())
                        {
                            Reading = true;
                        }
                        else
                        {
                            Reading = false;
                        }
                    }

                    Timer readMenuTimer = new Timer(timerMethod, null, 0, 500);

                    Console.WriteLine("Waiting for VoiceText. If this takes longer than 10 seconds, try again with administrative privileges.");

                    int timePassed = 0;

                    while (!Reading)
                    {
                        Thread.Sleep(1000);
                        timePassed++;
                        if (timePassed >= 10)
                        {
                            Console.Error.WriteLine("VoiceText did not respond.");
                            readMenuTimer.Dispose();
                            return false;
                        }
                    }

                    timePassed = 0;

                    Console.WriteLine("The string is being read by VoiceText.");

                    while (Reading)
                    {
                        Thread.Sleep(100);
                        //timePassed++;
                        //if (timePassed >= 600) // 10 minutes
                        //{
                        //    Console.Error.WriteLine("VoiceText is taking too long to read the string. Returning early.");
                        //    return false;
                        //}
                    }
                    Console.WriteLine("The string is no longer being read by VoiceText.");

                    readMenuTimer.Dispose();
                }
                else
                {
                    TriggerStop(mainHandle);
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return false;
            }
        }

        static void SendText(IntPtr hWnd, string text)
        {
            IntPtr ptr = Marshal.StringToHGlobalUni(text);
            SendMessage(hWnd, WM_SETTEXT, IntPtr.Zero, ptr);
            Marshal.FreeHGlobal(ptr);
        }

        static void TriggerPlay(IntPtr mainWnd)
        {
            IntPtr wParam = (IntPtr)PlayButtonCommandId;
            SendMessage(mainWnd, WM_COMMAND, wParam, IntPtr.Zero);
        }
        
        static void TriggerStop(IntPtr mainWnd)
        {
            IntPtr wParam = (IntPtr)StopButtonCommandId;
            SendMessage(mainWnd, WM_COMMAND, wParam, IntPtr.Zero);
        }

        static readonly uint WM_INITMENUPOPUP = 0x0117;

        // returns true if speaking
        static bool CheckReadMenuState()
        {
            try
            {
                if (mainHandle == IntPtr.Zero)
                    return false;

                IntPtr hMenu = GetMenu(mainHandle);
                if (hMenu == IntPtr.Zero)
                {
                    Console.Error.WriteLine("No menu found. You may be using a version of VoiceText that is unsupported.");
                    return false;
                }

                IntPtr hSynthesizeMenu = GetSubMenu(hMenu, SynthMenuIndex);
                if (hSynthesizeMenu == IntPtr.Zero)
                {
                    Console.Error.WriteLine("No correct menu found. You may be using a version of VoiceText that is unsupported.");
                    return false;
                }

                // Force the menu to update by sending an init command to the window.
                // Otherwise, everything just doesn't work for some reason. I hate the Windows API.
                SendMessage(mainHandle, WM_INITMENUPOPUP, hSynthesizeMenu, new IntPtr(1));

                uint stateVal = GetMenuState(hSynthesizeMenu, StopButtonCommandId, MF_BYCOMMAND);
                bool isDisabled = (stateVal & MF_DISABLED) != 0 || (stateVal & MF_GRAYED) != 0;

                return !isDisabled;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
            }

            return false;
        }
    }
}
