using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace PowerMode
{
    /// <summary>
    /// This program allows for setting the Windows "power mode" or "power slider" value from the command line.
    /// </summary>
    /// <seealso cref="https://docs.microsoft.com/en-us/windows-hardware/customize/desktop/customize-power-slider"/>
    class SetPowerMode
    {
        private static Dictionary<string, Guid> PowerModes = new Dictionary<string, Guid>(StringComparer.InvariantCultureIgnoreCase);

        /// <summary>
        /// Execution starts here.
        /// </summary>
        /// <param name="args">Command line parameters.</param>
        /// <returns>Error status; 0 = success, non-zero = failure.</returns>
        static int Main(string[] args)
        {
            try
            {
                // Read from App.config.
                ReadConfig();

                if (args.Length == 0)
                {
                    // Report the current power mode.
                    uint result = PowerGetEffectiveOverlayScheme(out Guid currentMode);
                    if (result == 0)
                    {
                        //Console.WriteLine("Current Mode GUID: {0}", currentMode.ToString());
                        string Key = PowerModes.FirstOrDefault(x => x.Value == currentMode).Key;
                        if (Key.Length > 0)
                            Console.WriteLine($"{Key}");
                        else
                            Console.WriteLine("*** Power Mode not configured ***");
                    }
                    else
                    {
                        return (int)result;
                    }
                }
                else if (args.Length == 1)
                {
                    // Attempt to set the power mode.
                    string parameter = args[0];
                    Guid powerMode;

                    if (parameter == "/?" || parameter == "-?")
                    {
                        Usage();
                        return 1;
                    }
                    else if (PowerModes.ContainsKey(parameter))
                    {
                        powerMode = PowerModes[parameter];
                    }
                    else
                    {
                        try
                        {
                            powerMode = new Guid(parameter);
                        }
                        catch (Exception)
                        {
                            Console.Error.WriteLine("Failed to parse GUID.\n");
                            Usage();
                            return 1;
                        }
                    }
                    uint result = PowerSetActiveOverlayScheme(powerMode);

                    if (result == 0)
                    {
                        if(PowerModes.ContainsKey(parameter))
                            Console.WriteLine("Set power mode to {0}.", parameter);
                        else
                            Console.WriteLine("Set power mode to {0}.", powerMode);
                    }
                    else
                    {
                        Console.Error.WriteLine("Failed to set power mode.\n");
                        Usage();
                    }

                    return (int)result;
                }
                else
                {
                    Usage();
                    return 1;
                }
            }
            catch (Exception exception)
            {
                // Print error information to the console.
                Console.Error.WriteLine("{0}: {1}\n{2}", exception.GetType(), exception.Message, exception.StackTrace);
                Console.WriteLine();
                Usage();
                return 1;
            }

            return 0;
        }

        /// <summary>
        /// Print a usage message to the console.
        /// </summary>
        private static void Usage()
        {
            Console.WriteLine(
                    "PowerMode (GPLv3); used to set the active power mode on Windows 10, version 1709 or later\n" +
                    "https://github.com/AaronKelley/PowerMode\n"
                );
            Console.WriteLine("  PowerMode {0,-20} - Report the current power mode", "");
            foreach(KeyValuePair<string, Guid> entry in PowerModes)
            {
                string humanReadable = AddSpacesToSentence(entry.Key);
                Console.WriteLine("  PowerMode {0,-20} - Set the system to \"{1}\" mode", entry.Key, humanReadable);
            }
            
            Console.WriteLine("  PowerMode {0,-20} - Set the system to the mode identified by the GUID", "<GUID>");
        }

        /// <summary>
        /// Add space to a sentence before capital letters
        /// </summary>
        /// <param name="text">AddSpacesBeforeCapitalLETTERS</param>
        /// <returns>Add Spaces Before Capital Letters</returns>
        private static string AddSpacesToSentence(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            StringBuilder newText = new StringBuilder(text.Length * 2);
            newText.Append(text[0]);
            for (int i = 1; i < text.Length; i++)
            {
                if (char.IsUpper(text[i]) && text[i - 1] != ' ')
                    newText.Append(' ');
                newText.Append(text[i]);
            }
            return newText.ToString();
        }

        /// <summary>
        /// Read `App.config` and build dictionary
        /// </summary>
        private static void ReadConfig()
        {
            foreach (string key in ConfigurationManager.AppSettings)
                PowerModes.Add(key, new Guid(ConfigurationManager.AppSettings[key]));
        }


        /// <summary>
        /// Retrieves the active overlay power scheme and returns a GUID that identifies the scheme.
        /// </summary>
        /// <param name="EffectiveOverlayPolicyGuid">A pointer to a GUID structure.</param>
        /// <returns>Returns zero if the call was successful, and a nonzero value if the call failed.</returns>
        [DllImportAttribute("powrprof.dll", EntryPoint = "PowerGetEffectiveOverlayScheme")]
        private static extern uint PowerGetEffectiveOverlayScheme(out Guid EffectiveOverlayPolicyGuid);

        /// <summary>
        /// Sets the active power overlay power scheme.
        /// </summary>
        /// <param name="OverlaySchemeGuid">The identifier of the overlay power scheme.</param>
        /// <returns>Returns zero if the call was successful, and a nonzero value if the call failed.</returns>
        [DllImportAttribute("powrprof.dll", EntryPoint = "PowerSetActiveOverlayScheme")]
        private static extern uint PowerSetActiveOverlayScheme(Guid OverlaySchemeGuid);
    }
}
