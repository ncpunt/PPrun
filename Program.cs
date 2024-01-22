﻿// System
using System.Media;
using System.Diagnostics;

// Google Cloud TTS API
using Google.Cloud.TextToSpeech.V1;

// Add the following assembly references to the project for PowerPoint automation
//
// C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Office.dll
// C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Microsoft.Office.Interop.Excel.dll
// C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Microsoft.Office.Interop.PowerPoint.dll
//
// Repair Office 365 (Online mode) in case of any COM errors
//
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPrun
{
    internal class Program
    {
        // Start and stop slides 
        static int First = 1;
        static int Last  = 999;

        // Excel Application
        static Excel.Application exa;

        // PowerPoint Application
        static PowerPoint.Application ppa;

        // PowerPoint Presentation
        static PowerPoint.Presentation ppp;

        // Google text to speech client
        static TextToSpeechClient tts;

        // Sound player
        static SoundPlayer player;

        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        /// <param name="args">The arguments.</param>
        static void Main(string[] args)
        {
            try
            {
                // Run in a safe context
                MainSafe(args);
            }
            catch (Exception e)
            {
                // Show error message
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);

            }
            finally
            {
                // Kill orphaned processes
                Kill("EXCEL", "");
                Kill("POWERPNT", "PowerPoint");
                Console.ReadLine();
            }
        }

        /// <summary>
        /// Defines the entry point of the application (Safe Context).
        /// </summary>
        /// <param name="args">The arguments.</param>
        static void MainSafe(string[] args)
        {
            // Check for presentation name
            if (args.Length == 0) throw new Exception("Presentation name missing!");

            // Get the command line arguments
            string fileXLSX = Path.GetFullPath(args[0] + ".xlsx");
            string filePPTX = Path.GetFullPath(args[0] + ".pptx");

            // Sound player
            player = new SoundPlayer();

            // Create speech client
            tts = TextToSpeechClient.Create();

            // Progress message
            Speak("Launching Excel ....", 200);

            // Launch Excel
            exa = new Excel.Application();

            // Progress message
            Speak("Launching PowerPoint ....", 200);

            // Launch powerpoint
            ppa = new PowerPoint.Application(); 

            // Build temporary filenames
            string tempXLSX = Path.Combine(Path.GetDirectoryName(fileXLSX), "~" + Path.GetFileName(fileXLSX));
            string tempPPTX = Path.Combine(Path.GetDirectoryName(filePPTX), "~" + Path.GetFileName(filePPTX));

            // Copy the source files
            File.Copy(fileXLSX, tempXLSX, true);
            File.Copy(filePPTX, tempPPTX, true);

            // Use temporary files (automation is notorious for corrupting files!!!!!)
            fileXLSX = tempXLSX;
            filePPTX = tempPPTX;

            // Open workbook 
            exa.Workbooks.Open(fileXLSX);

            // Make application visible            
            exa.Visible = true;

            // Open presentation 
            ppp = ppa.Presentations.Open(filePPTX);

            // Make application visible
            ppa.Visible = MsoTriState.msoTrue;

            // Start video recorder
            Speak("Now generating speech ....", 200);

            // Parse the script
            PPScript script = ParseScript();

            // Progress message
            Speak("Closing Excel ....", 200);

            // Quit Excel
            try { exa.ActiveWorkbook.Close(); } catch { }
            try { exa.Quit(); } catch { }
            try { Kill("EXCEL", ""); } catch { }

            // Start video recorder
            Speak("Please disable multi monitor software like display fusion");
            Speak("Please also disable all PowerPoint Add-ins");
            Speak("Start your video recorder and press enter");
                      
            // Wait for key
            Console.ReadLine();

            // Execute script
            script.Run(First);

            // Progress message
            Speak("Closing PowerPoint ....", 200);

            // Quit PowerPoint          
            try { ppp.Close(); } catch { }
            try { ppa.Quit();  } catch { }
            try { Kill("POWERPNT", "PowerPoint"); } catch { }

            // Delete temporary files
            File.Delete(fileXLSX);
            File.Delete(filePPTX);

            // Stop video recorder
            Speak("Please stop your video recorder and press enter", 200);
        }

        /// <summary>
        /// Speaks the specified text using an English female voice (system messages).
        /// </summary>
        /// <param name="text">The text.</param>
        static void Speak(string text, int delay = 0)
        {
            Console.WriteLine(text);

            var input = new SynthesisInput
            {
                Text = text
            };

            // Voice selection parameters
            var vspp = new VoiceSelectionParams
            {
                Name = "en-US-Neural2-E",                   // Voice name (male or female)
                LanguageCode = "en-US"                      // Take language code from voice name
            };

            // Specify the type of audio file.
            var ac = new AudioConfig
            {
                AudioEncoding = AudioEncoding.Linear16,         // Wave 
                VolumeGainDb = 10
            };

            // Make the API call
            var response = tts.SynthesizeSpeech(input, vspp, ac);

            // Connect as memory stream to the sound player
            player.Stream = new MemoryStream(response.AudioContent.ToByteArray(), true);

            // Wait
            if (delay > 0) Thread.Sleep(delay);

            // Play
            player.PlaySync();
        }

        /// <summary>
        /// Parses the settings and script from the Excel workbook.
        /// </summary>
        static PPScript ParseScript()
        {
            // Slide counter
            int s = 0;

            // Create a script object
            PPScript script = new PPScript(ppp, player);

            // Initialize the list of actions
            script.Actions = new List<PPAction>();

            var vspp = new VoiceSelectionParams();
            var ac = new AudioConfig();

            // Iterate all worksheets
            foreach (Excel.Worksheet ws in exa.ActiveWorkbook.Worksheets)
            {
                // Settings
                if (ws.Name.StartsWith("Settings", StringComparison.OrdinalIgnoreCase))
                {
                    First = (int)(double)ws.Range["B8"].Value;
                    Last  = (int)(double)ws.Range["B9"].Value;

                    string name   = ws.Range["B3"].Text.ToString();
                    string[] segs = name.Split('-');

                    // Voice selection parameters
                    vspp = new VoiceSelectionParams
                    {
                        Name = name,                                // Voice name (male or female)
                        LanguageCode = segs[0] + "-" + segs[1]      // Take language code from voice name
                    };

                    // Specify the type of audio file.
                    ac = new AudioConfig
                    {
                        AudioEncoding = AudioEncoding.Linear16,         // Wave 
                        VolumeGainDb = (double)ws.Range["B4"].Value,    // Volume gain in Db  (-96 .. +16)
                        Pitch = (double)ws.Range["B5"].Value,           // Pitch in semitones (-20 .. +20)
                        SpeakingRate = (double)ws.Range["B6"].Value     // Rate factor        (1/4 ..  4 )
                    };
                }
                // Slides
                else if (ws.Name.StartsWith("Slide", StringComparison.OrdinalIgnoreCase))
                {
                    int i = 0; s++;
                    if (s >= First && s <= Last)
                    {
                        while (ws.Range["B2"].Offset[i, 0].Text.ToString() != "")
                        {
                            // Get delay and argument
                            int del = (int)(double)ws.Range["A2"].Offset[i, 0].Value;
                            string arg = ws.Range["B2"].Offset[i, 0].Text.ToString();

                            // Create and add the action
                            PPAction action = new PPAction(del, arg);
                            script.Actions.Add(action);

                            // Prefetch voices
                            if (action.Com == EPCommand.Speak)
                            {
                                Console.WriteLine(arg);
                                var input = new SynthesisInput { Text = arg };
                                var response = tts.SynthesizeSpeech(input, vspp, ac);
                                action.Wav = new MemoryStream(response.AudioContent.ToByteArray(), true);
                            }

                            // Increment action counter
                            i++;
                        }
                    }
                }
            }

            // Return the script
            return script;
        }

        /// <summary>
        /// Kills the process indentified by its signature.
        /// </summary>
        /// <param name="signature"></param>
        static void Kill(string signature, string title)
        {
            // Create an array of all running Excel processes
            Process[] processes = Process.GetProcessesByName(signature);

            // Loop over these processes
            foreach (var process in processes)
            {
                // Only look at the instance with an empty window title
                if (process.MainWindowTitle == title)
                {
                    // Kill the process
                    process.Kill();
                }
            }
        }
    }
}
