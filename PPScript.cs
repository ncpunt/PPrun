using Microsoft.Office.Interop.PowerPoint;
using System.Media;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPrun
{
    public delegate void Notify();  // Notifiation event

    public class PPScript
    {
        public List<PPAction>           Actions;    // List of actions
        public SoundPlayer              Player;     // Sound player
        public PowerPoint.SlideShowView SSV;        // PowerPoint slide show view
        PowerPoint.Presentation         PPP;        // PowerPoint presentation

        public event Notify PresentationStarted;    // Fired when in presentation mode

        public PPScript(PowerPoint.Presentation ppp, SoundPlayer player)
        {
            PPP    = ppp;
            Player = player;
        }

        public void Run(int first = 1)
        {
            // Start the presentation and get the View object
            SSV = PPP.SlideShowSettings.Run().View;

            // Hide the media controls
            PPP.SlideShowSettings.ShowMediaControls = Microsoft.Office.Core.MsoTriState.msoFalse;

            // Move to the starting slide
            PPP.SlideShowWindow.View.GotoSlide(first);

            // Fire notification
            PresentationStarted?.Invoke();

            // Iterate all actions
            foreach (var action in Actions) 
            {
                // Was a key pressed?
                if (Console.KeyAvailable)
                {
                    // Get that key
                    ConsoleKey key = Console.ReadKey(true).Key;

                    // Process the key
                    if (key == ConsoleKey.Spacebar)
                    {
                        // Pause script
                        Console.WriteLine("Press any key to continue.");
                        Console.ReadKey();
                    }
                    else if (key == ConsoleKey.Escape) 
                    {
                        // Abort script
                        Console.WriteLine("Presentation was aborted.");
                        break;
                    }
                }

                // All action can be preceded wit a delay
                if (action.Del > 0) Thread.Sleep(action.Del);

                // Execute action
                switch (action.Com)
                {
                    case EPCommand.Click : SSV.Next();         break;
                    case EPCommand.Speak : Speak(action);      break;
                    case EPCommand.Pause : /* do nothing */    break;
                    case EPCommand.Stop  : /* exit       */    return;
                }
            }

            // Stop the slide show
            SSV.Exit();
        }

        public void Speak(PPAction action)
        {
            // Show progress
            Console.WriteLine(action.Arg);

            // Connect as memory stream to the sound player
            Player.Stream = action.Wav;

            // Play
            Player.PlaySync();
        }
    }
}
