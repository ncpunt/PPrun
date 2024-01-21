namespace PPrun
{
    public enum EPCommand
    {
        Pause = 0,
        Click = 1,
        Speak = 2,
        Stop  = 99
    }

    public class PPAction
    {
        public int          Del;   // Delay in msec
        public EPCommand    Com;   // Command inferred from argument
        public string       Arg;   // Command argument
        public MemoryStream Wav;   // Audio in Linear16 format

        public PPAction(int del, string arg)
        {
            Del = del;
            Arg = arg;

            switch (Arg[0])
            {
                case '#': Com = EPCommand.Click; break;
                case '@': Com = EPCommand.Pause; break;
                case '~': Com = EPCommand.Stop;  break;
                default : Com = EPCommand.Speak; break; 
            }
        }
    }
}
