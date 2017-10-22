using System;

namespace VisualStudioLauncher
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                var launcher = new Launcher();

                try
                {
                    launcher.ParseArguments(args);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e.Message);
                    Console.Error.WriteLine(e.StackTrace);
                    ExitWithCode(403);
                }

                launcher.ExecuteParameterCommand();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
                Console.Error.WriteLine(e.StackTrace);
                ExitWithCode(500);
            }
        }


        static void ExitWithCode(int errorCode)
        {
            Environment.Exit(errorCode);
        }


    }
}