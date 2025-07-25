using System;

namespace Printune
{
    class Program
    {
        static int Main(string[] args)
        {
            HelpInvocation.Register();
            DriverInvocation.Register();
            PrinterInvocation.Register();
            IntunePackageInvocation.Register();
            VerifyInvocation.Register();

            try
            {
                string log = null;
                if (ParameterParser.GetFlag(args, "-Log"))
                {
                    ParameterParser.GetParameterValue(args, "-Log", out log);
                    Log.Initialize(log);
                }

                var invocation = Invocation.Parse(args);
                return invocation.Invoke();
            }
            catch (Exception ex)
            {
                var exit = new HelpInvocation(ex).Invoke();
                Console.ResetColor();
                return exit;
            }
        }
    }
}