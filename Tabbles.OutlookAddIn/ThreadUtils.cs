using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace Tabbles.OutlookAddIn
{
    public static class ThreadUtils
    {

        /*
        let execInThread f =
        if Threading.Thread.CurrentThread.IsBackground then // modo rapido per capire se siamo già in un thread
                f ()
        else
                let body () = 
                        try
                                f()
                        with
                        | ecc  ->
                             showCrashDialog ecc None "execInThread - "
                (  // faccio partire il thread
                let th = new System.Threading.Thread(body)
                th.CurrentUICulture <-  new Globalization.CultureInfo(g_lang) 
                th.Priority <- Threading.ThreadPriority.Normal
                th.IsBackground <- true
                th.Start ()                
                )

         * 
         * */


        /// <summary>
        /// Execute a given piece of code in background, in a thread, in order to return quickly.
        /// </summary>
        /// <param name="a"></param>
        public static void execInThread( Action a)
        {
            if (System.Threading.Thread.CurrentThread.IsBackground)
            {
                // we are already in a thread, no need to start another.

                a.Invoke();
            }
            else
            {

                var th = new Thread(new ThreadStart(() =>
                {
                    try
                    {
                        a.Invoke();
                    }
                    catch (Exception e)
                    {
                        Log.log(">>> execInThread: exception:" + e.GetType().ToString() + ", " + e.Message);
                    }

                }));

                // th.CurrentUICulture = ...
                th.Priority = ThreadPriority.Normal;
                th.IsBackground = true;
                th.Start();
            }

        }


        public static void execInThreadForceNewThread(Action a)
        {

            var th = new Thread(new ThreadStart(() =>
            {
                try
                {
                    a.Invoke();
                }
                catch (Exception e)
                {
                    Log.log(">>> execInThreadForceNewThread: exception:" + e.GetType().ToString() + ", " + e.Message);
                }

            }));

            // th.CurrentUICulture = ...
            th.Priority = ThreadPriority.Normal;
            th.IsBackground = true;
            th.Start();

        }
    }
}
