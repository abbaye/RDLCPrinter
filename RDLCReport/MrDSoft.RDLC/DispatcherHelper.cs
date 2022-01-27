using System.Windows.Threading;

namespace TimePunch.Rdlc
{
    /// <summary>
    /// WPF UI Dispatcher
    /// <remarks>
    /// CREDIT : Alex Petrescu 2010, Modified 2014 and 2018 by Derek Tremblay
    /// PROFIL : http://social.msdn.microsoft.com/profile/alex%20petrescu/?ws=usercard-inline
    /// http://social.msdn.microsoft.com/Forums/fr-FR/248a3258-e3ec-4ba2-9085-2fda2f0b0058/wpf-faq-applicationdoevents-dans-wpf?forum=wpffr
    /// </remarks>
    /// </summary>
    public static class DispatcherHelper
    {
        private static readonly DispatcherOperationCallback ExitFrameCallback = ExitFrame;

        /// <summary>
        /// Execute all message in message Queud
        /// </summary>
        public static void DoEvents(DispatcherPriority priority = DispatcherPriority.Background)
        {
            var nestedFrame = new DispatcherFrame();
            var exitOperation = Dispatcher.CurrentDispatcher.BeginInvoke(priority, ExitFrameCallback, nestedFrame);

            try
            {
                //execute all next message
                Dispatcher.PushFrame(nestedFrame);

                //If not completed, will stop it
                if (exitOperation.Status != DispatcherOperationStatus.Completed)
                    exitOperation.Abort();
            }
            catch
            {
                exitOperation.Abort();
            }

        }

        private static object ExitFrame(object state)
        {
            // exit the message loop
            if (state is DispatcherFrame frame)
                frame.Continue = false;

            return null;
        }
    }
}
