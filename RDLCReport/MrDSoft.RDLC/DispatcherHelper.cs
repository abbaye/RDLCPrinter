using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace DSoft.RDLCReport
{
    /// <summary>
    /// WPF UI Dispatcher
    /// <remarks>
    /// CREDIT : Alex Petrescu 2010 
    /// PROFIL : http://social.msdn.microsoft.com/profile/alex%20petrescu/?ws=usercard-inline
    /// http://social.msdn.microsoft.com/Forums/fr-FR/248a3258-e3ec-4ba2-9085-2fda2f0b0058/wpf-faq-applicationdoevents-dans-wpf?forum=wpffr
    /// </remarks>
    /// </summary>
    public class DispatcherHelper
    {
        private static DispatcherOperationCallback exitFrameCallback = new DispatcherOperationCallback(ExitFrame);
        
        /// <summary>
        /// Execute all message in message Queud
        /// </summary>
        public static void DoEvents()
        {            
            DispatcherFrame nestedFrame = new DispatcherFrame();
            
            DispatcherOperation exitOperation = Dispatcher.CurrentDispatcher.BeginInvoke( DispatcherPriority.Background, exitFrameCallback, nestedFrame);

            //execute all next message
            Dispatcher.PushFrame(nestedFrame);

            //If not completed, will stop it
            if (exitOperation.Status != DispatcherOperationStatus.Completed)
                exitOperation.Abort();
            
        }

        private static Object ExitFrame(Object state)
        {
            DispatcherFrame frame = state as DispatcherFrame;

            // exit the message loop
            frame.Continue = false;
            return null;
        }
    }
}
