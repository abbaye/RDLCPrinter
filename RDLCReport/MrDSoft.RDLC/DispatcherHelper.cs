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
