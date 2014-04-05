using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace DSoft.RDLCReport
{
    /// <summary>
    /// encapsule un répartiteur wpf avec des fonctionalites supplementaires.
    /// </summary>
    public class DispatcherHelper
    {
        private static DispatcherOperationCallback exitFrameCallback = new
           DispatcherOperationCallback(ExitFrame);

        /// <summary>
        /// Traite tous les messages de l'IU actuellement dans la file d'attente de message. 
        /// </summary>
        public static void DoEvents()
        {
            // creer une nouvelle pompe de messages imbriques.
            DispatcherFrame nestedFrame = new DispatcherFrame();

            //Envoie un callback a la file de messages
            //lorsqu’il est appelle, ce callback arrêtera la boucle de messages imbriques
            //la priorité de ce callback doit être inferieure à celle des messages d’évènement IU
            DispatcherOperation exitOperation = Dispatcher.CurrentDispatcher.BeginInvoke(
              DispatcherPriority.Background, exitFrameCallback, nestedFrame);

            // Avancer la boucle de messages, la boucle de messages
            //va traiter les messages restés a l’intérieur du fil des messages
            Dispatcher.PushFrame(nestedFrame);

            //si le callback ExitFrame n’a pas fini, on l’arrete
            if (exitOperation.Status != DispatcherOperationStatus.Completed)
            {
                exitOperation.Abort();
            }
        }

        private static Object ExitFrame(Object state)
        {
            DispatcherFrame frame = state as DispatcherFrame;

            // Sort de la boucle de messages
            frame.Continue = false;
            return null;
        }
    }
}
