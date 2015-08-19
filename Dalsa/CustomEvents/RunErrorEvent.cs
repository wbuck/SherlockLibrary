using System;
using IpeEngCtrlLib;

namespace Dalsa.CustomEvents
{
    namespace Sherlock
    {
        /// <summary>
        /// Occurs when an error stops Sherlock's execution.
        /// </summary>
        /// <param name="sender">Reference to object that raised the event.</param>
        /// <param name="e">Contains error status information.</param>
        public delegate void RunErrorEventHandler( object sender, RunErrorEventArgs e );

        /// <summary>
        /// Custom event class for raising a Sherlock error message.
        /// </summary>
        public class RunErrorEventArgs : EventArgs
        {
            // Holds the error message sent from Sherlock.
            private readonly string _sherlockError;

            /// <summary>
            /// Custom event class for raising a Sherlock error message.
            /// </summary>
            public RunErrorEventArgs( I_EXEC_ERROR sherlockError )
            {
                _sherlockError = sherlockError.ToString( );
            }
            /// <summary>
            /// Contains the error message sent from Sherlock.
            /// </summary>
            /// <returns>Error message sent from Sherlock.</returns>
            public string ErrorMessage
            {
                get { return _sherlockError; }
            }
        }
    }
}
