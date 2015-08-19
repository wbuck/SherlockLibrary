using System;

namespace Dalsa.CustomEvents
{
    namespace Sherlock
    {
        /// <summary>
        /// Occurs when the reporter within Sherlock updates.
        /// </summary>
        /// <param name="sender">Reference to object that raised the event.</param>
        /// <param name="e">Contains a message sent by the reporter within Sherlock.</param>
        public delegate void ReporterUpdateEventHandler( object sender, ReporterUpdateEventArgs e );

        /// <summary>
        /// 
        /// </summary>
        public class ReporterUpdateEventArgs : EventArgs
        {
            // Holds message sent from reporter.
            private readonly string _reporterMessage;

            /// <summary>
            /// Occurs when the reporter within Sherlock updates.
            /// </summary>
            /// <param name="reporterMessage">Contains a message sent by the reporter within Sherlock.</param>
            public ReporterUpdateEventArgs( string reporterMessage )
            {
                _reporterMessage = reporterMessage;
            }

            /// <summary>
            /// Occurs when the reporter within Sherlock updates.
            /// </summary>
            /// <returns>Contains a message sent by the reporter within Sherlock.</returns>
            public string ReporterMessage
            {
                get { return _reporterMessage; }
            }
        }
    }
}
