using System;

namespace Dalsa.CustomEvents
{
    namespace Sherlock
    {
        /// <summary>
        /// Occurs when a user programmed event from Sherlock is raised.
        /// </summary>
        /// <param name="sender">Reference to object that raised the event.</param>
        /// <param name="e">Contains the event ID of the raised event.</param>
        public delegate void UserDefinedEventHandler( object sender, UserDefinedEventArgs e );

        /// <summary>
        /// Custom event class for raising a user define event.
        /// </summary>
        public class UserDefinedEventArgs : EventArgs
        {
            // The integar value comes from Sherlock.
            // It lets the developeer know which
            // user defined event was raised in Sherlock.
            private readonly int _eventId;

            /// <summary>
            /// Custom event class for raising a user defined event.
            /// </summary>
            /// <param name="eventId">An integer value to denote which user programmed event fired in Sherlock.</param>
            public UserDefinedEventArgs( int eventId )
            {
                _eventId = eventId;
            }

            /// <summary>
            /// A value indicating which user event was raised in Sherlock.
            /// </summary>
            /// <returns>An integer value to denote which user programmed event fired in Sherlock.</returns>
            public int EventId
            {
                get { return _eventId; }
            }
        }
    }
}