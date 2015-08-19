using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using AxIpeDspCtrlLib;
using Dalsa.CustomEvents.Sherlock;
using IpeEngCtrlLib;
using Timer = System.Windows.Forms.Timer;

namespace Dalsa
{
    namespace Sherlock
    {
        /// <summary>
        /// A wrapper for the Sherlock API which simplifies it's use. 
        /// Class cannot be inherited.
        /// </summary>
        public sealed class SherlockEngine
        {
            #region Private Variables

            // Sherlock engine.
            private static Engine _sherlock;

            // Sherlock error status return.        
            private static I_ENG_ERROR _iReturn;

            // Current sherlock mode.       
            private static I_MODE _iCurrMode;

            // This timer is used to poll the mode that Sherlock is in.
            // A Forms.Timer was used instead of a timer.timer to
            // prevent any issues caused by multi threading.
            private static Timer _timer;

            // Variable will be used to store the context of 
            // the main UI thread.
            // It will allow us to execute a section of code in the UI context 
            // so that GUI controls can be manipulated.
            private SynchronizationContext _synchronizationContext;

            #endregion

            #region Constructors
            /// <summary>
            /// Initializes Sherlock engine.
            /// </summary>
            /// <exception cref="ExternalException"></exception>
            public SherlockEngine( )
            {
                // Set the timer to check Sherlocks mode of operation.
                _timer = new Timer { Interval = 500 };

                // Hook up tick event.
                _timer.Tick += _timer_Tick;
            }
            #endregion

            #region Private Delegate Variables

            // Delegate variable used to determine if the event
            // has subscribers.
            private UserDefinedEventHandler _userDefinedHandler;

            // Delegate variable used to determine if the event
            // has subscribers.
            private EventHandler _runCompletedHandler;

            // Delegate variable used to determine if the event
            // has subscribers.
            private RunErrorEventHandler _runErrorHandler;

            // Delegate variable used to determine if the event
            // has subscribers.
            private ReporterUpdateEventHandler _reporterUpdateHandler;


            #endregion

            #region Event Delegates
            /// <summary>
            /// Occurs when Sherlock has halted.
            /// </summary>
            public event EventHandler Halted;

            /// <summary>
            /// Occurs when Sherlock has been set to run continuously.
            /// </summary>
            public event EventHandler RunningContinuous;

            /// <summary>
            /// Occurs when Sherlock has been set to run through one iteration.
            /// </summary>
            public event EventHandler RanOnce;

            /// <summary>
            /// Occurs when Sherlock has successfully loaded an investigation.
            /// </summary>
            public event EventHandler InvestigationLoaded;

            /// <summary>
            /// Occurs when a user programmed event from Sherlock is raised.
            /// </summary>
            public event UserDefinedEventHandler UserDefinedEvent
            {
                add
                {
                    _userDefinedHandler += value;
                    // Only subscribe the the Dalsa supplied event once.
                    if( _userDefinedHandler.GetInvocationList( ).Length == 1 )
                        // Subscribe to the Dalse supplied event.
                        _sherlock.UserProgramEvent += OnDalsaUserDefinedEvent;
                }
                remove
                {
                    if( _userDefinedHandler == null )
                        return;

                    _userDefinedHandler -= value;
                    // Unsubscribe from Dalsa event if there are no subscribers
                    // to the custom UserDefined event.
                    if( _userDefinedHandler == null )
                        _sherlock.UserProgramEvent -= OnDalsaUserDefinedEvent;
                }
            }

            /// <summary>
            /// Occurs when Sherlock execution has run to completion.
            /// </summary>
            public event EventHandler RunCompleted
            {
                add
                {
                    _runCompletedHandler += value;
                    // Only subscribe the the Dalsa supplied event once.
                    if( _runCompletedHandler.GetInvocationList( ).Length == 1 )
                        // Subscribe to the Dalse supplied event.
                        _sherlock.RunCompleted += OnDalsaRunCompletedEvent;
                }
                remove
                {
                    if( _runCompletedHandler == null )
                        return;

                    _runCompletedHandler -= value;

                    if( _runCompletedHandler == null )
                        _sherlock.RunCompleted -= OnDalsaRunCompletedEvent;
                }
            }

            /// <summary>
            /// Occurs when the reporter within Sherlock updates.
            /// </summary>
            public event ReporterUpdateEventHandler ReporterUpdateEvent
            {
                add
                {
                    _reporterUpdateHandler += value;
                    if ( _reporterUpdateHandler.GetInvocationList( ).Length == 1 )
                        _sherlock.ReporterDisplay += OnDalsaReporterUpdate;
                }
                remove
                {
                    if ( _reporterUpdateHandler == null )
                        return;

                    _reporterUpdateHandler -= value;

                    if ( _reporterUpdateHandler == null )
                        _sherlock.ReporterDisplay -= OnDalsaReporterUpdate;
                }
            }

            /// <summary>
            /// Occurs when an error stops Sherlock's execution.
            /// </summary>
            public event RunErrorEventHandler RunError
            {
                add
                {
                    _runErrorHandler += value;
                    // Only subscribe to the Dalsa supplied event once.
                    if( _runErrorHandler.GetInvocationList( ).Length == 1 )
                        // Subscribe to the Dalse supplied event.
                        _sherlock.OnRunError += OnDalsaRunErrorEvent;
                }
                remove
                {
                    if( _runErrorHandler == null )
                        return;

                    _runErrorHandler -= value;

                    if( _runErrorHandler == null )
                        _sherlock.OnRunError -= OnDalsaRunErrorEvent;
                }

            }
            #endregion

            #region Enums
            /// <summary>
            /// Enumeration of Sherlock modes of operation.
            /// </summary>
            public enum Mode
            {
                /// <summary>
                /// Unspecified Sherlock mode of operation.
                /// </summary>
                Error = -1,
                /// <summary>
                /// Halt Sherlock operation immediately.
                /// </summary>
                Halt = 1,
                /// <summary>
                /// Run Sherlock continuously.
                /// </summary>
                Continuous = 2,
                /// <summary>
                /// Run Sherlock once and halt on program completion.
                /// </summary>
                RunOnce = 4,
            }


            /// <summary>
            /// Enumeration of different magnification settings for the Sherlock display control.
            /// </summary>
            public enum Magnification
            {
                /// <summary>
                /// Fit image to display control size.
                /// </summary>
                ImageFit,
                /// <summary>
                /// One times zoom.
                /// </summary>
                Image1X,
                /// <summary>
                /// Two times zoom.
                /// </summary>
                Image2X,
                /// <summary>
                /// Three times zoom.
                /// </summary>
                Image3X,
                /// <summary>
                /// Four times zoom.
                /// </summary>
                Image4X,
                /// <summary>
                /// Five times zoom.
                /// </summary>
                Image5X,
                /// <summary>
                /// Sixe times zoom.
                /// </summary>
                Image6X,
                /// <summary>
                /// Seven times zoom.
                /// </summary>
                Image7X
            }
            #endregion

            #region Public Properties
            /// <summary>
            /// Gets Sherlocks mode of operation.
            /// </summary>
            public Mode SherlockMode
            {
                get
                {
                    _iReturn = _sherlock.InvModeGet( out _iCurrMode );
                    switch( _iCurrMode )
                    {
                        case I_MODE.I_EXE_MODE_HALT:
                            return Mode.Halt;

                        case I_MODE.I_EXE_MODE_CONT:
                            return Mode.Continuous;

                        case I_MODE.I_EXE_MODE_ONCE:
                            return Mode.RunOnce;

                        default:
                            return Mode.Error;
                    }
                }
            }
            #endregion

            #region Public Methods
            /// <summary>
            /// Set Sherlocks mode to continuous.
            /// </summary>
            public void Continuous( )
            {
                _sherlock.InvModeSet( I_MODE.I_EXE_MODE_CONT );
                // Fire run continuous event for immediate updating of form controls.
                OnRunningContinuous( this, EventArgs.Empty );
            }

            /// <summary>
            /// Set Sherlocks mode to run once.
            /// </summary>
            public void RunOnce( )
            {
                _sherlock.InvModeSet( I_MODE.I_EXE_MODE_ONCE );
                // Fire run once event for immediate updating of form controls.
                OnRanOnce( this, EventArgs.Empty );
            }

            // Halt Sherlocks mode of operation. In order to avoid
            // using Dalsa's method of halting the investigation 
            // which involed using Application.DoEvents() I set and
            // wait for Sherlock to halt in another task. This prevents
            // the main UI thread from deadlocking and not updating _iCurrMode.
            /// <summary>
            /// Set Sherlocks mode to halt.
            /// </summary>
            public void Halt( )
            {
                // Halt Sherlock and wait for halted state in new task.
                Task.Factory.StartNew( ( ) =>
                {
                    _sherlock.InvModeSet( I_MODE.I_EXE_MODE_HALT );
                    // Remain in loop until Sherlock has halted.
                    while( _iCurrMode != I_MODE.I_EXE_MODE_HALT )
                    {
                        _iReturn = _sherlock.InvModeGet( out _iCurrMode );
                        // Giving the cpu a second to catch its breath.
                        Thread.Sleep( 1000 );
                    }
                } );
            }

            /// <summary>
            /// Initializes Sherlock and then loads the supplied investigation.
            /// </summary>
            /// <param name="investigationPath">Path and name of Sherlock investigation to be loaded.</param>
            /// <exception cref="ArgumentNullException"></exception>
            /// <exception cref="DirectoryNotFoundException"></exception>
            /// <exception cref="ExternalException"></exception>
            public void InitializeAndLoadInvestigation( string investigationPath )
            {
                if( investigationPath == null )
                {
                    throw new ArgumentNullException(
                        "investigationPath",
                        "The value of \"string investigation\" passed to the \"LoadInvestigation( string investigation )\" method was null." );
                }

                if( !File.Exists( investigationPath ) )
                {
                    throw new DirectoryNotFoundException(
                        string.Format( "The supplied Sherlock path {0} could not be found", investigationPath ) );
                }

                // Get UI context to marshal messages from
                // a worker thread back to the main UI thread.
                _synchronizationContext = SynchronizationContext.Current;

                // Because loading the Sherlock investigation takes so long
                // the UI would deadlock during the loading process. I now
                // instantiate the sherlock engine and load the investigation
                // in another thread to prevent the UI from deadlocking.
                Task.Factory.StartNew( ( ) =>
                {
                    InitializeEngine( );
                    _iReturn = _sherlock.InvLoad( investigationPath );

                } ).ContinueWith( task =>
                {
                    // Execute code in the main UI context.
                    _synchronizationContext.Post( callback => AfterInitializationAndInvestigationLoad( ), null );
                } );
            }

            /// <summary>
            /// Compares the currently loaded investigation name to the supplied investigation name.
            /// </summary>
            /// <param name="investigation">The name of the investigation to be compared.</param>
            /// <returns>A boolean result is returned based on the comparison.</returns>
            /// <exception cref="ArgumentNullException"></exception>
            public bool CompareInvestigation( string investigation )
            {
                if( investigation == null )
                {
                    throw new ArgumentNullException(
                        "investigation",
                        "The value of \"string investigation\" passed to the \"CompareInvestigation( string investigation )\" method was null." );
                }

                string isLoaded;
                _sherlock.InvGetName( out isLoaded );

                // Parse out the investigation name from the file path.
                var inDex = investigation.LastIndexOf( @"\", investigation.Length, StringComparison.Ordinal );
                var toLoad = investigation.Substring( inDex + 1 );

                return string.CompareOrdinal( isLoaded, toLoad ) == 0;
            }

            /// <summary>
            /// Connects the AxIpeDspCtrl image control to an image window in Sherlock.
            /// </summary>
            /// <param name="displayName">The name of the image window in Sherlock.</param>
            /// <param name="imgDisplay">A reference to the AxeIpeDspCtrl display control used.</param>
            public void ConnectDisplay( string displayName, AxIpeDspCtrl imgDisplay )
            {
                imgDisplay.ConnectEngine( _sherlock.GetEngineObj( ) );
                imgDisplay.ConnectImgWindow( displayName );
            }

            /// <summary>
            /// Connects the AxIpeDspCtrl image control to an image window in Sherlock.
            /// </summary>
            /// <param name="displayName">The name of the image window in Sherlock.</param>
            /// <param name="imgDisplay">A reference to the AxeIpeDspCtrl display control used.</param>
            /// <exception cref="ArgumentNullException"></exception>
            public void ConnectDisplay( string displayName, object imgDisplay )
            {
                var display = imgDisplay as AxIpeDspCtrl;

                if( display == null )
                    throw new ArgumentNullException( "imgDisplay",
                        "The value of \"object imgDisplay\" cannot be null" );

                display.ConnectEngine( _sherlock.GetEngineObj( ) );
                display.ConnectImgWindow( displayName );
            }

            /// <summary>
            /// Connects the AxIpeDspCtrl image control to an image window in Sherlock and then sets the magnification.
            /// </summary>
            /// <param name="displayName">The name of the image window in Sherlock.</param>
            /// <param name="imgDisplay">A reference to the AxeIpeDspCtrl display control used.</param>
            /// <param name="magnification">Sets the inital magnification of the AxIpeDspCtrl control.</param>
            public void ConnectDisplay( string displayName, AxIpeDspCtrl imgDisplay, Magnification magnification )
            {
                imgDisplay.ConnectEngine( _sherlock.GetEngineObj( ) );
                imgDisplay.ConnectImgWindow( displayName );
                imgDisplay.SetZoom( ( double )magnification );
            }

            /// <summary>
            /// Sets the magnification of the AxIpeDspCtrl image control.
            /// </summary>
            /// <param name="imgDisplay">A reference to the AxeIpeDspCtrl display control used.</param>
            /// <param name="magnification">Sets the magnification of the AxIpeDspCtrl control.</param>
            public void SetZoom( AxIpeDspCtrl imgDisplay, Magnification magnification )
            {
                imgDisplay.SetZoom( ( double )magnification );
            }

            /// <summary>
            /// Disconnect the AxIpeDspCtrl image control from the image window in Sherlock.
            /// </summary>
            /// <param name="imgDisplay">A reference to the AxeIpeDspCtrl display control to be disconnected.</param>
            public void DisconnectDisplay( AxIpeDspCtrl imgDisplay )
            {
                imgDisplay.DisconnectImgWindow( );
            }

            /// <summary>
            /// Disconnect the AxIpeDspCtrl image control from the image window in Sherlock.
            /// </summary>
            /// <param name="imgDisplay">A reference to the AxeIpeDspCtrl display control to be disconnected.</param>
            /// <exception cref="ArgumentNullException"></exception>
            public void DisconnectDisplay( object imgDisplay )
            {
                var display = imgDisplay as AxIpeDspCtrl;
                if( display == null )
                    throw new ArgumentNullException( "imgDisplay",
                        "The value of \"object imgDisplay\" cannot be null" );

                display.DisconnectImgWindow( );

            }

            /// <summary>
            /// Generic method for retrieving a Sherlock variables value.
            /// </summary>
            /// <typeparam name="T">The Sherlock variables type.</typeparam>
            /// <param name="variableName">The name of the variable in Sherlock.</param>
            /// <returns>The Sherlock variable value.</returns>
            /// <exception cref="ArgumentException"></exception>
            public T GetVariable<T>( string variableName )
            {
                object retVal;

                if( typeof( T ) == typeof( double ) )
                {
                    double tempReturnVal;
                    _iReturn = _sherlock.VarGetDouble( variableName, out tempReturnVal );

                    retVal = Convert.ChangeType( tempReturnVal, typeof( T ) );
                }
                else if( typeof( T ) == typeof( string ) )
                {
                    string tempReturnVal;
                    _iReturn = _sherlock.VarGetString( variableName, out tempReturnVal );

                    retVal = Convert.ChangeType( tempReturnVal, typeof( T ) );
                }
                else if( typeof( T ) == typeof( bool ) )
                {
                    bool tempReturnVal;
                    _iReturn = _sherlock.VarGetBool( variableName, out tempReturnVal );

                    retVal = Convert.ChangeType( tempReturnVal, typeof( T ) );
                }
                else if( typeof( T ) == typeof( double[ ] ) )
                {
                    Array tempReturnVal;
                    _iReturn = _sherlock.VarGetDoubleArray( variableName, out tempReturnVal );

                    retVal = Convert.ChangeType( tempReturnVal, typeof( T ) );
                }
                else if( typeof( T ) == typeof( string[ ] ) )
                {
                    Array tempReturnVal;
                    _iReturn = _sherlock.VarGetStringArray( variableName, out tempReturnVal );

                    retVal = Convert.ChangeType( tempReturnVal, typeof( T ) );
                }
                else if ( typeof( T ) == typeof( bool[ ] ) )
                {
                    Array tempReturnVal;
                    _iReturn = _sherlock.VarGetBoolArray( variableName, out tempReturnVal );

                    retVal = Convert.ChangeType( tempReturnVal, typeof( T ) );
                }
                else
                {
                    throw new ArgumentException(
                        string.Concat( "The type ", typeof( T ).ToString( ), " is not supported" ) );
                }

                if ( _iReturn != I_ENG_ERROR.I_OK )
                {
                    throw new ArgumentException(
                        string.Concat(
                            "An issue was encountered while attemping to retrieve the value of the \"",
                            variableName, "\" variable. Sherlock returned an error message of: ", _iReturn.ToString( ) ) );
                }
                return ( T )retVal;
            }

            /// <summary>
            /// Generic method for setting a Sherlock variable value.
            /// </summary>
            /// <typeparam name="T">The Sherlock variable type.</typeparam>
            /// <param name="value">The value to set the Sherlock variable to.</param>
            /// <param name="variableName">The name of the variable in Sherlock.</param>
            /// <exception cref="ArgumentException"></exception>
            public void SetVariable<T>( string variableName, T value )
            {
                if ( value is double || value is int )
                {
                    _iReturn = _sherlock.VarSetDouble( variableName, Convert.ToDouble( value ) );
                }
                else if ( value is string )
                {
                    _iReturn = _sherlock.VarSetString( variableName, Convert.ToString( value ) );
                }
                else if ( value is bool )
                {
                    _iReturn = _sherlock.VarSetBool( variableName, Convert.ToBoolean( value ) );
                }
                else if ( value is double[ ] )
                {
                    _iReturn = _sherlock.VarSetDoubleArray( variableName, ( double[ ] )( ( object )value ) );
                }
                else if ( value is string[ ] )
                {
                    _iReturn = _sherlock.VarSetStringArray( variableName, ( string[ ] )( ( object )value ) );
                }
                else if ( value is bool[ ] )
                {
                    _iReturn = _sherlock.VarSetBoolArray( variableName, ( bool[ ] )( ( object )value ) );
                }
                else
                {
                    throw new ArgumentException(
                        string.Concat( "The type ", typeof( T ).ToString( ), " is not supported" ) );
                }

                if ( _iReturn != I_ENG_ERROR.I_OK )
                {
                    throw new ArgumentException(
                        string.Concat(
                            "An issue was encountered while attemping to retrieve the value of the \"",
                            variableName, "\" variable. Sherlock returned an error message of: ", _iReturn.ToString( ) ) );
                }
            }

            /// <summary>
            /// Saves the Sherlock investigation.
            /// </summary>
            /// <param name="investigationPath">The full path and name of the investigation including the .ivs file extention.</param>
            /// <exception cref="ArgumentException"></exception>
            public void SaveInvestigationChanges( string investigationPath )
            {
                _iReturn = _sherlock.InvSave( investigationPath );

                if ( _iReturn != I_ENG_ERROR.I_OK )
                {
                    throw new ArgumentException(
                        string.Format( "An error was encountered while save the investigation. Sherlock return: {0}",
                            _iReturn ) );
                }
            }


            /// <summary>
            /// Disposes of the Sherlock engine. Must be called prior to application termination.
            /// </summary>
            public void DisposeSherlock( )
            {
                CleanUp( );
            }

            #endregion

            #region Private Methods
            // The tick event handler checks to see what the current mode Sherlock is in.
            // I did this to ensure the Sherlocks mode is always known and is not
            // dependant on the developer explicitly checking the current mode.
            private void _timer_Tick( object sender, EventArgs e )
            {
                // Stop timer to prevent tick events while
                // handling current tick event.
                _timer.Stop( );
                switch ( SherlockMode )
                {
                    case Mode.Continuous:
                        OnRunningContinuous( this, EventArgs.Empty );
                        break;

                    case Mode.Halt:
                        OnHalted( this, EventArgs.Empty );
                        break;

                    case Mode.RunOnce:
                        OnRanOnce( this, EventArgs.Empty );
                        break;
                }
            }

            // Raise continuous event method.
            private void OnRunningContinuous( object sender, EventArgs e )
            {
                // There is a possibility that the event could be cleared (by code executing in another thread) 
                //between the test for null and the line that actually raises the event. 
                //This scenario constitutes a race condition. So I create, test, 
                //and raise a copy of the event's event handler (delegate).
                var handler = RunningContinuous;
                if ( handler != null )
                {
                    // Re-enable the timer.
                    _timer.Start( );
                    handler( sender, e );
                }
            }

            // Raise event once Sherlock has loaded an investigation.
            private void OnInvestigationLoaded( object sender, EventArgs e )
            {
                var handler = InvestigationLoaded;
                if( handler != null )
                {
                    // Re-enable the timer.
                    _timer.Start( );
                    handler( sender, e );
                }
            }

            // Raise error event method.
            private void OnRunError( object sender, RunErrorEventArgs e )
            {
                var handler = _runErrorHandler;
                if( handler != null )
                {
                    // Re-enable the timer.
                    _timer.Start( );
                    handler( sender, e );
                }
            }

            // Raise halt event method.
            private void OnHalted( object sender, EventArgs e )
            {
                var handler = Halted;
                if( handler != null )
                {
                    // Re-enable the timer.
                    _timer.Start( );
                    handler( sender, e );
                }
            }

            // Raise RunOnce event method.
            private void OnRanOnce( object sender, EventArgs e )
            {
                var handler = RanOnce;
                if ( handler != null )
                {
                    // Re-enable the timer.
                    _timer.Start( );
                    handler( sender, e );
                }
            }

            // Raise UserDefined event method.
            private void OnUserDefinedEvent( object sender, UserDefinedEventArgs e )
            {
                var handler = _userDefinedHandler;

                if ( handler == null )
                    return;

                // Get a list of all subscribers to this event.
                var eventHandlers = handler.GetInvocationList( );
                // Raise the event for each subscriber and check
                // if any of them throws an exception.
                foreach ( var currentHandler in eventHandlers )
                {
                    // Current event subscriber.
                    var currentSubscriber = ( UserDefinedEventHandler )currentHandler;
                    // Raise event for each subscriber and catch exceptions thrown.
                    try
                    {
                        currentSubscriber( sender, e );
                    }
                    catch ( Exception ex )
                    {
                        // Halt Sherlock execution on exception
                        // to prevent the screen from being flooded
                        // with text boxes.
                        Halt( );
                        // I have to use message boxes to alert users of possible exceptions.
                        // This was done because if an exception is thrown in the event handler
                        // the Sherlock COM object swallows the exception.
                        // I used the top level Exception to ensure I catch
                        // any possible exception thrown. If I didn't do this
                        // the exception would go un-noticed.
                        MessageBox.Show(
                            ex.ToString( ), "Encountered Exception",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation );
                    }
                }
            }

            // Raise ReporterUpdate event method.
            private void OnReporterUpdateEvent( object sender, ReporterUpdateEventArgs e )
            {
                var handler = _reporterUpdateHandler;

                if ( handler == null )
                    return;

                // Get a list of all subscribers to this event.
                var eventHandlers = handler.GetInvocationList( );
                // Raise the event for each subscriber and check
                // if any of them throws an exception.
                foreach ( var currentHandler in eventHandlers )
                {
                    // Current event subscriber.
                    var currentSubscriber = ( ReporterUpdateEventHandler )currentHandler;
                    // Raise event for each subscriber and catch exceptions thrown.
                    try
                    {
                        currentSubscriber( sender, e );
                    }
                    catch ( Exception ex )
                    {
                        // Halt Sherlock execution on exception
                        // to prevent the screen from being flooded
                        // with text boxes.
                        Halt( );
                        // I have to use message boxes to alert users of possible exceptions.
                        // This was done because if an exception is thrown in the event handler
                        // the Sherlock COM object swallows the exception.
                        // I used the top level Exception to ensure I catch
                        // any possible exception thrown. If I didn't do this
                        // the exception would go un-noticed.
                        MessageBox.Show(
                            ex.ToString( ), "Encountered Exception",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation );
                    }
                }
            }

            // Raise RunCompleted event method.
            private void OnRunCompleted( object sender, EventArgs e )
            {
                var handler = _runCompletedHandler;

                if ( handler == null )
                    return;

                var eventHandlers = handler.GetInvocationList( );

                foreach ( var currentSubscriber in eventHandlers.Cast<EventHandler>( ) )
                {
                    try
                    {
                        currentSubscriber( sender, e );
                    }
                    catch ( Exception ex )
                    {
                        // Halt Sherlock execution on exception
                        // to prevent the screen from being flooded
                        // with text boxes.
                        Halt( );
                        MessageBox.Show(
                            ex.ToString( ), "Encountered Exception",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation );
                    }
                }
            }

            // Run completed Sherlock event handler.
            // A custom event class was created for this event
            // to make it eaiser to use.
            private void OnDalsaRunCompletedEvent( )
            {
                // Because the Sherlock object was created
                // on a different thread I have to marshall
                // all messages from Sherlock back to the main
                // UI thread.
                _synchronizationContext.Post( callback => OnRunCompleted( this, EventArgs.Empty ), null );
            }

            // Reporter update Sherlock event handler.
            // A custom event class was created for this event
            // to make it eaiser to use.
            private void OnDalsaReporterUpdate( string reporterMessage )
            {
                // Because the Sherlock object was created
                // on a different thread I have to marshall
                // all messages from Sherlock back to the main
                // UI thread.
                _synchronizationContext.Post(
                    callBack => OnReporterUpdateEvent( this, new ReporterUpdateEventArgs( reporterMessage ) ), null );
            }

            // Sherlock on error event.
            // If Sherlock encounters an error which halts execution
            // this event handler is called.
            private void OnDalsaRunErrorEvent( I_EXEC_ERROR sherlockError )
            {
                // Because the Sherlock object was created
                // on a different thread I have to marshall
                // all messages from Sherlock back to the main
                // UI thread.
                _synchronizationContext.Post( callback => OnRunError( this, new RunErrorEventArgs( sherlockError ) ), null );
            }

            // User defined Sherlock event handler.
            // The eventId comes from Sherlock and denotes which 
            // of the user defined events has fired.
            // I pass this eventId to my custom event handler.
            private void OnDalsaUserDefinedEvent( int eventId )
            {
                // Because the Sherlock object was created
                // on a different thread I have to marshall
                // all messages from Sherlock back to the main
                // UI thread.
                _synchronizationContext.Post( callback => OnUserDefinedEvent( this, new UserDefinedEventArgs( eventId ) ), null );
            }

            // Initialize the Sherlock engine.
            private static void InitializeEngine( )
            {
                _sherlock = new Engine( );
                _iReturn = _sherlock.EngInitialize( );
            }

            // This method is called from another thread. It
            // has been sync'd back to the main UI thread so that
            // it can manipulate winforms controls.
            private void AfterInitializationAndInvestigationLoad( )
            {
                if( _iReturn != I_ENG_ERROR.I_OK )
                    throw new ExternalException(
                        string.Concat( "Sherlock failed to load the supplied investigation. Sherlock returned: ", _iReturn.ToString( ) ) );

                // Enable timer used to poll sherlocks 
                // mode of operation.
                _timer.Enabled = true;
                OnInvestigationLoaded( this, EventArgs.Empty );
            }

            // Clean up once finished with object.
            // Because of an issue with the polling timer still
            // firing its elasped event after the Sherlock
            // object was disposed, causing a NullReferenceException,
            // I now stop, dispose, and then wait 500 ms before
            // disposing of the sherlock object.
            private static void CleanUp( )
            {
                if( _timer != null )
                {
                    _timer.Stop( );
                    _timer.Dispose( );
                }
                if( _sherlock != null )
                {
                    _sherlock.EngTerminate( );
                    _sherlock = null;
                }
            }
            #endregion
        }
    }
}

