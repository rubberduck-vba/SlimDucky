using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Vbe.Interop;

namespace SlimDucky
{
    [ComVisible(true)]
    [Guid(Guid)]
    [ProgId(ProgId)]
    [EditorBrowsable(EditorBrowsableState.Never)]
    public class Connect : IDTExtensibility2
    {
        public const string Guid = "3971959D-003F-4D83-8747-3CE434E7553D";
        public const string ProgId = "SlimDucky.Connect";

        private VBE _vbe;
        private AddIn _addIn;
        private App _app;

        private bool _isInitialized;
        private bool _isBeginShutdownInvoked;

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _vbe = Application as VBE;
            _addIn = AddInInst as AddIn;

            _isBeginShutdownInvoked = false;

            switch(ConnectMode)
            {
                case ext_ConnectMode.ext_cm_Startup:
                    // normal execution path - don't initialize just yet, wait for OnStartupComplete to be called by the host.
                    break;
                case ext_ConnectMode.ext_cm_AfterStartup:
                    InitializeAddIn();
                    break;
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            switch(RemoveMode)
            {
                case ext_DisconnectMode.ext_dm_UserClosed:
                    ShutdownAddIn();
                    break;

                case ext_DisconnectMode.ext_dm_HostShutdown:
                    if(_isBeginShutdownInvoked)
                    {
                        // this is the normal case: nothing to do here, we already ran ShutdownAddIn.
                    }
                    else
                    {
                        // some hosts do not call OnBeginShutdown: this mitigates it.
                        ShutdownAddIn();
                    }
                    break;
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            InitializeAddIn();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            _isBeginShutdownInvoked = true;
            ShutdownAddIn();
        }

        private void InitializeAddIn()
        {
            if(_isInitialized)
            {
                // The add-in is already initialized. See:
                // The strange case of the add-in initialized twice
                // http://msmvps.com/blogs/carlosq/archive/2013/02/14/the-strange-case-of-the-add-in-initialized-twice.aspx
                return;
            }

            _app = new App(_vbe, _addIn);
            _app.Startup();

            _isInitialized = true;
        }

        private void ShutdownAddIn()
        {
            _app.Shutdown();
            _app = null;
            _isInitialized = false;
        }
    }
}
