using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace SlimDucky
{
    public class DockableToolwindowPresenter
    {
        private readonly AddIn _addin;
        private readonly Window _window;
        private readonly VBE _vbe;

        private object _userControlObject;

        public DockableToolwindowPresenter(VBE vbe, AddIn addin, IDockableUserControl view)
        {
            _vbe = vbe;
            _addin = addin;
            UserControl = view;
            _window = CreateToolWindow(view);
        }

        public IDockableUserControl UserControl { get; }

        private Window CreateToolWindow(IDockableUserControl view)
        {
            Window toolWindow;
            try
            {
                object control = null;
                toolWindow = _vbe.Windows.CreateToolWindow(_addin, _DockableWindowHost.ProgId, view.Caption, view.ClassId, ref control);
                _userControlObject = control;
            }
            catch(COMException exception)
            {
                Debug.WriteLine(exception);
                throw;
            }

            var userControlHost = (_DockableWindowHost)_userControlObject;
            toolWindow.Visible = true; //window resizing doesn't work without this

            EnsureMinimumWindowSize(toolWindow);

            toolWindow.Visible = true; // here Rubberduck checks window settings to show toolwindow at startup

            userControlHost.AddUserControl(view as UserControl, new IntPtr(_vbe.MainWindow.HWnd));
            return toolWindow;
        }

        private void EnsureMinimumWindowSize(Window window)
        {
            const int defaultWidth = 350;
            const int defaultHeight = 200;

            if(!window.Visible || window.LinkedWindows != null)
            {
                return;
            }

            if(window.Width < defaultWidth)
            {
                window.Width = defaultWidth;
            }

            if(window.Height < defaultHeight)
            {
                window.Height = defaultHeight;
            }
        }

        public virtual void Show() => _window.Visible = true;
        public virtual void Hide() => _window.Visible = false;

        ~DockableToolwindowPresenter()
        {
            // destructor for tracking purposes only - do not suppress unless 
            Debug.WriteLine("DockableToolwindowPresenter finalized.");
        }
    }
}
