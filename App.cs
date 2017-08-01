using Microsoft.Vbe.Interop;

namespace SlimDucky
{
    public class App
    {
        private readonly VBE _vbe;
        private readonly AddIn _adddIn;

        public App(VBE vbe, AddIn addIn)
        {
            _vbe = vbe;
            _adddIn = addIn;

            _winFormsUserControl = new WinFormsControl();
            _wpfUserControl = new WpfHostControl();
        }

        private IDockableUserControl _wpfUserControl;
        private IDockableUserControl _winFormsUserControl;

        public void Startup()
        {
            var winFormsPresenter = new DockableToolwindowPresenter(_vbe, _adddIn, _winFormsUserControl);
            var wpfPresenter = new DockableToolwindowPresenter(_vbe, _adddIn, _wpfUserControl);

            winFormsPresenter.Show();
            wpfPresenter.Show();
        }

        public void Shutdown()
        {
        }
    }
}
