using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SlimDucky
{
    [ComVisible(true)]
    [Guid("AA0811DA-4C0E-45FE-A017-994298C675E5")]
    public partial class WinFormsControl : UserControl, IDockableUserControl
    {
        public WinFormsControl()
        {
            InitializeComponent();
        }

        public string ClassId => "AA0811DA-4C0E-45FE-A017-994298C675E5";
        public string Caption => "WinForms ToolWindow";
    }
}
