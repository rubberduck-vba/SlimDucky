using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SlimDucky
{
    [ComVisible(true)]
    [Guid("23D58F4F-FEE2-4D63-A5AE-EBBD360F26ED")]
    public partial class WpfHostControl : UserControl, IDockableUserControl
    {
        public WpfHostControl()
        {
            InitializeComponent();
        }

        public string ClassId => "23D58F4F-FEE2-4D63-A5AE-EBBD360F26ED";
        public string Caption => "WPF ToolWindow";
    }
}
