using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace SlimDucky
{
    public class App
    {
        private readonly VBE _vbe;

        public App(VBE vbe)
        {
            _vbe = vbe;
        }

        public void Startup()
        {
            
        }

        public void Shutdown()
        {
            
        }
    }
}
