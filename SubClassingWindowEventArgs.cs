using System;

namespace SlimDucky
{
    public class SubClassingWindowEventArgs : EventArgs
    {
        private readonly IntPtr _lparam;

        public IntPtr LParam
        {
            get { return _lparam; }
        }

        public bool Closing { get; set; }

        public SubClassingWindowEventArgs(IntPtr lparam)
        {
            _lparam = lparam;
        }
    }
}