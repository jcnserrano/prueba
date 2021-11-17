using System;
using System.Runtime.InteropServices;

namespace Correos.SimuladorOfertas.Common
{
    internal static class NativeMethods
    {
        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi)]
        internal static extern int MAPISendMail(IntPtr sess, IntPtr hwnd, MapiMessage message, int flg, int rsv);
    }
}
