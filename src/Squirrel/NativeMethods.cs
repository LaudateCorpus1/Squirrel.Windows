using System;
using System.Runtime.InteropServices;

namespace Squirrel
{
    static class NativeMethods
    {
        [DllImport("version.dll", SetLastError = true)]
        [return:MarshalAs(UnmanagedType.Bool)] public static extern bool GetFileVersionInfo(
            string lpszFileName, 
            IntPtr dwHandleIgnored,
            int dwLen, 
            [MarshalAs(UnmanagedType.LPArray)] byte[] lpData);

        [DllImport("version.dll", SetLastError = true)]
        public static extern int GetFileVersionInfoSize(
            string lpszFileName,
            IntPtr dwHandleIgnored);

        [DllImport("version.dll")]
        [return:MarshalAs(UnmanagedType.Bool)] public static extern bool VerQueryValue(
            byte[] pBlock, 
            string pSubBlock, 
            out IntPtr pValue, 
            out int len);
    }
}
