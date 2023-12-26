using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace SampleShellComCSharp
{
    class Program
    {
        /// <summary>
        /// HRESULTの判別が必要なメソッドを使用しないため、PreserveSigを使わずエラーを例外に変換する
        /// </summary>
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
        public interface IShellItem
        {
            void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        static extern void SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string pszPath, IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);

        static void Main()
        {
            string inputPath = @"C:\Users\user\Desktop";
            //string inputPath = @"D:\Data";

            try
            {

                Guid clsidShellItem = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");
                SHCreateItemFromParsingName(inputPath, IntPtr.Zero, clsidShellItem, out var shellItem);
                try
                {
                    // GetDisplayNameを呼び出し
                    shellItem.GetDisplayName(SIGDN.DESKTOPABSOLUTEEDITING, out var pathPtr);
                    string path = Marshal.PtrToStringUni(pathPtr);

                    Console.WriteLine($"Name retrieved successfully. Path: {path}");
                }
                finally
                {
                    Marshal.ReleaseComObject(shellItem);
                }
            }
            catch (COMException e)
            {
                Console.WriteLine($"COM Exception: {e.Message}");
            }
        }

        // SIGDN (Shell Item Get Display Name) 定数
        public enum SIGDN : uint
        {
            NORMALDISPLAY = 0,
            PARENTRELATIVEPARSING = 0x80018001,
            PARENTRELATIVEFORADDRESSBAR = 0x8001c001,
            DESKTOPABSOLUTEPARSING = 0x80028000,
            PARENTRELATIVEEDITING = 0x80031001,
            DESKTOPABSOLUTEEDITING = 0x8004c000,
            FILESYSPATH = 0x80058000,
            URL = 0x80068000
        }
    }
}
