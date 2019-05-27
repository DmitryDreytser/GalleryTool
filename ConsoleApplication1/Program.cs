using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace GalleryTool
{
    class Program
    {
        static void Main(string[] args)
        {
            string mode = null;
            string directory = null;
            string filename2 = null;

            if (args.Length > 0)
            {
                mode = args[0].ToLower();
                directory = args[1];
                filename2 = args[2];
            }


            ImageList Gallery;
            switch (mode)
            {
                case "-c": //Собрать галерею из картинок
                    Gallery = new ImageList(80, 80);
                    Gallery.Construct(directory);
                    Console.WriteLine("Gallery constructed from {0} images.", Gallery.ImageCount);
                    if(Gallery.Write(filename2))
                        Console.WriteLine("Gallery written to {0}.", filename2);
                    else
                        Console.WriteLine("Gallery write error.");
                    break;
                case "-d": //разобрать на картинки
                    Gallery = new ImageList(directory);
                    Console.WriteLine("Gallery loaded from {0}. {1} images found.", directory, Gallery.ImageCount);
                    Gallery.ExportImages(filename2);
                        Console.WriteLine("Gallery icons exported to {0}.", filename2);
                    break;
                case "-f": //загрузить из bmp после GComp
                    Gallery = ImageList.LoadFromFile(directory, 80, 80);
                    Console.WriteLine("Gallery converted from {0}. {1} images found.", directory, Gallery.ImageCount);
                    if (Gallery.Write(filename2))
                        Console.WriteLine("Gallery written to {0}.", filename2);
                    else
                        Console.WriteLine("Gallery write error.");
                    break;
                case "-r": //загрузить из bmp после GComp
                    Gallery = new ImageList(directory);
                    Console.WriteLine("Gallery converted from {0}. {1} images found.", directory, Gallery.ImageCount);
                    File.Move(directory, filename2);
                    if (Gallery.Write(directory))
                        Console.WriteLine("Gallery written to {0}.", directory);
                    else
                        Console.WriteLine("Gallery write error.");
                    break;
                default:
                    Console.WriteLine("1C PictureGallery tool (c) MadDAD 2017");
                    Console.WriteLine("Usage:\r\n {0} -c path\\PictureGallery path\\outFileName - construct picture gallery from image files", Path.GetFileName(Application.ExecutablePath));
                    Console.WriteLine("{0} -d path\\inFileName path\\out pictures directory - export all icons to bitmap files", Path.GetFileName(Application.ExecutablePath));
                    Console.WriteLine("{0} -f path\\inFileName  path\\outFileName - conver GComp Gallery.bmp file to native ImageList gallery ", Path.GetFileName(Application.ExecutablePath));
                    Console.WriteLine("{0} -r path\\inFileName  path\\BKPFileNAme - repair - conver Imagelist 6.0 to 5.0", Path.GetFileName(Application.ExecutablePath));

                    break;

            }
        }
    }
}


namespace Win32FromForms
{
    delegate IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    class Win32Window
    {
        const UInt32 WS_OVERLAPPEDWINDOW = 0xcf0000;
        const UInt32 WS_VISIBLE = 0x10000000;
        const UInt32 WS_CHILDWINDOW = 0x40000000;

        const UInt32 CS_USEDEFAULT = 0x80000000;
        const UInt32 CS_DBLCLKS = 8;
        const UInt32 CS_VREDRAW = 1;
        const UInt32 CS_HREDRAW = 2;
        const UInt32 COLOR_WINDOW = 5;
        const UInt32 COLOR_BACKGROUND = 1;
        const UInt32 IDC_CROSS = 32515;
        const UInt32 WM_DESTROY = 2;
        const UInt32 WM_PAINT = 0x0f;
        const UInt32 WM_LBUTTONUP = 0x0202;
        const UInt32 WM_LBUTTONDBLCLK = 0x0203;


        [Flags]
        public enum WindowStyles : uint
        {
            WS_OVERLAPPED = 0x00000000,
            WS_POPUP = 0x80000000,
            WS_CHILD = 0x40000000,
            WS_MINIMIZE = 0x20000000,
            WS_VISIBLE = 0x10000000,
            WS_DISABLED = 0x08000000,
            WS_CLIPSIBLINGS = 0x04000000,
            WS_CLIPCHILDREN = 0x02000000,
            WS_MAXIMIZE = 0x01000000,
            WS_BORDER = 0x00800000,
            WS_DLGFRAME = 0x00400000,
            WS_VSCROLL = 0x00200000,
            WS_HSCROLL = 0x00100000,
            WS_SYSMENU = 0x00080000,
            WS_THICKFRAME = 0x00040000,
            WS_GROUP = 0x00020000,
            WS_TABSTOP = 0x00010000,

            WS_MINIMIZEBOX = 0x00020000,
            WS_MAXIMIZEBOX = 0x00010000,

            WS_CAPTION = WS_BORDER | WS_DLGFRAME,
            WS_TILED = WS_OVERLAPPED,
            WS_ICONIC = WS_MINIMIZE,
            WS_SIZEBOX = WS_THICKFRAME,
            WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW,

            WS_OVERLAPPEDWINDOW = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX,
            WS_POPUPWINDOW = WS_POPUP | WS_BORDER | WS_SYSMENU,
            WS_CHILDWINDOW = WS_CHILD,


            BS_RADIOBUTTON = 0x00000004,
            BS_CHECKBOX = 0x00000002,
            BS_AUTOCHECKBOX = 0x00000003,

            RBS_TOOLTIPS = 0x00000100,
            RBS_VARHEIGHT = 0x00000200,
            RBS_BANDBORDERS = 0x00000400,
            RBS_FIXEDORDER = 0x00000800,
            RBS_REGISTERDROP = 0x00001000,
            RBS_AUTOSIZE = 0x00002000,
            RBS_VERTICALGRIPPER = 0x00004000,  // this always has the vertical gripper (default for horizontal mode)
            RBS_DBLCLKTOGGLE = 0x00008000,

            CCS_TOP = 0x00000001,
            CCS_NOMOVEY = 0x00000002,
            CCS_BOTTOM = 0x00000003,
            CCS_NORESIZE = 0x00000004,
            CCS_NOPARENTALIGN = 0x00000008,
            CCS_ADJUSTABLE = 0x00000020,
            CCS_NODIVIDER = 0x00000040,
            CCS_VERT = 0x00000080,
            CCS_LEFT = (CCS_VERT | CCS_TOP),
            CCS_RIGHT = (CCS_VERT | CCS_BOTTOM),
            CCS_NOMOVEX = (CCS_VERT | CCS_NOMOVEY),

             TBSTYLE_BUTTON = 0x0000,  // obsolete; use BTNS_BUTTON instead,
            TBSTYLE_SEP = 0x0001, // obsolete; use BTNS_SEP instead,
            TBSTYLE_CHECK = 0x0002,  // obsolete; use BTNS_CHECK instead,
            TBSTYLE_GROUP = 0x0004,  // obsolete; use BTNS_GROUP instead,
            TBSTYLE_CHECKGROUP = (TBSTYLE_GROUP | TBSTYLE_CHECK),     // obsolete; use BTNS_CHECKGROUP instead,
            TBSTYLE_DROPDOWN = 0x0008,  // obsolete; use BTNS_DROPDOWN instead,
            TBSTYLE_AUTOSIZE = 0x0010,  // obsolete; use BTNS_AUTOSIZE instead,
            TBSTYLE_NOPREFIX = 0x0020,  // obsolete; use BTNS_NOPREFIX instead,
            TBSTYLE_TOOLTIPS = 0x0100,
            TBSTYLE_WRAPABLE = 0x0200,
            TBSTYLE_ALTDRAG = 0x0400,
            TBSTYLE_FLAT = 0x0800,
            TBSTYLE_LIST = 0x1000,
            TBSTYLE_CUSTOMERASE = 0x2000,
            TBSTYLE_REGISTERDROP = 0x4000,
            TBSTYLE_TRANSPARENT = 0x8000

            /*
             #define BS_PUSHBUTTON       0x00000000L
             #define BS_DEFPUSHBUTTON    0x00000001L
             #define BS_CHECKBOX         0x00000002L
             #define BS_AUTOCHECKBOX     0x00000003L
             #define 
             #define BS_3STATE           0x00000005L
             #define BS_AUTO3STATE       0x00000006L
             #define BS_GROUPBOX         0x00000007L
             #define BS_USERBUTTON       0x00000008L
             #define BS_AUTORADIOBUTTON  0x00000009L
             #define BS_PUSHBOX          0x0000000AL
             #define BS_OWNERDRAW        0x0000000BL
             #define BS_TYPEMASK         0x0000000FL
             #define BS_LEFTTEXT         0x00000020L
             #if(WINVER >= 0x0400)
             #define BS_TEXT             0x00000000L
             #define BS_ICON             0x00000040L
             #define BS_BITMAP           0x00000080L
             #define BS_LEFT             0x00000100L
             #define BS_RIGHT            0x00000200L
             #define BS_CENTER           0x00000300L
             #define BS_TOP              0x00000400L
             #define BS_BOTTOM           0x00000800L
             #define BS_VCENTER          0x00000C00L
             #define BS_PUSHLIKE         0x00001000L
             #define BS_MULTILINE        0x00002000L
             #define BS_NOTIFY           0x00004000L
             #define BS_FLAT             0x00008000L
             #define BS_RIGHTBUTTON      BS_LEFTTEXT
             */
        }
        [Flags]
        public enum WindowStylesEx : uint
        {
            //Extended Window Styles

            WS_EX_DLGMODALFRAME = 0x00000001,
            WS_EX_NOPARENTNOTIFY = 0x00000004,
            WS_EX_TOPMOST = 0x00000008,
            WS_EX_ACCEPTFILES = 0x00000010,
            WS_EX_TRANSPARENT = 0x00000020,

            //#if(WINVER >= 0x0400)

            WS_EX_MDICHILD = 0x00000040,
            WS_EX_TOOLWINDOW = 0x00000080,
            WS_EX_WINDOWEDGE = 0x00000100,
            WS_EX_CLIENTEDGE = 0x00000200,
            WS_EX_CONTEXTHELP = 0x00000400,

            WS_EX_RIGHT = 0x00001000,
            WS_EX_LEFT = 0x00000000,
            WS_EX_RTLREADING = 0x00002000,
            WS_EX_LTRREADING = 0x00000000,
            WS_EX_LEFTSCROLLBAR = 0x00004000,
            WS_EX_RIGHTSCROLLBAR = 0x00000000,

            WS_EX_CONTROLPARENT = 0x00010000,
            WS_EX_STATICEDGE = 0x00020000,
            WS_EX_APPWINDOW = 0x00040000,

            WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE),
            WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST),
            WS_EX_LAYERED = 0x00080000,
            WS_EX_NOINHERITLAYOUT = 0x00100000, // Disable inheritence of mirroring by children
            WS_EX_LAYOUTRTL = 0x00400000, // Right to left mirroring
            WS_EX_COMPOSITED = 0x02000000,
            WS_EX_NOACTIVATE = 0x08000000
        }


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        struct WNDCLASSEX
        {
            [MarshalAs(UnmanagedType.U4)]
            public int cbSize;
            [MarshalAs(UnmanagedType.U4)]
            public int style;
            public IntPtr lpfnWndProc;
            public int cbClsExtra;
            public int cbWndExtra;
            public IntPtr hInstance;
            public IntPtr hIcon;
            public IntPtr hCursor;
            public IntPtr hbrBackground;
            public string lpszMenuName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpszClassName;
            public IntPtr hIconSm;
        }

        enum TBMessages
        {
            WM_USER = 1024,
            TB_ADDBUTTONSA = (WM_USER + 20),
            TB_INSERTBUTTONA = (WM_USER + 21),
            TB_DELETEBUTTON = (WM_USER + 22),
            TB_GETBUTTON = (WM_USER + 23),
            TB_BUTTONCOUNT = (WM_USER + 24),
            TB_COMMANDTOINDEX = (WM_USER + 25),
            TB_SAVERESTOREA = (WM_USER + 26),
            TB_SAVERESTOREW = (WM_USER + 76),
            TB_CUSTOMIZE = (WM_USER + 27),
            TB_ADDSTRINGA = (WM_USER + 28),
            TB_ADDSTRINGW = (WM_USER + 77),
            TB_GETITEMRECT = (WM_USER + 29),
            TB_BUTTONSTRUCTSIZE = (WM_USER + 30),
            TB_SETBUTTONSIZE = (WM_USER + 31),
            TB_SETBITMAPSIZE = (WM_USER + 32),
            TB_AUTOSIZE = (WM_USER + 33),
            TB_GETTOOLTIPS = (WM_USER + 35),
            TB_SETTOOLTIPS = (WM_USER + 36),
            TB_SETPARENT = (WM_USER + 37),
            TB_SETROWS = (WM_USER + 39),
            TB_GETROWS = (WM_USER + 40),
            TB_SETCMDID = (WM_USER + 42),
            TB_CHANGEBITMAP = (WM_USER + 43),
            TB_GETBITMAP = (WM_USER + 44),
            TB_GETBUTTONTEXTA = (WM_USER + 45),
            TB_GETBUTTONTEXTW = (WM_USER + 75),
            TB_REPLACEBITMAP = (WM_USER + 46),
            TB_SETINDENT = (WM_USER + 47),
            TB_SETIMAGELIST = (WM_USER + 48),
            TB_GETIMAGELIST = (WM_USER + 49),
            TB_LOADIMAGES = (WM_USER + 50),
            TB_GETRECT = (WM_USER + 51), // wParam is the Cmd instead of index,
            TB_SETHOTIMAGELIST = (WM_USER + 52),
            TB_GETHOTIMAGELIST = (WM_USER + 53),
            TB_SETDISABLEDIMAGELIST = (WM_USER + 54),
            TB_GETDISABLEDIMAGELIST = (WM_USER + 55),
            TB_SETSTYLE = (WM_USER + 56),
            TB_GETSTYLE = (WM_USER + 57),
            TB_GETBUTTONSIZE = (WM_USER + 58),
            TB_SETBUTTONWIDTH = (WM_USER + 59),
            TB_SETMAXTEXTROWS = (WM_USER + 60),
            TB_GETTEXTROWS = (WM_USER + 61),
            TB_ADDBUTTONSW = (WM_USER + 68),
            TB_SETEXTENDEDSTYLE  =   (WM_USER + 84)
        }


        public struct NativeMacros
        {
            /// <summary>
            /// C++ макрос для GET_X_LPARAM.
            /// </summary>
            public static int GET_X_LPARAM(int x)
            {
                return x & 0xffff;
            }
            /// <summary>
            /// C++ макрос для GET_Y_LPARAM.
            /// </summary>
            public static int GET_Y_LPARAM(int x)
            {
                return (x >> 16) & 0xffff;
            }
            /// <summary>
            /// C++ макрос для MAKELONG.
            /// </summary>
            public static int MAKELONG(int x, int y)
            {
                return (x & 0xffff) | ((y & 0xffff) << 16);
            }
        }

        [StructLayout(LayoutKind.Explicit, CharSet = CharSet.Auto, Pack = 1, Size = 20)]
        struct TBBUTTON
        {
            [FieldOffset(0)]
            public int iBitmap;

            [FieldOffset(4)]
            public int idCommand;

            [FieldOffset(8)]
            public byte fsState;

            [FieldOffset(9)]
            public byte fsStyle;

            [FieldOffset(10)]
            //[MarshalAs(UnmanagedType.I4, SizeConst = 2)]
            public int bReserved;

            [FieldOffset(12)]
            public IntPtr dwData;

            [FieldOffset(16)]
            [MarshalAs(UnmanagedType.LPWStr)]
            public string iString;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }

            public POINT(System.Drawing.Point pt) : this(pt.X, pt.Y) { }

            public static implicit operator System.Drawing.Point(POINT p)
            {
                return new System.Drawing.Point(p.X, p.Y);
            }

            public static implicit operator POINT(System.Drawing.Point p)
            {
                return new POINT(p.X, p.Y);
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public IntPtr hwnd;
            public UInt32 message;
            public IntPtr wParam;
            public IntPtr lParam;
            public UInt32 time;
            public POINT pt;
        }

        private WndProc delegWndProc = myWndProc;

        [DllImport("user32.dll")]
        static extern bool UpdateWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        static extern bool DestroyWindow(IntPtr hWnd);


        [DllImport("user32.dll", SetLastError = true, EntryPoint = "CreateWindowEx")]
        public static extern IntPtr CreateWindowEx(
           WindowStylesEx dwExStyle,
           IntPtr regResult,
           string lpWindowName,
           WindowStyles dwStyle,
           int x,
           int y,
           int nWidth,
           int nHeight,
           IntPtr hWndParent,
           IntPtr hMenu,
           IntPtr hInstance,
           IntPtr lpParam);

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "CreateWindowEx")]
        public static extern IntPtr CreateWindowEx(
           WindowStylesEx dwExStyle,
           [MarshalAs(UnmanagedType.LPStr)]
           string lpClassName,
           [MarshalAs(UnmanagedType.LPStr)]
           string lpWindowName,
           WindowStyles dwStyle,
           int x,
           int y,
           int nWidth,
           int nHeight,
           IntPtr hWndParent,
           IntPtr hMenu,
           IntPtr hInstance,
           IntPtr lpParam);

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "RegisterClassEx")]
        static extern IntPtr RegisterClassEx([In] ref WNDCLASSEX lpWndClass);

        [DllImport("kernel32.dll")]
        static extern uint GetLastError();

        [DllImport("user32.dll")]
        static extern IntPtr DefWindowProc(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern int SendMessage(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern void PostQuitMessage(int nExitCode);

        [DllImport("user32.dll", SetLastError = true)]
        static extern sbyte GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin,
           uint wMsgFilterMax);

        [DllImport("user32.dll")]
        static extern IntPtr LoadCursor(IntPtr hInstance, int lpCursorName);

        [DllImport("user32.dll")]
        static extern bool TranslateMessage([In] ref MSG lpMsg);

        [DllImport("user32.dll")]
        static extern IntPtr DispatchMessage([In] ref MSG lpmsg);

        [DllImport("comctl32.dll", SetLastError = true)]
        static extern IntPtr ImageList_Create(int cx, int cy, uint flags, int cInitial, int cGrow);

        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern IntPtr ImageList_Read(IStream pstm);

        public class STATSTG
        {

            [MarshalAs(UnmanagedType.LPWStr)]
            public string pwcsName = null;

            public int type;
            [MarshalAs(UnmanagedType.I8)]
            public long cbSize;
            [MarshalAs(UnmanagedType.I8)]
            public long mtime = 0;
            [MarshalAs(UnmanagedType.I8)]
            public long ctime = 0;
            [MarshalAs(UnmanagedType.I8)]
            public long atime = 0;
            [MarshalAs(UnmanagedType.I4)]
            public int grfMode = 0;
            [MarshalAs(UnmanagedType.I4)]
            public int grfLocksSupported;

            public int clsid_data1 = 0;
            [MarshalAs(UnmanagedType.I2)]
            public short clsid_data2 = 0;
            [MarshalAs(UnmanagedType.I2)]
            public short clsid_data3 = 0;
            [MarshalAs(UnmanagedType.U1)]
            public byte clsid_b0 = 0;
            [MarshalAs(UnmanagedType.U1)]
            public byte clsid_b1 = 0;
            [MarshalAs(UnmanagedType.U1)]
            public byte clsid_b2 = 0;
            [MarshalAs(UnmanagedType.U1)]
            public byte clsid_b3 = 0;
            [MarshalAs(UnmanagedType.U1)]
            public byte clsid_b4 = 0;
            [MarshalAs(UnmanagedType.U1)]
            public byte clsid_b5 = 0;
            [MarshalAs(UnmanagedType.U1)]
            public byte clsid_b6 = 0;
            [MarshalAs(UnmanagedType.U1)]
            public byte clsid_b7 = 0;
            [MarshalAs(UnmanagedType.I4)]
            public int grfStateBits = 0;
            [MarshalAs(UnmanagedType.I4)]
            public int reserved = 0;
        }

        public class ComStreamFromDataStream : IStream
        {
            protected Stream dataStream;

            // to support seeking ahead of the stream length...
            private long virtualPosition = -1;

            public ComStreamFromDataStream(Stream dataStream)
            {
                if (dataStream == null) throw new ArgumentNullException("dataStream");
                this.dataStream = dataStream;
            }

            private void ActualizeVirtualPosition()
            {
                if (virtualPosition == -1) return;

                if (virtualPosition > dataStream.Length)
                    dataStream.SetLength(virtualPosition);

                dataStream.Position = virtualPosition;

                virtualPosition = -1;
            }

            public IStream Clone()
            {
                NotImplemented();
                return null;
            }

            public void Commit(int grfCommitFlags)
            {
                dataStream.Flush();
                // Extend the length of the file if needed.
                ActualizeVirtualPosition();
            }

            public long CopyTo(IStream pstm, long cb, long[] pcbRead)
            {
                int bufsize = 4096; // one page
                IntPtr buffer = Marshal.AllocHGlobal(bufsize);
                if (buffer == IntPtr.Zero) throw new OutOfMemoryException();
                long written = 0;
                try
                {
                    while (written < cb)
                    {
                        int toRead = bufsize;
                        if (written + toRead > cb) toRead = (int)(cb - written);
                        int read = Read(buffer, toRead);
                        if (read == 0) break;
                        if (pstm.Write(buffer, read) != read)
                        {
                            throw EFail("Wrote an incorrect number of bytes");
                        }
                        written += read;
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
                if (pcbRead != null && pcbRead.Length > 0)
                {
                    pcbRead[0] = written;
                }

                return written;
            }

            public Stream GetDataStream()
            {
                return dataStream;
            }

            public void LockRegion(long libOffset, long cb, int dwLockType)
            {
            }

            protected static ExternalException EFail(string msg)
            {
                ExternalException e = new ExternalException(msg, E_FAIL);
                throw e;
            }

            protected static void NotImplemented()
            {
                ExternalException e = new ExternalException("Не реализовано", E_NOTIMPL);
                throw e;
            }

            public int Read(IntPtr buf, /* cpr: int offset,*/  int length)
            {
                //        System.Text.Out.WriteLine("IStream::Read(" + length + ")");
                byte[] buffer = new byte[length];
                int count = Read(buffer, length);
                Marshal.Copy(buffer, 0, buf, count);
                return count;
            }

            public int Read(byte[] buffer, /* cpr: int offset,*/  int length)
            {
                ActualizeVirtualPosition();
                return dataStream.Read(buffer, 0, length);
            }

            public void Revert()
            {
                NotImplemented();
            }

            public const int E_NOTIMPL = unchecked((int)0x80004001),
                E_OUTOFMEMORY = unchecked((int)0x8007000E),
                E_INVALIDARG = unchecked((int)0x80070057),
                E_NOINTERFACE = unchecked((int)0x80004002),
                E_POINTER = unchecked((int)0x80004003),
                E_FAIL = unchecked((int)0x80004005),
                STREAM_SEEK_SET = 0x0,
                STREAM_SEEK_CUR = 0x1,
                STREAM_SEEK_END = 0x2;

            public long Seek(long offset, int origin)
            {
                // Console.WriteLine("IStream::Seek("+ offset + ", " + origin + ")");
                long pos = virtualPosition;
                if (virtualPosition == -1)
                {
                    pos = dataStream.Position;
                }
                long len = dataStream.Length;
                switch (origin)
                {
                    case STREAM_SEEK_SET:
                        if (offset <= len)
                        {
                            dataStream.Position = offset;
                            virtualPosition = -1;
                        }
                        else
                        {
                            virtualPosition = offset;
                        }
                        break;
                    case STREAM_SEEK_END:
                        if (offset <= 0)
                        {
                            dataStream.Position = len + offset;
                            virtualPosition = -1;
                        }
                        else
                        {
                            virtualPosition = len + offset;
                        }
                        break;
                    case STREAM_SEEK_CUR:
                        if (offset + pos <= len)
                        {
                            dataStream.Position = pos + offset;
                            virtualPosition = -1;
                        }
                        else
                        {
                            virtualPosition = offset + pos;
                        }
                        break;
                }
                if (virtualPosition != -1)
                {
                    return virtualPosition;
                }
                else
                {
                    return dataStream.Position;
                }
            }

            public void SetSize(long value)
            {
                dataStream.SetLength(value);
            }

            public void Stat(STATSTG pstatstg, int grfStatFlag)
            {
                pstatstg.type = 2; // STGTY_STREAM
                pstatstg.cbSize = dataStream.Length;
                pstatstg.grfLocksSupported = 2; //LOCK_EXCLUSIVE
            }

            public void UnlockRegion(long libOffset, long cb, int dwLockType)
            {
            }

            public int Write(IntPtr buf, /* cpr: int offset,*/ int length)
            {
                byte[] buffer = new byte[length];
                Marshal.Copy(buf, buffer, 0, length);
                return Write(buffer, length);
            }

            public int Write(byte[] buffer, /* cpr: int offset,*/ int length)
            {
                ActualizeVirtualPosition();
                dataStream.Write(buffer, 0, length);
                return length;
            }
        }

        [ComImport(), Guid("0000000C-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IStream
        {

            int Read(

                    IntPtr buf,

                    int len);


            int Write(

                    IntPtr buf,

                    int len);

            [return: MarshalAs(UnmanagedType.I8)]
            long Seek(
                    [In, MarshalAs(UnmanagedType.I8)]
                 long dlibMove,

                     int dwOrigin);


            void SetSize(
                   [In, MarshalAs(UnmanagedType.I8)]
                 long libNewSize);

            [return: MarshalAs(UnmanagedType.I8)]
            long CopyTo(
                    [In, MarshalAs(UnmanagedType.Interface)]
                    IStream pstm,
                    [In, MarshalAs(UnmanagedType.I8)]
                 long cb,
                    [Out, MarshalAs(UnmanagedType.LPArray)]
                 long[] pcbRead);


            void Commit(

                    int grfCommitFlags);


            void Revert();


            void LockRegion(
                   [In, MarshalAs(UnmanagedType.I8)]
                 long libOffset,
                   [In, MarshalAs(UnmanagedType.I8)]
                 long cb,

                    int dwLockType);


            void UnlockRegion(
                   [In, MarshalAs(UnmanagedType.I8)]
                 long libOffset,
                   [In, MarshalAs(UnmanagedType.I8)]
                 long cb,

                    int dwLockType);


            void Stat(
                    [Out]
                 
                    STATSTG pStatstg,
                    int grfStatFlag);

            [return: MarshalAs(UnmanagedType.Interface)]
            IStream Clone();
        }



        const int ImageListID = 0;
        const int numButtons = 3;
        const int bitmapSize = 16;

        internal bool create()
        {
            WNDCLASSEX wind_class = new WNDCLASSEX();
            wind_class.cbSize = Marshal.SizeOf(typeof(WNDCLASSEX));
            wind_class.style = (int)(CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS); //Doubleclicks are active
            wind_class.hbrBackground = (IntPtr)COLOR_BACKGROUND + 1; //Black background, +1 is necessary
            wind_class.cbClsExtra = 0;
            wind_class.cbWndExtra = 0;
            wind_class.hInstance = Marshal.GetHINSTANCE(this.GetType().Module); ;// alternative: Process.GetCurrentProcess().Handle;
            wind_class.hIcon = IntPtr.Zero;
            wind_class.hCursor = LoadCursor(IntPtr.Zero, (int)IDC_CROSS);// Crosshair cursor;
            wind_class.lpszMenuName = null;
            wind_class.lpszClassName = "ToolbarWindow32"; //  "Afx:400000:8:10005:10:0";//
            wind_class.lpfnWndProc = Marshal.GetFunctionPointerForDelegate(delegWndProc);
            wind_class.hIconSm = IntPtr.Zero;

            IntPtr regResult = RegisterClassEx(ref wind_class);

            if (regResult == IntPtr.Zero)
            {
                uint error = GetLastError();
                return false;
            }
            string wndClass = wind_class.lpszClassName;

            //The next line did NOT work with me! When searching the web, the reason seems to be unclear! 
            //It resulted in a zero hWnd, but GetLastError resulted in zero (i.e. no error) as well !!??)
            //IntPtr hWnd = CreateWindowEx(0, wind_class.lpszClassName, "MyWnd", WS_OVERLAPPEDWINDOW | WS_VISIBLE, 0, 0, 30, 40, IntPtr.Zero, IntPtr.Zero, wind_class.hInstance, IntPtr.Zero);

            //This version worked and resulted in a non-zero hWnd
            IntPtr ParentWnd = new IntPtr(0x000C054E);
            //IntPtr ParentWnd = IntPtr.Zero;
            uint DwStyle = 0x54002834; //WS_OVERLAPPEDWINDOW | WS_VISIBLE | WS_CHILDWINDOW
            WindowStyles Style = WindowStyles.WS_CHILD | WindowStyles.TBSTYLE_SEP| WindowStyles.TBSTYLE_FLAT;
            //IntPtr hWnd = CreateWindowEx(0, wndClass, "Hello Win32", DwStyle , 547, 0, 100, 30, ParentWnd, IntPtr.Zero, wind_class.hInstance, IntPtr.Zero);
            IntPtr hWnd = CreateWindowEx(WindowStylesEx.WS_EX_TOOLWINDOW, wndClass, "Hello Win32", Style, 547, 30, 0, 0, ParentWnd, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);

            if (hWnd == IntPtr.Zero)
            {
                uint error = GetLastError();
                return false;
            }

            int Res;
            IntPtr g_hImageList = ImageList_Create(bitmapSize, bitmapSize,   // Dimensions of individual bitmaps.
                                    0x00000010 | 0x00000001,   // Ensures transparent background.
                                    numButtons, 0);

            Res = SendMessage(hWnd, (uint)TBMessages.TB_SETIMAGELIST, (IntPtr)ImageListID, g_hImageList);
            Res = SendMessage(hWnd, (uint)TBMessages.TB_LOADIMAGES, IntPtr.Zero, wind_class.hInstance);

            TBBUTTON[] tbButtons = new TBBUTTON[]
            {
                new TBBUTTON { iBitmap= NativeMacros.MAKELONG(1,  ImageListID), idCommand= 1, fsState =  0x04, fsStyle = 0x10, bReserved =0, dwData= IntPtr.Zero, iString = "New" },
                new TBBUTTON { iBitmap= NativeMacros.MAKELONG(2,  ImageListID), idCommand= 2, fsState =  0x04, fsStyle = 0x10, bReserved = 0, dwData= IntPtr.Zero, iString = "New" },
                new TBBUTTON { iBitmap= NativeMacros.MAKELONG(3,  ImageListID), idCommand= 3, fsState =  0x04, fsStyle = 0x10, bReserved = 0, dwData= IntPtr.Zero, iString = "New" },
            };

            IntPtr lpButtons = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(TBBUTTON)) * tbButtons.Length);
            Res = SendMessage(hWnd, (uint)TBMessages.TB_BUTTONSTRUCTSIZE, new IntPtr(Marshal.SizeOf(typeof(TBBUTTON))), IntPtr.Zero);
            //Res = SendMessage(hWnd, (uint)TBMessages.TB_GETBUTTONSIZE, IntPtr.Zero, IntPtr.Zero);


            try
            {
                for (int i = 0; i < tbButtons.Length; i++)
                    Marshal.StructureToPtr(tbButtons[i], lpButtons + i * Marshal.SizeOf(typeof(TBBUTTON)), false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //SendMessage(hWnd, (uint)TBMessages.TB_SETEXTENDEDSTYLE, IntPtr.Zero, new IntPtr(0x00000008 | 0x00000001 | 0x00000080 | 0x00000010));

            Res = SendMessage(hWnd, (uint)TBMessages.TB_ADDBUTTONSW, new IntPtr(numButtons), lpButtons);
            Res = SendMessage(hWnd, (uint)TBMessages.TB_SETBUTTONSIZE, new IntPtr(NativeMacros.MAKELONG(21, 21)), new IntPtr(NativeMacros.MAKELONG(21, 21)));

            ShowWindow(hWnd, 1);
            UpdateWindow(hWnd);

            MSG msg;
            while (GetMessage(out msg, IntPtr.Zero, 0, 0) != 0)
            {
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }
            Marshal.FreeHGlobal(lpButtons);

            return true;
        }

        private static IntPtr myWndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
        {
            switch (msg)
            {
                // All GUI painting must be done here
                case WM_PAINT:
                    break;

                case WM_LBUTTONDBLCLK:
                    MessageBox.Show("Doubleclick");
                    break;

                case WM_DESTROY:
                    DestroyWindow(hWnd);

                    //If you want to shutdown the application, call the next function instead of DestroyWindow
                    PostQuitMessage(0);
                    break;

                default:
                    break;
            }
            return DefWindowProc(hWnd, msg, wParam, lParam);
        }
    }
}
