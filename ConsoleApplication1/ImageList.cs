using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GalleryTool
{
    class ImageList
    {
        private IntPtr Handle;
        public int ImageCount = 0;
        public string filename;

        [StructLayout(LayoutKind.Sequential)]
        public class IMAGEINFO
        {
            public IntPtr hbmImage = IntPtr.Zero;
            public IntPtr hbmMask = IntPtr.Zero;
            public int Unused1 = 0;
            public int Unused2 = 0;
            public int rcImage_left = 0;
            public int rcImage_top = 0;
            public int rcImage_right = 0;
            public int rcImage_bottom = 0;
        }
        [StructLayout(LayoutKind.Sequential)]
        // This is not our convention for managed resources.
        public class BITMAP
        {
            public int bmType = 0;
            public int bmWidth = 0;
            public int bmHeight = 0;
            public int bmWidthBytes = 0;
            public short bmPlanes = 0;
            public short bmBitsPixel = 0;
            public IntPtr bmBits = IntPtr.Zero;
        }
        [StructLayout(LayoutKind.Sequential)]
        public class ICONINFO
        {
            public int fIcon = 0;
            public int xHotspot = 0;
            public int yHotspot = 0;
            public IntPtr hbmMask = IntPtr.Zero;
            public IntPtr hbmColor = IntPtr.Zero;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        struct COLORREF
        {
            public byte R;
            public byte G;
            public byte B;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct GalleryFAT
        {
            public uint Header;
            public int Length;
            public int[] Indexes;
            public short LastIndex;
            public int IdStringLength;
            public string IdString;

            public void Write(Stream ms)
            {
                //using 
                BinaryWriter bw = new BinaryWriter(ms);
                {
                    bw.Write(Header);
                    bw.Write(Length);
                    for (int i = 0; i < Length; i++)
                        bw.Write(Indexes[i]);
                    bw.Write(LastIndex);

                    IdString = IdString.Replace("{}", "");
                    IdString = IdString.TrimEnd(new char[] { ' ' });
                    IdString = IdString.Replace("{{", "{");
                    IdString = IdString.Replace(' ', ',');
                    IdString = "{" + IdString + "}";
                    IdStringLength = IdString.Length;

                    if (IdString.Length > 0xFF)
                    {
                        bw.Write(new byte[] { 0x00, 0x00, 0xFF });
                        bw.Write(((short)IdStringLength));
                    }

                    if (IdString.Length <= 0xFF)
                    {
                        bw.Write(new byte[]{ 0x00, 0x0F, (byte)IdStringLength} );
                    }

                    bw.Write(Encoding.GetEncoding(1251).GetBytes(IdString));
                    
                }

            }
        }

        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern void InitCommonControls();

        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern bool ImageList_GetImageInfo(HandleRef himl, int i, IMAGEINFO pImageInfo);

        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern bool ImageList_GetIconSize(HandleRef himl, out int x, out int y);

        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern int ImageList_GetImageCount(HandleRef himl);

        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern IntPtr ImageList_Create(int cx, int cy, uint flags, int cInitial, int cGrow);

        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern IntPtr ImageList_Read(IStream pstm);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetObject(HandleRef hObject, int nSize, [In, Out] BITMAP bm);

        [DllImport("Gdi32.dll", SetLastError = true, ExactSpelling = true, EntryPoint = "DeleteObject", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        internal static extern bool DeleteObject(HandleRef hObject);

        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern int ImageList_WriteEx(HandleRef himl, int dwFlags, IStream pstm);

        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern bool ImageList_Write(HandleRef himl, IStream pstm);

        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern bool ImageList_Draw(HandleRef himl, int index, HandleRef hdcDst, int x, int y, int fStyle);
        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern int ImageList_Add(HandleRef himl, HandleRef hBitmap, HandleRef hMask);

        const uint LR_LOADFROMFILE = 16;
        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern IntPtr ImageList_LoadImage(IntPtr hres, [MarshalAs(UnmanagedType.LPStr)] string lpbmp, int cx, int cGrow, IntPtr crMask, uint uType, uint uFlags = LR_LOADFROMFILE);

        
        [DllImport("comctl32.dll", SetLastError = true)]
        public static extern IntPtr ImageList_LoadImage(HandleRef hres, IntPtr lpbmp, int cx, int cGrow, IntPtr crMask, uint uType = 0, uint uFlags = LR_LOADFROMFILE);

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

        public void GetInfo()
        {
            ImageCount = ImageList_GetImageCount(new HandleRef(this, this.Handle));

            IMAGEINFO imageInfo = new IMAGEINFO();
            if (ImageList_GetImageInfo(new HandleRef(this, this.Handle), 0, imageInfo))
            {
                BITMAP bmp = new BITMAP();
                GetObject(new HandleRef(null, imageInfo.hbmImage), Marshal.SizeOf(bmp), bmp);
            }
        }



        Bitmap[] images = null;
        GalleryFAT FAT;

        public void GetImages()
        {
            ImageCount = ImageList_GetImageCount(new HandleRef(this, this.Handle));
            images = new Bitmap[ImageCount];
            int x;
            int y;
            ImageList_GetIconSize(new HandleRef(this, this.Handle), out x, out y);

            for (int i = 0; i < ImageCount; i++)
            {
                IMAGEINFO imageInfo = new IMAGEINFO();
                if (ImageList_GetImageInfo(new HandleRef(this, this.Handle), i, imageInfo))
                {
                    Bitmap bm = new Bitmap(x, y);
                    Graphics g = Graphics.FromImage(bm);
                    IntPtr dc = g.GetHdc();
                    try
                    {
                        ImageList_Draw(new HandleRef(this, this.Handle), i, new HandleRef(g, dc), 0, 0, 0);
                    }
                    finally
                    {
                        g.ReleaseHdcInternal(dc);
                    }
                    g.Flush();
                    
                    images[i] = bm;
                }
            }
        }

        const int ILP_DOWNLEVEL = 1;

        public bool Write(string filename)
        {
            if(this.Handle != IntPtr.Zero)
            {

                FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.ReadWrite);
                ComStreamFromDataStream cs = new ComStreamFromDataStream(fs);
                ImageList_WriteEx(new HandleRef(this, this.Handle), ILP_DOWNLEVEL, cs);

                
                byte[] header = new byte[28];
                byte[] Bitmap = new byte[fs.Length - 28];
                

                fs.Seek(0, SeekOrigin.Begin);
                fs.Read(header, 0, 28);
                fs.Read(Bitmap, 0, Bitmap.Length);
                FAT.Write(fs);
                byte[] Tail = new byte[fs.Length - Bitmap.Length - 28];
                fs.Seek(header.Length + Bitmap.Length, SeekOrigin.Begin);
                fs.Read(Tail, 0, Tail.Length);
                fs.Flush();
                fs.Dispose();

                File.WriteAllBytes("Gallery.head", header);
                File.WriteAllBytes("Gallery.bmp", Bitmap);
                File.WriteAllBytes("Gallery.tail", Tail);
                return true;
            }
            return false;
        }

        public void ExportImages(string directory = null)
        {
            if (!string.IsNullOrEmpty(directory))
            {
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                for (int i = 0; i < ImageCount; i++)
                    images[i].Save(string.Format("{0}\\_picture.{1}.bmp", directory, i));
            }
        }

        public ImageList(string filename): this(new MemoryStream(File.ReadAllBytes(filename)))
        {
            this.filename = filename;
        }


        int ReadInt(Stream ms)
        {
            //int Result;
            byte[] buffer = new byte[4];
            ms.Read(buffer, 0, 4);
            return BitConverter.ToInt32(buffer, 0);
        }

        short ReadWord(Stream ms)
        {
            return (short)(ms.ReadByte() | ms.ReadByte() << 8);
        }

        public ImageList(Stream ms)
        {
            
            //ms.Seek(0, SeekOrigin.Begin);
            InitCommonControls();
            ComStreamFromDataStream cs = new ComStreamFromDataStream(ms);
            this.Handle = ImageList_Read(cs);

            //Читаем длину BMP и находим его конец в потоке 
            //ms.Seek(10, SeekOrigin.Current); //28 - заголовок ImageList + 10 - указатель начала картинки после BM6
            //int IimageStart = ReadInt(ms);
            //ms.Seek(20, SeekOrigin.Current);
            //int IimageEnd = ReadInt(ms);
            ////IimageEnd += IimageStart + 38 + 28;
            //ms.Seek(IimageEnd + IimageStart - 38 + 44, SeekOrigin.Current);

            FAT.Header = (uint)ReadInt(ms);
            FAT.Length = ReadInt(ms);
            FAT.Indexes = new int[FAT.Length];
            for (int i = 0; i < FAT.Length; i++)
            {
                FAT.Indexes[i] = ReadInt(ms);
            }
            FAT.LastIndex = ReadWord(ms);
            int StrL = ms.ReadByte();

            while (StrL == 0)
                StrL = ms.ReadByte();

            if (StrL == 0xFF)
            {
                FAT.IdStringLength = ReadWord(ms);
            }

            if (StrL == 0x0F)
                FAT.IdStringLength = ms.ReadByte();

            byte[] buffer = new byte[FAT.IdStringLength];
            ms.Read(buffer, 0, FAT.IdStringLength);

            FAT.IdString = Encoding.GetEncoding(1251).GetString(buffer);

            GetInfo();
            GetImages();
        }

        public ImageList(int x, int y)
        {
            this.Handle = ImageList_Create(x, y, 0x00000004 , 1, 1);
        }

        public ImageList()
        {

        }

        public static ImageList LoadFromFile(string filename, int x = 80, int y = 80)
        {
            ImageList result = new ImageList(x, y);

            using (Bitmap data = new Bitmap(filename))
            {
                Bitmap region = new Bitmap(x, y);
                using (Graphics g = Graphics.FromImage(region))
                {
                    for (int j = 0; j < data.Height; j += y)
                        for (int i = 0; i < data.Width; i += x)
                        {
                            g.DrawImage(data, new Rectangle(0, 0, x, y), new Rectangle(i, j, x, y), GraphicsUnit.Pixel);
                            result.Add(region);
                        }
                }
            }

            result.GetImages();
            return result;
        }

        public int Construct(string Dirname)
        {
            List<string> files = Directory.EnumerateFiles(Dirname).ToList();
            files.Sort();

            FAT = new GalleryFAT { Length = 0, Header = 0xFF000001, IdString = "{}", IdStringLength = 2, Indexes = new int[0], LastIndex = 0x0F };

            foreach (string filename in files)
            {
                if(!Path.GetFileName(filename).Contains("Gallery.") && !filename.EndsWith(".dat"))
                    Add(filename);
            }

            return ImageCount;
        }


        public int Add(string filename)
        {
            using (Bitmap original = new Bitmap(filename))
            {
                Add(original);
                int Index = 0;
                if (Int32.TryParse(Path.GetFileNameWithoutExtension(filename).Split(new char[] { '.' })[1], NumberStyles.HexNumber, CultureInfo.CurrentCulture, out Index))
                {
                    FAT.Indexes[FAT.Indexes.Length - 1] = Index;
                    FAT.LastIndex = (byte)(Index);
                }
                else
                    FAT.Indexes[FAT.Indexes.Length - 1] = FAT.LastIndex + ImageCount << 16;

                FAT.IdString += "{\"" + filename + "\",\"" + Index.ToString() + "\"} ";

                FAT.Length = ImageCount;
                return ImageCount;
            }
        }

        public static Bitmap ResizeImage(Image image, float width, float height)
        {

            float Wamp = width / Math.Max(image.Width, image.Height);

            int x = (int)(image.Width * Wamp);
            int y = (int)(image.Height * Wamp);

            y = y < 20 ? 20 : y;

            var destImage = new Bitmap((int)width, (int)height);
            var destRect = new Rectangle((destImage.Width - x) / 2, (destImage.Height - y) / 2, x, y);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceOver;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
                graphics.SmoothingMode = SmoothingMode.None;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphics.FillRectangle(Brushes.AntiqueWhite, new Rectangle(0,0, destImage.Width, destImage.Height));
                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.Clamp);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }


        public int Add(Bitmap original)
        {
            
            if (original.Width != 80 || original.Height != 80)
                return AddResize(original);
            else
            {
                //IntPtr hMask = ControlPaint.CreateHBitmapTransparencyMask(original);   // Calls GDI to create Bitmap.
                IntPtr hBitmap = original.GetHbitmap(); // Calls GDI+ to create Bitmap. Need to add handle to HandleCollector.
                int index = ImageList_Add(new HandleRef(this, Handle), new HandleRef(null, hBitmap), new HandleRef(null, IntPtr.Zero));

                Array.Resize<int>(ref FAT.Indexes, FAT.Indexes.Length + 1);

                DeleteObject(new HandleRef(null, hBitmap));
                //DeleteObject(new HandleRef(null, hMask));
                ImageCount = ImageList_GetImageCount(new HandleRef(this, this.Handle));

                return index;
            }
        }

        public int AddResize(Bitmap original)
        {
            using (Bitmap bitmap = ResizeImage(original, 80, 80))
            {
                return Add(bitmap); 
            }
        }
    }
}
