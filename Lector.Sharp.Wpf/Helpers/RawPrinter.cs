using System;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Lector.Sharp.Wpf.Helpers
{
    public class RawPrinter
    {
        #region Structure and API declarions ...

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        private class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
        }

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StartDocPrinter(IntPtr hPrinter, int level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

        [DllImport("winspool.drv", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern bool GetDefaultPrinter(StringBuilder pszBuffer, ref int size);

        #endregion

        #region Private static Methods ...

        // SendBytesToPrinter()
        // When the function is given a printer name and an unmanaged array
        // of bytes, the function sends those bytes to the print queue.
        // Returns true on success, false on failure.
        private static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, int dwCount)
        {
            int dwError = 0, dwWritten = 0;
            IntPtr hPrinter = new IntPtr(0);
            DOCINFOA di = new DOCINFOA();
            bool bSuccess = false; // Assume failure unless you specifically succeed.
            di.pDocName = "My C#.NET RAW Document";
            di.pDataType = "RAW";

            // Open the printer.
            if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
            {
                // Start a document.
                if (StartDocPrinter(hPrinter, 1, di))
                {
                    // Start a page.
                    if (StartPagePrinter(hPrinter))
                    {
                        // Write your bytes.
                        bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                        EndPagePrinter(hPrinter);
                    }
                    EndDocPrinter(hPrinter);
                }
                ClosePrinter(hPrinter);
            }
            // If you did not succeed, GetLastError may give more information
            // about why not.
            if (bSuccess == false)
            {
                dwError = Marshal.GetLastWin32Error();
            }
            return bSuccess;
        }

        private static bool SendFileToPrinter(string szPrinterName, string szFileName)
        {
            // Open the file.
            FileStream fs = new FileStream(szFileName, FileMode.Open);
            // Create a BinaryReader on the file.
            BinaryReader br = new BinaryReader(fs);
            // Dim an array of bytes big enough to hold the file's contents.
            byte[] bytes = new byte[fs.Length];
            bool bSuccess = false;
            // Your unmanaged pointer.
            IntPtr pUnmanagedBytes = new IntPtr(0);
            int nLength;

            nLength = Convert.ToInt32(fs.Length);
            // Read the contents of the file into the array.
            bytes = br.ReadBytes(nLength);
            // Allocate some unmanaged memory for those bytes.
            pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
            // Copy the managed byte array into the unmanaged array.
            Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
            // Send the unmanaged bytes to the printer.
            bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);
            // Free the unmanaged memory that you allocated earlier.
            Marshal.FreeCoTaskMem(pUnmanagedBytes);
            return bSuccess;
        }

        private static bool SendStringToPrinter(string szPrinterName, string szString)
        {
            IntPtr pBytes;
            int dwCount;

            // How many characters are in the string?
            // Fix from Nicholas Piasecki:
            // dwCount = szString.Length;
            dwCount = (szString.Length + 1) * Marshal.SystemMaxDBCSCharSize;

            // Assume that the printer is expecting ANSI text, and then convert
            // the string to ANSI text.
            pBytes = Marshal.StringToCoTaskMemAnsi(szString);

            // Assume that the printer is expecting ANSI text, and then convert
            // the string to Unicode text.
            //pBytes = Marshal.StringToCoTaskMemUni(szString);

            // Send the converted ANSI string to the printer.
            SendBytesToPrinter(szPrinterName, pBytes, dwCount);
            Marshal.FreeCoTaskMem(pBytes);
            return true;
        }

        #endregion

        #region ESC/POS Commands ...

        // ESC/POS Command Set for Reliance Thermal Printer
        // http://reliance-escpos-commands.readthedocs.io/en/latest/index.html
        public static readonly string Initialize = Convert.ToString((char)27) + Convert.ToString((char)64);
        public static readonly string CutPaper = Convert.ToString((char)27) + Convert.ToString((char)105);
        public static readonly string FontDefault = Convert.ToString((char)27) + Convert.ToString((char)77) + Convert.ToString((char)0);
        public static readonly string FontSmall = Convert.ToString((char)27) + Convert.ToString((char)77) + Convert.ToString((char)1);
        public static readonly string AlignCenter = Convert.ToString((char)27) + Convert.ToString((char)97) + Convert.ToString((char)1);
        public static readonly string AlignRight = Convert.ToString((char)27) + Convert.ToString((char)97) + Convert.ToString((char)2);
        public static readonly string AlignLeft = Convert.ToString((char)27) + Convert.ToString((char)97) + Convert.ToString((char)0);
        public static readonly string Underline = Convert.ToString((char)27) + Convert.ToString((char)45) + Convert.ToString((char)1);
        public static readonly string NoUnderline = Convert.ToString((char)27) + Convert.ToString((char)45) + Convert.ToString((char)0);
        public static readonly string Bold = Convert.ToString((char)27) + Convert.ToString((char)69) + Convert.ToString((char)1);
        public static readonly string NoBold = Convert.ToString((char)27) + Convert.ToString((char)69) + Convert.ToString((char)0);
        public static readonly string SizeH1 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)0);
        public static readonly string SizeH2 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)17);
        public static readonly string SizeH4 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)51);
        public static readonly string SizeH8 = Convert.ToString((char)29) + Convert.ToString((char)33) + Convert.ToString((char)119);
        public static readonly string NewLine = Convert.ToString((char)10);
        public static readonly string BarCodeHeightX20 = Convert.ToString((char)29) + Convert.ToString((char)104) + Convert.ToString((char)20);
        public static readonly string BarCodeBelow = Convert.ToString((char)29) + Convert.ToString((char)72) + Convert.ToString((char)2);
        public static readonly string BarCodeFormA = Convert.ToString((char)29) + Convert.ToString((char)107) + Convert.ToString((char)67) + Convert.ToString((char)13);
        public static readonly string BarCodeFormB = Convert.ToString((char)29) + Convert.ToString((char)107) + Convert.ToString((char)2);
        #endregion

        public string Text { get; private set; }
        public string PrinterName { get; set; }

        public RawPrinter()
        {
            Text = string.Empty;
            PrinterName = GetDefaultPrinter();
        }

        public string NormalizeCharacters(string text)
        {
            text = text.Replace('Á', 'A').Replace("&Aacute;", "A")
                .Replace('á', 'a').Replace("&aacute;", "a")
                .Replace('À', 'A').Replace("&Agrave;", "A")
                .Replace('à', 'a').Replace("&agrave;", "a")
                .Replace('Ä', 'A').Replace("&Auml;", "A")
                .Replace('ä', 'a').Replace("&auml;", "a")

                .Replace('É', 'E').Replace("&Eacute;", "E")
                .Replace('é', 'e').Replace("&eacute;", "e")
                .Replace('È', 'E').Replace("&Egrave;", "E")
                .Replace('è', 'e').Replace("&egrave;", "e")
                .Replace('Ë', 'E').Replace("&Euml;", "E")
                .Replace('ë', 'e').Replace("&euml;", "e")

                .Replace('Í', 'I').Replace("&Iacute;", "I")
                .Replace('í', 'i').Replace("&iacute;", "i")
                .Replace('Ì', 'I').Replace("&Igrave;", "I")
                .Replace('ì', 'i').Replace("&igrave;", "i")
                .Replace('Ï', 'I').Replace("&Iuml;", "I")
                .Replace('ï', 'i').Replace("&iuml;", "i")

                .Replace('Ó', 'O').Replace("&Oacute;", "O")
                .Replace('ó', 'o').Replace("&oacute;", "o")
                .Replace('Ò', 'O').Replace("&Ograve;", "O")
                .Replace('ò', 'o').Replace("&ograve;", "o")
                .Replace('Ö', 'O').Replace("&Ouml;", "O")
                .Replace('ö', 'o').Replace("&ouml;", "o")

                .Replace('Ú', 'U').Replace("&Uacute;", "U")
                .Replace('ú', 'u').Replace("&uacute;", "u")
                .Replace('Ù', 'U').Replace("&Ugrave;", "U")
                .Replace('ù', 'u').Replace("&ugrave;", "u")
                .Replace('Ü', 'U').Replace("&Uuml;", "U")
                .Replace('ü', 'u').Replace("&uuml;", "u")

                .Replace('Ñ', 'N').Replace("&Ntilde;", "N")
                .Replace('ñ', 'n').Replace("&ntilde;", "n")

                .Replace('Ç', 'C').Replace("&Ccedil;", "C")
                .Replace('ç', 'c').Replace("&ccedil;", "c")

                .Replace('€', 'E').Replace("&euro;", "E")

                .Replace('º', '#').Replace("&ordm;", "#")
                .Replace('ª', '#').Replace("&ordf;", "#");

            return text;
        }

        public string GetDefaultPrinter()
        {
            PrintDocument pd = new PrintDocument();
            StringBuilder dp = new StringBuilder(256);
            int size = dp.Capacity;
            if (GetDefaultPrinter(dp, ref size))
            {
                pd.PrinterSettings.PrinterName = dp.ToString().Trim();
            }
            return pd.PrinterSettings.PrinterName;
        }

        public void Clear()
        {
            Text = string.Empty;
        }

        public void Draw(string text)
        {
            Text += text;
        }

        public void Draw(string text, Align align, uint width)
        {
            string textDrawing = string.Empty;

            switch (align)
            {
                case Align.Center:
                    var margin = new string(' ', (int)Math.Truncate(width - text.Length / (decimal)2));
                    textDrawing = margin + text + margin;
                    break;
                case Align.Left:
                    var left = new string(' ', (int)width - text.Length);
                    textDrawing = text + left;
                    break;
                case Align.Right:
                    var right = new string(' ', (int)width - text.Length);
                    textDrawing = right + text;
                    break;
            }

            Text += textDrawing;
        }

        public void DrawLine(string text)
        {
            Text += NewLine + text;
        }

        public void DrawLine(string text, Align align)
        {
            switch (align)
            {
                case Align.Center:
                    Text += NewLine + AlignCenter + text;
                    break;
                case Align.Left:
                    Text += NewLine + AlignLeft + text;
                    break;
                case Align.Right:
                    Text += NewLine + AlignRight + text;
                    break;
            }
        }

        public void DrawLine(string text, Align align, uint width)
        {
            Draw(NewLine);
            Draw(text, align, width);
        }

        public void SetAlign(Align align)
        {
            switch (align)
            {
                case Align.Center:
                    Text += NewLine + AlignCenter;
                    break;
                case Align.Left:
                    Text += NewLine + AlignLeft;
                    break;
                case Align.Right:
                    Text += NewLine + AlignRight;
                    break;
            }
        }

        public void BarCodeModelA(string barcode)
        {
            Text += NewLine + BarCodeHeightX20 + BarCodeBelow + BarCodeFormA + barcode + NewLine;
        }

        public void BarCodeModelB(string barcode)
        {
            Text += NewLine + BarCodeHeightX20 + BarCodeBelow + BarCodeFormB + barcode + NewLine;
        }

        public void BarCode(string bardcode, int model)
        {
            if (model == 765 || model == 605)
                BarCodeModelA(bardcode);
            else
                BarCodeModelB(bardcode);
        }

        public void Print()
        {
            string text = Initialize + Text + CutPaper;
            text = NormalizeCharacters(text);
            SendStringToPrinter(PrinterName, text);
        }

        public enum Align
        {
            Center,
            Left,
            Right
        }
    }
}