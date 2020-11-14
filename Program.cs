using System;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;

namespace ClickCaptureAnalyze
{
    public class MouseOperations
    {
        [Flags]
        public enum MouseEventFlags
        {
            LeftDown = 0x00000002,
            LeftUp = 0x00000004,
            MiddleDown = 0x00000020,
            MiddleUp = 0x00000040,
            Move = 0x00000001,
            Absolute = 0x00008000,
            RightDown = 0x00000008,
            RightUp = 0x00000010
        }

        // Used for SendInput
        // //////////////////
        internal struct MouseInput
        {
            public int X;
            public int Y;
            public uint MouseData;
            public uint Flags;
            public uint Time;
            public IntPtr ExtraInfo;
        }

        internal struct InputType
        {
            public static int INPUT_MOUSE = 0;
            public static int INPUT_KEYBOARD = 1;
            public static int INPUT_HARDWARE = 2;
        }

        internal struct Input
        {
            public int Type;
            public MouseInput MouseInput;
        }
        // //////////////////

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out MousePoint lpMousePoint);

        [DllImport("user32.dll")]
        // This function is deprecated, in the docs it says to use SendInput instead.
        // The reason for it, on stackoverflow somebody wrote that this function doesn't
        // have any way to indicate failure.
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern uint SendInput(uint nInputs, Input[] pInputs, int size);


        public static void SetCursorPosition(int x, int y)
        {
            SetCursorPos(x, y);
        }

        public static void SetCursorPosition(MousePoint point)
        {
            SetCursorPos(point.X, point.Y);
        }

        public static MousePoint GetCursorPosition()
        {
            MousePoint currentMousePoint;
            var gotPoint = GetCursorPos(out currentMousePoint);
            if (!gotPoint) { currentMousePoint = new MousePoint(0, 0); }
            return currentMousePoint;
        }

        public static void MouseEvent(MouseEventFlags value)
        {
            MousePoint position = GetCursorPosition();

            mouse_event((int)value, position.X, position.Y, 0, 0);
        }

        public static void MoveAndClickMouse(int x, int y)
        {
            // Get screen width and height
            var w = WIN32_API.GetSystemMetrics(WIN32_API.SM_CXSCREEN);
            var h = WIN32_API.GetSystemMetrics(WIN32_API.SM_CYSCREEN);

            var i = new Input[3];

            i[0].Type = InputType.INPUT_MOUSE;
            i[0].MouseInput.X = (x * 65535) / w;
            i[0].MouseInput.Y = (y * 65535) / h;
            i[0].MouseInput.Flags = (uint) (MouseEventFlags.Absolute | MouseEventFlags.Move);

            i[1].Type = InputType.INPUT_MOUSE;
            i[1].MouseInput.Flags = (uint)MouseEventFlags.LeftDown;

            i[2].Type = InputType.INPUT_MOUSE;
            i[2].MouseInput.Flags = (uint)MouseEventFlags.LeftUp;

            SendInput(3, i, Marshal.SizeOf(i[0]));
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MousePoint
        {
            public int X;
            public int Y;

            public MousePoint(int x, int y)
            {
                X = x;
                Y = y;
            }
        }
    }

    public class ScreenCapture
    {
        public static Bitmap GetEntireDesktopImage()
        {
            IntPtr m_HBitmap;
            WIN32_API.SIZE size;

            IntPtr hDC = WIN32_API.GetDC(WIN32_API.GetDesktopWindow());
            IntPtr hMemDC = WIN32_API.CreateCompatibleDC(hDC);

            size.cx = WIN32_API.GetSystemMetrics(WIN32_API.SM_CXSCREEN);
            size.cy = WIN32_API.GetSystemMetrics(WIN32_API.SM_CYSCREEN);

            m_HBitmap = WIN32_API.CreateCompatibleBitmap(hDC, size.cx, size.cy);

            if (m_HBitmap != IntPtr.Zero)
            {
                IntPtr hOld = (IntPtr)WIN32_API.SelectObject(hMemDC, m_HBitmap);
                WIN32_API.BitBlt(hMemDC, 0, 0, size.cx, size.cy, hDC, 0, 0, WIN32_API.SRCCOPY);
                WIN32_API.SelectObject(hMemDC, hOld);
                WIN32_API.DeleteDC(hMemDC);
                WIN32_API.ReleaseDC(WIN32_API.GetDesktopWindow(), hDC);
                return Image.FromHbitmap(m_HBitmap);
            }
            return null;
        }

        public static Bitmap GetPartialDesktopImage(int x, int y, int w, int h)
        {
            IntPtr m_HBitmap;

            IntPtr hDC = WIN32_API.GetDC(WIN32_API.GetDesktopWindow());
            IntPtr hMemDC = WIN32_API.CreateCompatibleDC(hDC);

            m_HBitmap = WIN32_API.CreateCompatibleBitmap(hDC, w, h);

            if (m_HBitmap != IntPtr.Zero)
            {
                IntPtr hOld = (IntPtr)WIN32_API.SelectObject(hMemDC, m_HBitmap);
                WIN32_API.BitBlt(hMemDC, 0, 0, w, h, hDC, x, y, WIN32_API.SRCCOPY);
                WIN32_API.SelectObject(hMemDC, hOld);
                WIN32_API.DeleteDC(hMemDC);
                WIN32_API.ReleaseDC(WIN32_API.GetDesktopWindow(), hDC);
                return Image.FromHbitmap(m_HBitmap);
            }
            return null;
        }

    }
    

    public class WIN32_API
    {
        public struct SIZE
        {
            public int cx;
            public int cy;
        }
        public const int SRCCOPY = 13369376;
        public const int SM_CXSCREEN = 0;
        public const int SM_CYSCREEN = 1;

        [DllImport("gdi32.dll", EntryPoint = "DeleteDC")]
        public static extern IntPtr DeleteDC(IntPtr hDc);

        [DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
        public static extern IntPtr DeleteObject(IntPtr hDc);

        [DllImport("gdi32.dll", EntryPoint = "BitBlt")]
        public static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest, IntPtr hdcSource, int xSrc, int ySrc, int RasterOp);

        [DllImport("gdi32.dll", EntryPoint = "CreateCompatibleBitmap")]
        public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll", EntryPoint = "CreateCompatibleDC")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll", EntryPoint = "SelectObject")]
        public static extern IntPtr SelectObject(IntPtr hdc, IntPtr bmp);

        [DllImport("user32.dll", EntryPoint = "GetDesktopWindow")]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll", EntryPoint = "GetDC")]
        public static extern IntPtr GetDC(IntPtr ptr);

        [DllImport("user32.dll", EntryPoint = "GetSystemMetrics")]
        public static extern int GetSystemMetrics(int abc);

        [DllImport("user32.dll", EntryPoint = "GetWindowDC")]
        public static extern IntPtr GetWindowDC(Int32 ptr);

        [DllImport("user32.dll", EntryPoint = "ReleaseDC")]
        public static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDc);
    }

class Program
    {
        static int screen_x, screen_y, screen_w, screen_h;
        static MouseOperations.MousePoint pos1, pos2;
        static int action_delay = 0;
        static int number_of_cycles = 0;

        static void Main(string[] args)
        {
            while (true)
            {
                Console.Write("Mouse info \"mi\", param set \"psm\"(manual) or \"psd\"(dynamic), analysis \"an\", exit \"x\"?\n>");
                var input = Console.ReadLine();
                
                if (input == "mi")
                    MouseInfo();
                else if (input == "psm")
                    ParamSetManual();
                else if (input == "psd")
                    ParamSetDynamic();
                else if (input == "an")
                    Analysis();
                else if (input == "x" || input == "exit" || input == "quit" || input == "q")
                    break;
                else
                    Console.WriteLine("Unknown command");
            }            
        }

        static void MouseInfo()
        {
            Console.WriteLine("Press any key to stop...");
            while (true)
            {
                if (Console.KeyAvailable)                
                {
                    Console.ReadKey(true); // intercept key so to not display it
                    break;
                }
                var pos = MouseOperations.GetCursorPosition();
                Console.WriteLine(pos.X + " " + pos.Y);
                Thread.Sleep(100);
            }
        }

        static void ParamSetManual()
        {
            // Read screen capture params
            Console.Write("Enter parameters for screen capture (x, y, w, h):");            
            
            var input = Console.ReadLine();
            string[] tokens = input.Split(',');

            if (tokens.Length == 4)
            {
                screen_x = Convert.ToInt32(tokens[0]);
                screen_y = Convert.ToInt32(tokens[1]);
                screen_w = Convert.ToInt32(tokens[2]);
                screen_h = Convert.ToInt32(tokens[3]);
            }

            // Read mouse position 1
            Console.Write("Enter parameters for mouse position 1: (x, y):");
            
            input = Console.ReadLine();
            tokens = input.Split(',');

            if (tokens.Length == 2)
            {
                pos1.X = Convert.ToInt32(tokens[0]);
                pos1.Y = Convert.ToInt32(tokens[1]);
            }            

            // Read mouse position 2
            Console.Write("Enter parameters for mouse position 2: (x, y):");
            
            input = Console.ReadLine();
            tokens = input.Split(',');

            if (tokens.Length == 2)
            {
                pos2.X = Convert.ToInt32(tokens[0]);
                pos2.Y = Convert.ToInt32(tokens[1]);
            }            

            // Get delay
            Console.Write("Enter delay in ms:");

            input = Console.ReadLine();

            try
            {
                int x = Convert.ToInt32(input);
                action_delay = x;
            }
            catch
            {
            }

            // Get number of cycles
            Console.Write("Enter number of cycles:");

            input = Console.ReadLine();

            try
            {
                int x = Convert.ToInt32(input);
                number_of_cycles = x;
            }
            catch
            {
            }

            Console.WriteLine("--------------\n" + screen_x + "," + screen_y + "," + screen_w + "," + screen_h + "\n" +
                pos1.X + "," + pos1.Y + "\n" +
                pos2.X + "," + pos2.Y + "\n" +
                action_delay + " ms\n" +
                number_of_cycles + "\n--------------");
        }

        static void ParamSetDynamic()
        {
            static MouseOperations.MousePoint GetMousePoint()
            {
                while (!Console.KeyAvailable) ;
                Console.ReadKey(true);
                return MouseOperations.GetCursorPosition();
            }

            // Get screen capture params
            Console.Write("Move to top left and press a key...");
            var pos = GetMousePoint();
            screen_x = pos.X;
            screen_y = pos.Y;
            Console.WriteLine(" \t\t(" + screen_x + "," + screen_y + ")");

            Console.Write("Move to bottom right and press a key...");
            pos = GetMousePoint();
            screen_w = pos.X - screen_x;
            screen_h = pos.Y - screen_y;
            Console.WriteLine(" \t(" + pos.X + "," + pos.Y + ")");

            // Get mouse position 1
            Console.Write("Move mouse to location 1 and press a key...");
            pos1 = GetMousePoint();
            Console.WriteLine(" \t(" + pos1.X + "," + pos1.Y + ")");

            // Get mouse position 2
            Console.Write("Move mouse to location 2 and press a key...");
            pos2 = GetMousePoint();
            Console.WriteLine(" \t(" + pos2.X + "," + pos2.Y + ")");

            // Get delay
            Console.Write("Enter delay in ms:");

            var input = Console.ReadLine();

            try
            {
                int x = Convert.ToInt32(input);
                action_delay = x;
            }
            catch
            {
            }

            // Get number of cycles
            Console.Write("Enter number of cycles:");

            input = Console.ReadLine();

            try
            {
                int x = Convert.ToInt32(input);
                number_of_cycles = x;
            }
            catch
            {
            }

            Console.WriteLine("--------------\n" + screen_x + "," + screen_y + "," + screen_w + "," + screen_h + "\n" +
                pos1.X + "," + pos1.Y + "\n" +
                pos2.X + "," + pos2.Y + "\n" +
                action_delay + " ms\n" +
                number_of_cycles + "\n--------------");
        }

        static void Analysis()
        {
            Console.WriteLine("Press any key to stop...");

            for (int it = 1; it <= number_of_cycles; ++it)
            {
                if (Console.KeyAvailable)
                {
                    Console.ReadKey(true); // intercept key so to not display it
                    break;
                }

                //MouseOperations.SetCursorPosition(pos1);
                //MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
                //MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);
                MouseOperations.MoveAndClickMouse(pos1.X, pos1.Y);

                Thread.Sleep(action_delay);

                var total = CaptureAndAnalyzeImage(true);
                Console.Write(it + ". " + total);

                //MouseOperations.SetCursorPosition(pos2);
                //MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
                //MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);
                MouseOperations.MoveAndClickMouse(pos2.X, pos2.Y);

                Thread.Sleep(action_delay);

                total = CaptureAndAnalyzeImage(true);
                Console.WriteLine(" " + total);
            }
        }

        static int CaptureAndAnalyzeImage(bool save)
        {
            var image = ScreenCapture.GetPartialDesktopImage(screen_x, screen_y,
                    screen_w, screen_h);

            var bmpData = image.LockBits(Rectangle.FromLTRB(0, 0, screen_w, screen_h),
                ImageLockMode.ReadOnly, PixelFormat.Format8bppIndexed);

            // Get the address of the first line.
            IntPtr ptr = bmpData.Scan0;

            // Declare an array to hold the bytes of the bitmap.
            int bytes = Math.Abs(bmpData.Stride) * image.Height;
            byte[] byte_array = new byte[bytes];

            // Copy the RGB values into the array.
            Marshal.Copy(ptr, byte_array, 0, bytes);

            int total = 0;
            for (int i = 0; i < byte_array.Length; ++i)
                total += byte_array[i];            

            if (save)
            {
                string dir_name = "CCA_images/";
                if (!Directory.Exists(dir_name))
                    Directory.CreateDirectory(dir_name);
                DateTime timestamp = DateTime.Now;
                String tsstr = timestamp.ToString("CCA_yyyy-MM-dd_HH-mm-ss.ff", CultureInfo.CreateSpecificCulture("en-US"));
                String tfname = tsstr + "_" + total + ".png";
                image.Save(dir_name + tfname, ImageFormat.Png);
            }

            return total;
        }
    }
}
