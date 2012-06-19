using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JuraScriptLibrary
{

    public class JSConsole {
        public static void WriteLine(object arg) {
            Console.WriteLine(arg);
        }
        public static string ReadLine() {
            return Console.ReadLine();
        }
        public static ConsoleColor BackgroundColor {
            get {
                return Console.BackgroundColor;
            }
            set {
                Console.BackgroundColor = value;
            }
        }
        public static ConsoleColor ForegroundColor {
            get {
                return Console.ForegroundColor;
            }
            set {
                Console.ForegroundColor = value;
            }
        }
        public static void Beep() {
            Console.Beep();
        }
        public static void Beep(int freq, int dur) {
            Console.Beep(freq, dur);
        }
        public static int BufferHeight {
            get {
                return Console.BufferHeight;
            }
            set {
                Console.BufferHeight = value;
            }
        }
        public static int BufferWidth {
            get {
                return Console.BufferWidth;
            }
            set {
                Console.BufferWidth = value;
            }
        }
        public static bool CapsLock { get { return Console.CapsLock; } }
        public static void Clear() {
            Console.Clear();
        }
        public static int CursorLeft {
            get {
                return Console.CursorLeft;
            }
            set {
                Console.CursorLeft = value;
            }
        }
        public static int CursorSize {
            get {
                return Console.CursorSize;
            }
            set {
                Console.CursorSize = value;
            }
        }
        public static int CursorTop {
            get {
                return Console.CursorTop;
            }
            set {
                Console.CursorTop = value;
            }
        }
        public static bool CursorVisible {
            get {
                return Console.CursorVisible;
            }
            set {
                Console.CursorVisible = value;
            }
        }
        public static object Error {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }
        public static object In {
            get {
                throw new NotImplementedException();
            }
            set {
                throw new NotImplementedException();
            }
        }
        public static Encoding InputEncoding {
            get {
                return Console.InputEncoding;
            }
            set {
                Console.InputEncoding = value;
            }

        }
        public static bool KeyAvailable {
            get {
                return Console.KeyAvailable;
            }
        }
        public static int LargestWindowHeight {
            get {
                return Console.LargestWindowHeight;
            }
        }
    }
}
