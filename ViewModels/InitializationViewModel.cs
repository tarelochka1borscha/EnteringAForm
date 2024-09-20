using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Entering.ViewModels
{
    internal class InitializationViewModel : INotifyPropertyChanged
    {
        // сворованное
        #region

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int WH_KEYBOARD_LL = 13;
        private LowLevelKeyboardProcDelegate m_callback;
        private IntPtr m_hHook;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProcDelegate lpfn, IntPtr hMod, int dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("Kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetModuleHandle(IntPtr lpModuleName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr LowLevelKeyboardHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var khs = (KeyboardHookStruct)
                          Marshal.PtrToStructure(lParam,
                          typeof(KeyboardHookStruct));
                if (Convert.ToInt32("" + wParam) == 256)
                {
                    if ((int)khs.VirtualKeyCode == 113)//F2
                    {
                        _work = true;
                        Unhook();
                    }
                }
            }
            return CallNextHookEx(m_hHook, nCode, wParam, lParam);

        }
        [StructLayout(LayoutKind.Sequential)]
        private struct KeyboardHookStruct
        {
            public readonly int VirtualKeyCode;
            public readonly int ScanCode;
            public readonly int Flags;
            public readonly int Time;
            public readonly IntPtr ExtraInfo;
        }
        private delegate IntPtr LowLevelKeyboardProcDelegate(
            int nCode, IntPtr wParam, IntPtr lParam);

        public void SetHook()
        {
            m_callback = LowLevelKeyboardHookProc;
            m_hHook = SetWindowsHookEx(WH_KEYBOARD_LL,
                m_callback,
                GetModuleHandle(IntPtr.Zero), 0);
        }
        public void Unhook()
        {
            UnhookWindowsHookEx(m_hHook);
        }

        #endregion

        public event PropertyChangedEventHandler PropertyChanged;
        private void NotifyPropertyChanged(String propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public string Instruction => "1. Скопируйте ответы в буфер обмена\n" +
                "! Каждый ответ должен начинаться с новой строки\n" +
                "2. Установите курсор в первом поле, где начинается ввод ответов\n\n" +
                "F2 - запуск скрипта\n" +
                "После пробежки по всем ответам скрипт остановится самостоятельно";
        public string Buffer
        {
            get => _buffer;
            set
            {
                _buffer = value;
                NotifyPropertyChanged(nameof(Buffer));
            }
        }
        private string _buffer;

        private string all_answers;

        private bool _work = false;
        public InitializationViewModel()
        {
            Updating();
        }

        public async Task Updating()
        {
            while (!_work)
            {
                SetHook();
                var delayTask = Task.Delay(700);
                string buffer = System.Windows.Clipboard.GetText();
                if (Buffer != buffer) Buffer = System.Windows.Clipboard.GetText();
                await delayTask;
                Unhook();
            }
            MakeClick();
        }

        private string[] CutBuffer()
        {
            all_answers = Buffer;
            string[] answers = Buffer.Split('\n');
            for (int i = 0; i < answers.Length; i++)
            {
                int index = FirstChar(answers[i]);
                if (index < 0)
                {
                    answers[i] = "";
                }
                else
                {
                    answers[i] = answers[i].Substring(index, answers[i].Length - index);
                    answers[i] = answers[i].Trim();
                }
            }
            return answers;
        }

        private int FirstChar(string str)
        {
            Regex regex = new Regex("^[А-я]|[A-z]");
            int index = -1;
            for (int i = 0; i < str.Length; i++)
            {
                if (regex.IsMatch(str[i] + ""))
                {
                    index = i;
                    break;
                }
            }
            return index;
        }

        private async void MakeClick()
        {
            IntPtr windowHandler = GetForegroundWindow();
            SetForegroundWindow(windowHandler);
            string[] answers = CutBuffer();
            for (int i = 0; i < answers.Length; i++)
            {
                var delayTask = Task.Delay(500);
                System.Windows.Clipboard.SetText(answers[i]);
                SendKeys.SendWait("(^){v}");
                Task.Delay(500);
                SendKeys.SendWait("{TAB}");
                await delayTask;
            }
            _work = false;
            System.Windows.Clipboard.SetText(all_answers);
            Updating();
        }
    }
}
