using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        private string mp3Path = @"";
        private int selectedVirtualPlayKey = -1;
        private int selectedVirtualRestartKey = -1;
        private int selectedVirtualStopKey = -1;
        private int selectedMouseMessage = MouseHookConstants.WM_LBUTTONDOWN;
        private bool isKeyDown = false;
        private bool isMouseDown = false;
        private Thread spamThread;
        private int valTimer = 0;
        private float lastGotoSeconds = 0.0f;
        private bool useMouseMode = false;


        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            this.SetStyle(
                ControlStyles.UserPaint |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer, true);

            this.UpdateStyles();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams handleParam = base.CreateParams;
                handleParam.ExStyle |= 0x02000000;
                return handleParam;
            }
        }

        public Form1()
        {
            InitializeComponent();

            SetStyle(ControlStyles.UserPaint |
         ControlStyles.AllPaintingInWmPaint |
         ControlStyles.OptimizedDoubleBuffer, true);
            UpdateStyles();

            this.DoubleBuffered = true;

            SystemHooks._procMouse = HookMouseCallback;
            SystemHooks._procKeyboard = HookKeyboardCallback;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (VoicemeeterApi.VBVMR_Login() != 0)
            {
                label4.Text = "VoiceMeeter API: hooked";
                label4.ForeColor = Color.Green;
            }

            _ = voiceMeeterIsConnected();

            label3.Text = trackBar1.Value.ToString() + " ms";
            textBox1.Text = mp3Path;

            lastGotoSeconds = 00.00f;
            trackBar3.Value = (int)lastGotoSeconds;
            label10.Text = $"{lastGotoSeconds:F2} s";

            getSoundGain();

            comboBox1.Items.Clear();
            LoadEmbeddedSounds();
            LoadAllInputsIntoCombo();

            SystemHooks._hookMouseID = SetMouseHook(SystemHooks._procMouse);
            SystemHooks._hookKeyboardID = SetKeyboardHook(SystemHooks._procKeyboard);
        }


        public async Task voiceMeeterIsConnected()
        {
            while (true)
            {

                if (VoicemeeterApi.IsConnected())
                {
                    label4.Text = "VoiceMeeter API: hooked";
                    label4.ForeColor = Color.Green;
                }
                else
                {
                    label4.Text = "VoiceMeeter API: Not hooked";
                    label4.ForeColor = Color.Red;
                }

                await Task.Delay(1000);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            isMouseDown = false;
            isKeyDown = false;

            if (SystemHooks._hookMouseID != IntPtr.Zero)
                SystemHooks.UnhookWindowsHookEx(SystemHooks._hookMouseID);

            if (SystemHooks._hookKeyboardID != IntPtr.Zero)
                SystemHooks.UnhookWindowsHookEx(SystemHooks._hookKeyboardID);

            VoicemeeterApi.VBVMR_Logout();
            base.OnFormClosing(e);
        }

        private IntPtr SetMouseHook(SystemHooks.LowLevelMouseProc proc)
        {
            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SystemHooks.SetWindowsHookEx(
                    MouseHookConstants.WH_MOUSE_LL,
                    proc,
                    SystemHooks.GetModuleHandle(curModule.ModuleName),
                    0);
            }
        }

        private IntPtr SetKeyboardHook(SystemHooks.LowLevelKeyboardProc proc)
        {
            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SystemHooks.SetWindowsHookEx(
                    KeyboardHookConstants.WH_KEYBOARD_LL,
                    proc,
                    SystemHooks.GetModuleHandle(curModule.ModuleName),
                    0);
            }
        }

        // HOOK SOURIS
        private IntPtr HookMouseCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                if (!useMouseMode)
                    return SystemHooks.CallNextHookEx(SystemHooks._hookMouseID, nCode, wParam, lParam);

                int msg = wParam.ToInt32();

                if (msg == selectedMouseMessage)
                {
                    if (!isMouseDown)
                    {
                        isMouseDown = true;

                        if (spamThread == null || !spamThread.IsAlive)
                        {
                            valTimer = trackBar1.Value;
                            spamThread = new Thread(SpamRecorder)
                            {
                                IsBackground = true
                            };
                            spamThread.Start();
                        }
                    }
                }
                else
                {
                    int upMsg = 0;
                    if (selectedMouseMessage == MouseHookConstants.WM_LBUTTONDOWN)
                        upMsg = MouseHookConstants.WM_LBUTTONUP;

                    if (upMsg != 0 && msg == upMsg)
                    {
                        isMouseDown = false;
                    }
                }
            }

            return SystemHooks.CallNextHookEx(SystemHooks._hookMouseID, nCode, wParam, lParam);
        }

        // HOOK CLAVIER
        private IntPtr HookKeyboardCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int msg = wParam.ToInt32();

                if (msg == KeyboardHookConstants.WM_KEYDOWN || msg == KeyboardHookConstants.WM_SYSKEYDOWN)
                {
                    var kb = (SystemHooks.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(
                        lParam, typeof(SystemHooks.KBDLLHOOKSTRUCT));

                    if (kb.vkCode == selectedVirtualRestartKey && !isKeyDown)
                    {
                        isKeyDown = true;

                        if (spamThread == null || !spamThread.IsAlive)
                        {
                            valTimer = trackBar1.Value;
                            spamThread = new Thread(SpamRecorder)
                            {
                                IsBackground = true
                            };
                            spamThread.Start();
                        }
                    }

                    if (kb.vkCode == selectedVirtualPlayKey && !isKeyDown)
                    {
                        isKeyDown = false;

                        if (spamThread == null || !spamThread.IsAlive)
                        {
                            VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Play, "1");
                        }
                    }

                    if (kb.vkCode == selectedVirtualStopKey && !isKeyDown)
                    {
                        isKeyDown = false;

                        if (spamThread == null || !spamThread.IsAlive)
                        {
                            VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Stop, "1");
                        }
                    }
                }
                else if (msg == KeyboardHookConstants.WM_KEYUP || msg == KeyboardHookConstants.WM_SYSKEYUP)
                {
                    var kb = (SystemHooks.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(
                        lParam, typeof(SystemHooks.KBDLLHOOKSTRUCT));

                    if (kb.vkCode == selectedVirtualRestartKey)
                    {
                        isKeyDown = false;
                    }
                }
            }

            return SystemHooks.CallNextHookEx(SystemHooks._hookKeyboardID, nCode, wParam, lParam);
        }

        // SPAM
        private void SpamRecorder()
        {
            while (isMouseDown || isKeyDown)
            {
                VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Load, mp3Path);
                Thread.Sleep(valTimer);
            }

            if (checkBox1.Checked)
                VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Stop, "1");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Audio Files (*.mp3;*.wav)|*.mp3;*.wav|All Files (*.*)|*.*";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string shortPath = GetShortPathName(ofd.FileName); // 8.3 format = ASCII pur
                    textBox1.Text = Path.GetFileName(ofd.FileName);
                    mp3Path = ofd.FileName;

                    VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Load, shortPath);
                    Console.WriteLine($"Short path: {shortPath}"); // Debug
                }
            }
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            label3.Text = trackBar1.Value.ToString() + " ms";
            Console.WriteLine($"Successfully set Interval to {trackBar1.Value} ms");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            if (comboBox.SelectedItem != null)
            {
                string selectedFileName = comboBox.SelectedItem.ToString();

                string soundsDirectory = Path.Combine(Application.StartupPath, "Sounds");

                string subFolder = selectedFileName.Split('.')[0];

                string subFolderPath = Path.Combine(soundsDirectory, subFolder);
                string soundFilePath = Path.Combine(subFolderPath, selectedFileName);
                Console.WriteLine(subFolder);

                if (File.Exists(soundFilePath))
                {
                    mp3Path = soundFilePath;
                    VoicemeeterApi.VBVMR_SetParameterString("Recorder.load", mp3Path);
                    VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Load, mp3Path);

                    textBox1.Text = selectedFileName;

                    Console.WriteLine($"Song selected: {selectedFileName}");
                    Console.WriteLine($"File path: {soundFilePath}");
                }
                else
                {
                    MessageBox.Show($"File {selectedFileName} not found in 'Sounds' folder.");
                }
            }
        }

        private void getSoundGain()
        {
            float currentGain = (int)VoicemeeterApi.GetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Gain);
            int trackBarValue = (int)((currentGain + 60) * (trackBar2.Maximum - trackBar2.Minimum) / 72);
            trackBar2.Value = trackBarValue;
            label5.Text = currentGain.ToString() + " dB";
        }

        private void LoadEmbeddedSounds()
        {
            string soundsDirectory = Path.Combine(Application.StartupPath, "Sounds");
            if (!Directory.Exists(soundsDirectory))
            {
                Directory.CreateDirectory(soundsDirectory);
            }

            var assem = Assembly.GetExecutingAssembly();
            var resourceNames = assem.GetManifestResourceNames();

            var soundResources = resourceNames
                .Where(r => r.StartsWith("BetterSoundMeeter.Ressources.Sounds.") &&
                            (r.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase) ||
                             r.EndsWith(".wav", StringComparison.OrdinalIgnoreCase)))
                .ToList();

            comboBox1.Items.Clear();

            var groupedSounds = soundResources
                .Select(resourceName =>
                {
                    string soundName = resourceName.Substring("BetterSoundMeeter.Ressources.Sounds.".Length);
                    soundName = Path.GetFileName(soundName);
                    string prefix = soundName.Split('.')[0];
                    return new { SoundName = soundName, Prefix = prefix };
                })
                .GroupBy(x => x.Prefix)
                .OrderBy(g => g.Key)
                .ToList();

            foreach (var group in groupedSounds)
            {
                foreach (var item in group.OrderBy(x => x.SoundName))
                {
                    string soundName = item.SoundName;

                    using (var stream = assem.GetManifestResourceStream("BetterSoundMeeter.Ressources.Sounds." + soundName))
                    {
                        if (stream != null)
                        {
                            string folder = item.Prefix;
                            string soundFilePath = Path.Combine(soundsDirectory, folder);

                            if (!Directory.Exists(soundFilePath))
                            {
                                Directory.CreateDirectory(soundFilePath);
                            }

                            soundFilePath = Path.Combine(soundFilePath, soundName);

                            if (!File.Exists(soundFilePath))
                            {
                                using (var fileStream = new FileStream(soundFilePath, FileMode.Create))
                                {
                                    stream.CopyTo(fileStream);
                                    Console.WriteLine($"Sound {soundName} extracted and saved.");
                                }
                            }

                            comboBox1.Items.Add(soundName);
                        }
                    }
                }
            }


            Console.WriteLine("Sounds loaded into ComboBox:");
            foreach (var sound in comboBox1.Items)
            {
                Console.WriteLine(sound);
            }
        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            try
            {
                float minGain = -60.0f;
                float maxGain = 12.0f;

                float scaledValue = minGain + (trackBar2.Value / (float)trackBar2.Maximum) * (maxGain - minGain);

                int result = VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Gain, scaledValue);

                if (result == 0)
                {
                    Console.WriteLine($"Successfully set Recorder.Gain to {scaledValue:F2} dB");
                }
                else
                {
                    Console.WriteLine($"Failed to set Recorder.Gain. Error code: {result}");
                }

                label5.Text = $"{scaledValue:F2} dB";
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        // CHARGEMENT DES ENTRÉES DANS comboBox2
        private void LoadAllInputsIntoCombo()
        {
            comboBox2.Items.Clear();

            var keyFields = typeof(KeyboardHookConstants)
                .GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy)
                .Where(f => f.IsLiteral && !f.IsInitOnly && f.FieldType == typeof(int));

            foreach (var field in keyFields)
            {
                int vk = (int)field.GetRawConstantValue();
                string name = field.Name;
                string display = $"KEY_{name} (0x{vk:X2})";

                comboBox2.Items.Add(name.Substring("VK_".Length));
                comboBox3.Items.Add(name.Substring("VK_".Length));
                comboBox5.Items.Add(name.Substring("VK_".Length));
            }

            var mouseFields = typeof(MouseHookConstants)
                .GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy)
                .Where(f => f.IsLiteral && !f.IsInitOnly && f.FieldType == typeof(int));

            foreach (var field in mouseFields)
            {
                int code = (int)field.GetRawConstantValue();
                string name = field.Name;
                string display = $"MOUSE_{name} (0x{code:X4})";

                if (code != MouseHookConstants.WM_LBUTTONDOWN) continue;

                comboBox2.Items.Add(name.Substring("WM_".Length));
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem == null) return;

            string item = comboBox2.SelectedItem.ToString();

            if (item.StartsWith("KEY_"))
            {
                useMouseMode = false;
                selectedVirtualRestartKey = -1;

                int start = item.IndexOf("(0x");
                if (start >= 0)
                {
                    string hex = item.Substring(start + 3, item.Length - start - 4);
                    if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out int vk))
                    {
                        selectedVirtualRestartKey = vk;
                    }
                }
            }
            else if (item.StartsWith("MOUSE_"))
            {
                useMouseMode = true;
                selectedMouseMessage = -1;

                int start = item.IndexOf("(0x");
                if (start >= 0)
                {
                    string hex = item.Substring(start + 3, item.Length - start - 4);
                    if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out int code))
                    {
                        selectedMouseMessage = code;
                    }
                }
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedItem == null) return;

            string item = comboBox3.SelectedItem.ToString();

            if (item.StartsWith("KEY_"))
            {
                //useMouseMode = false;
                selectedVirtualStopKey = -1;

                int start = item.IndexOf("(0x");
                if (start >= 0)
                {
                    string hex = item.Substring(start + 3, item.Length - start - 4);
                    if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out int vk))
                    {
                        selectedVirtualStopKey = vk;
                        return;
                    }
                }
            }
        }

        private void trackBar3_Scroll(object sender, EventArgs e)
        {
            float newPosSeconds = trackBar3.Value;
            lastGotoSeconds = newPosSeconds;

            int setResult = VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Goto, newPosSeconds);
            if (setResult != 0)
            {
                Console.WriteLine($"Error setting playback position: {setResult}");
                return;
            }

            label10.Text = $"{lastGotoSeconds:F2} s";
            Console.WriteLine($"Last requested goto: {lastGotoSeconds:F2} s");
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.SelectedItem == null) return;

            string item = comboBox5.SelectedItem.ToString();

            if (item.StartsWith("KEY_"))
            {
                useMouseMode = false;
                selectedVirtualPlayKey = -1;

                int start = item.IndexOf("(0x");
                if (start >= 0)
                {
                    string hex = item.Substring(start + 3, item.Length - start - 4);
                    if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out int vk))
                    {
                        selectedVirtualPlayKey = vk;
                    }
                }
            }
        }

        private void trackBar3_MouseDown(object sender, MouseEventArgs e)
        {
            float newPosSeconds = trackBar3.Value;
            lastGotoSeconds = newPosSeconds;

            int setResult = VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Goto, newPosSeconds);
            if (setResult != 0)
            {
                Console.WriteLine($"Error setting playback position: {setResult}");
                return;
            }

            label10.Text = $"{lastGotoSeconds:F2} s";
            Console.WriteLine($"Last requested goto: {lastGotoSeconds:F2} s");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Play, "1");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Load, mp3Path);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            VoicemeeterApi.SetParameter(VoicemeeterApi.VoicemeeterCommand.Recorder_Stop, "1");
        }


        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        static extern uint GetShortPathName(string lpszLongPath, StringBuilder lpszShortPath, int cchBuffer);

        private static string GetShortPathName(string path)
        {
            StringBuilder sb = new StringBuilder(260);
            GetShortPathName(path, sb, sb.Capacity);
            return sb.ToString();
        }
    }
}
