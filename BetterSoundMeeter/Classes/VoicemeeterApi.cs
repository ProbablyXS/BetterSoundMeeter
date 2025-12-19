using System;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsFormsApp3
{
    public static class VoicemeeterApi
    {
        public static int retryConnection = 0;

        // Voicemeeter API functions
        [DllImport("VoicemeeterRemote64.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int VBVMR_Login();

        [DllImport("VoicemeeterRemote64.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int VBVMR_Logout();

        [DllImport("VoicemeeterRemote64.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int VBVMR_SetParameterString(string param, string value);

        [DllImport("VoicemeeterRemote64.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int VBVMR_SetParameterFloat(string param, float value);

        [DllImport("VoicemeeterRemote64.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int VBVMR_GetParameterString(string param, StringBuilder value, int maxLength);

        [DllImport("VoicemeeterRemote64.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int VBVMR_GetParameterFloat(string param, ref float value);

        // Enum representing Voicemeeter parameters
        public enum VoicemeeterCommand
        {
            Bus0_Gain = 0,
            Bus1_Gain = 1,
            Bus2_Gain = 2,
            Bus3_Gain = 3,
            Recorder_Gain = 10,
            Recorder_Goto = 11,
            MasterVolume = 20,
            A1Gain = 30,
            A2Gain = 31,
            A3Gain = 32,
            A4Gain = 33,
            MuteAll = 50,
            Recorder_Load = 60,
            Recorder_Play = 61,
            Recorder_Stop = 62,
        }

        // Method to get the parameter name based on the enum
        public static string GetParameterName(VoicemeeterCommand command)
        {
            switch (command)
            {
                case VoicemeeterCommand.Bus0_Gain:
                    return "Bus[0].Gain";
                case VoicemeeterCommand.Bus1_Gain:
                    return "Bus[1].Gain";
                case VoicemeeterCommand.Bus2_Gain:
                    return "Bus[2].Gain";
                case VoicemeeterCommand.Bus3_Gain:
                    return "Bus[3].Gain";
                case VoicemeeterCommand.Recorder_Gain:
                    return "Recorder.Gain";
                case VoicemeeterCommand.Recorder_Goto:
                    return "Recorder.goto";
                case VoicemeeterCommand.MasterVolume:
                    return "MasterVolume";
                case VoicemeeterCommand.A1Gain:
                    return "A1Gain";
                case VoicemeeterCommand.A2Gain:
                    return "A2Gain";
                case VoicemeeterCommand.A3Gain:
                    return "A3Gain";
                case VoicemeeterCommand.A4Gain:
                    return "A4Gain";
                case VoicemeeterCommand.MuteAll:
                    return "MuteAll";
                case VoicemeeterCommand.Recorder_Load:
                    return "Recorder.Load";
                case VoicemeeterCommand.Recorder_Play:
                    return "Recorder.Play";
                case VoicemeeterCommand.Recorder_Stop:
                    return "Recorder.Stop";
                default:
                    throw new ArgumentException("Unknown command");
            }
        }

        public static int SetParameter<T>(VoicemeeterCommand command, T value)
        {
            string parameterName = GetParameterName(command);

            if (value is float f)
                return VBVMR_SetParameterFloat(parameterName, f);
            if (value is double d)
                return VBVMR_SetParameterFloat(parameterName, (float)d);
            if (value is string s)
                return VBVMR_SetParameterString(parameterName, s);

            throw new ArgumentException($"Unsupported type {typeof(T)} for {parameterName}");
        }




        public static float GetParameter(VoicemeeterCommand command)
        {
            string parameterName = GetParameterName(command);
            float value = 0.0f;
            VBVMR_GetParameterFloat(parameterName, ref value);
            return value;
        }


        public static bool IsConnected()
        {
            StringBuilder value = new StringBuilder(1024);
            int result = VBVMR_GetParameterString("Strip[1]/Volume", value, value.Capacity);

            if (result == -3)
            {
                //Console.WriteLine($"Voicemeeter Connected!");
                return true;
            }

            if (result == -1 || result == -2)
            {
                Console.WriteLine("Voicemeeter not connected / not running / crash ??");

                if (retryConnection >= 2)
                {
                    VoicemeeterApi.VBVMR_Logout();
                    VoicemeeterApi.VBVMR_Login();
                    retryConnection = 0;
                    return false;
                }

                retryConnection += 1;
                return false;
            }

            Console.WriteLine($"Voicemeeter API error: {result}");
            return false;
        }
    }
}
