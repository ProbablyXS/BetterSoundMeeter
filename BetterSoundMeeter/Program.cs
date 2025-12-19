using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    internal static class Program
    {

        [STAThread]
        static void Main()
        {

            // SINGLE INSTANCE APP
            var appName = Assembly.GetEntryAssembly().GetName().Name;
            bool createdNew;
            var mutex = new Mutex(true, appName, out createdNew);

            if (!createdNew)
            {
                return;
            }




            var assem = Assembly.GetExecutingAssembly();
            var resourceNames = assem.GetManifestResourceNames();
            Console.WriteLine("List of embedded resources:");
            foreach (var name in resourceNames)
            {
                Console.WriteLine(name);
            }

            foreach (var resourceName in resourceNames)
            {
                if (resourceName.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
                {
                    using (var stream = assem.GetManifestResourceStream(resourceName))
                    {
                        if (stream != null)
                        {
                            byte[] assemblyData = new byte[stream.Length];
                            stream.Read(assemblyData, 0, assemblyData.Length);

                            string dllName = resourceName.Substring(resourceName.LastIndexOf("Dlls.") + 5);

                            dllName = Path.GetFileName(dllName);

                            string outputPath = Path.Combine(Application.StartupPath, dllName);

                            try
                            {
                                if (!File.Exists(outputPath))
                                {
                                    File.WriteAllBytes(outputPath, assemblyData);
                                    Console.WriteLine($"The DLL {dllName} has been successfully extracted to {outputPath}");
                                }
                            }
                            catch
                            {
                                Console.WriteLine($"Error extracting {dllName}");
                                Application.Exit();
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Error: The resource {resourceName} was not found.");
                        }
                    }
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
