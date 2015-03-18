using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ImageResizer
{
    internal static class Program
    {
        //MUTEX for unique process identifier
        private const uint BCM_SETSHIELD = 0x160C;
        private static Mutex m = new Mutex(true, "{841e2cfc-fc06-4764-a892-3b7e818b35f9}");
        public static String[] arguments = {""};

        [DllImport("user32", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, IntPtr lParam);

        public static void startServer2()
        {
            Task.Factory.StartNew(() =>
            {
                using (var pipeStream = new NamedPipeServerStream("PipeToImageResizer"))
                {
                    pipeStream.WaitForConnection();

                    using (var sr = new StreamReader(pipeStream))
                    {
                        string message;
                        while ((message = sr.ReadLine()) != null)
                        {
                        }
                        startServer();
                    }
                }
            });
        }

        //runs pipeline server for inter process communication
        public static void startServer()
        {
            String str = "";
            var bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.DoWork += new DoWorkEventHandler(delegate(object o, DoWorkEventArgs args)
            {
                using (var pipeStream = new NamedPipeServerStream("PipeToImageResizer"))
                {
                    while (!pipeStream.IsConnected)
                    {
                        pipeStream.WaitForConnection();
                    }


                    using (var sr = new StreamReader(pipeStream))
                    {
                        string message;
                        while ((message = sr.ReadLine()) != null)
                        {
                            str = message;
                        }
                    }
                }
            });
            bw.ProgressChanged += new ProgressChangedEventHandler(delegate(object o, ProgressChangedEventArgs args) { });
            bw.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(delegate(object o, RunWorkerCompletedEventArgs args)
                {
                    var regPath = new Regex("@(.*)");
                    var regFilename = new Regex("(.*)@");
                    Match matchPath = regPath.Match(str);
                    Match matchFilename = regFilename.Match(str);
                    string path = "";
                    string filename = "";
                    if (matchPath.Success && matchFilename.Success)
                    {
                        path = matchPath.Value.ToString().Replace("@", "");
                        filename = matchFilename.Value.ToString().Replace("@", "");
                    }
                    Form1._Form1.addToList(filename, path);
                    startServer();
                });
            bw.RunWorkerAsync();
        }

        //runs client for inter process communication methods
        public static void startClient(string str)
        {
            using (var pipeStream = new NamedPipeClientStream("PipeToImageResizer"))
            {
                if (!pipeStream.IsConnected)
                {
                    pipeStream.Connect();
                }

                using (var sw = new StreamWriter(pipeStream))
                {
                    sw.AutoFlush = true;
                    string message = str;
                    sw.WriteLine(message);
                }
            }
        }

        public static void Register(string fileType, string shellKeyName, string menuText, string menuCommand)
        {
            string regPath = string.Format(@"{0}\shell\{1}", fileType, shellKeyName);

            if (Registry.ClassesRoot.OpenSubKey(regPath, false) == null)
            {
                try
                {
                    using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(regPath))
                    {
                        key.SetValue(null, menuText);
                    }

                    // add command that is invoked to the registry
                    using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(string.Format(@"{0}\command", regPath)))
                    {
                        key.SetValue(null, menuCommand);
                    }
                }
                catch (Exception)
                {
                    //if (string.IsNullOrEmpty((from o in arguments where o == "engage" select o).FirstOrDefault()))
                    //{
                    //    var btnElevate = new Button();
                    //    btnElevate.FlatStyle = FlatStyle.System;

                    //    SendMessage(btnElevate.Handle, BCM_SETSHIELD, 0, (IntPtr) 1);

                    //    var processInfo = new ProcessStartInfo();
                    //    processInfo.Verb = "runas";
                    //    processInfo.FileName = Application.ExecutablePath;
                    //    processInfo.Arguments = string.Join(" ", arguments.Concat(new[] {"engage"}).ToArray());

                    //    Process p = Process.Start(processInfo);
                    //    p.WaitForExit();
                    //    Application.Exit();
                    //}
                }
            }
        }

        public static void Unregister(string fileType, string shellKeyName)
        {
            Debug.Assert(!string.IsNullOrEmpty(fileType) &&
                         !string.IsNullOrEmpty(shellKeyName));

            // path to the registry location
            string regPath = string.Format(@"{0}\shell\{1}", fileType, shellKeyName);

            // remove context menu from the registry
            Registry.ClassesRoot.DeleteSubKeyTree(regPath);
        }


        [STAThread]
        private static void Main(string[] args)
        {
            string menuCommand = string.Format("\"{0}\" \"%L\"", Application.ExecutablePath);

            Register("*", "Resize Images", "Resize Images", menuCommand);

            //checks wether process with same mutex exists
            if (m.WaitOne(TimeSpan.Zero, true))
            {
                startServer();

                if (args.Length > 0)
                {
                    arguments = args;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                Application.Run(new Form1());
            }
            //sends commandline args if process running
            else
            {
                String arg = "";
                string filename = "";
                string path = "";
                if (args.Length > 0)
                {
                    try
                    {
                        filename = Path.GetFileName(args[0]);
                        path = args[0];
                        arg = filename + "@" + path;
                    }
                    catch (Exception)
                    {
                    }
                }
                startClient(arg);
            }
        }
    }
}