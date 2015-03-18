using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ImageResizer.Properties;
using SHDocVw;
using Shell32;
using Encoder = System.Drawing.Imaging.Encoder;
using Timer = System.Windows.Forms.Timer;

namespace ImageResizer
{
    public partial class Form1 : Form
    {
        public static Form1 _Form1;
        public static long ramUsage = 0;


        private readonly Timer checkWindowTimer = new Timer();
        private readonly Timer timer1 = new Timer();
        public bool _shouldStop = false;
        public float aspectRatio = 0f;
        public bool selected = false;
        public bool windowActive = false;

        public Form1()
        {
            InitializeComponent();

            Timer_start();
            Application.ApplicationExit += Application_Exit;
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 1;
            comboBox4.SelectedIndex = 0;
            checkWindowTimer_start();
            checkBox1.FlatAppearance.BorderSize = 0;

            _Form1 = this;
            if (Program.arguments != null && Program.arguments.Length > 0)
            {
                string filename = "";
                string path = "";
                try
                {
                    filename = Path.GetFileName(Program.arguments[0]);
                    path = Program.arguments[0];
                }
                catch (Exception)
                {
                }

                if (filename != "" && path != "")
                {
                    _Form1.addToList(filename, path);
                }
            }
            loadSettings();
        }


        public void loadSettings()
        {
            comboBox2.SelectedIndex = (int) Settings.Default["resize"];
            comboBox1.SelectedIndex = (int) Settings.Default["interpolation"];
            if (comboBox2.SelectedIndex == 0)
            {
                textBox1.Text = Settings.Default["width"].ToString();
                textBox2.Text = Settings.Default["height"].ToString();
                checkBox1.Checked = (bool) Settings.Default["aspectratio"];
            }
            trackBar1.Value = (int) Settings.Default["slider"];
            label9.Text = trackBar1.Value.ToString();
        }

        public void saveSettings()
        {
            Settings.Default["resize"] = comboBox2.SelectedIndex;
            Settings.Default["interpolation"] = comboBox1.SelectedIndex;

            Settings.Default["slider"] = trackBar1.Value;
            if (comboBox2.SelectedIndex == 0)
            {
                Settings.Default["aspectratio"] = checkBox1.Checked;
            }

            Settings.Default.Save();
        }

        public static void checkRAM()
        {
            String str = "";
            var bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.DoWork +=
                new DoWorkEventHandler(
                    delegate(object o, DoWorkEventArgs args)
                    {
                        ramUsage = (Process.GetCurrentProcess().WorkingSet64/1024)/1024;
                    });
            bw.ProgressChanged += new ProgressChangedEventHandler(delegate(object o, ProgressChangedEventArgs args) { });
            bw.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(delegate(object o, RunWorkerCompletedEventArgs args)
                {
                    while (ramUsage >= 400)
                    {
                        MessageBox.Show("sleeping due to overload: " + ramUsage.ToString() + "MB");
                        Thread.Sleep(100);
                        ramUsage = (Process.GetCurrentProcess().WorkingSet64/1024)/1024;
                    }
                });
            bw.RunWorkerAsync();
        }


        public void Timer_start()
        {
            timer1.Tick += Timer_tick;
            timer1.Interval = 500;
            timer1.Start();
        }

        public void checkWindowTimer_start()
        {
            checkWindowTimer.Tick += checkWindowTimer_tick;
            checkWindowTimer.Interval = 5;
            checkWindowTimer.Start();
        }


        public void checkWindowTimer_tick(object sender, EventArgs e)
        {
            String filename;

            foreach (InternetExplorer window in new ShellWindows())
            {
                filename = Path.GetFileNameWithoutExtension(window.FullName).ToLower();
                //if (window.FullName.ToLower().Contains("explorer"))
                //{
                if (window.HWND == GetForegroundWindow().ToInt32())
                {
                    // MessageBox.Show(window.FullName);
                    if (window.FullName.Contains("Form1"))
                    {
                        timer1.Stop();
                    }

                    if (windowActive == false)
                    {
                        timer1.Start();
                        windowActive = true;
                    }
                }
                else
                {
                    if (windowActive)
                    {
                        //timer1.Stop();
                        //windowActive = false;
                    }
                }
                //}
            }
        }


        public void Timer_tick(object sender, EventArgs e)
        {
            ArrayList items = getSelectedFiles();

            if (listView1.Items.Count != items.Count)
            {
                if (windowActive)
                {
                    listView1.Items.Clear();
                }
            }
            foreach (FolderItem key in items)
            {
                if (listView1.Items.Count != items.Count)
                {
                    if (!listView1.Items.Contains(new ListViewItem(key.Name)))
                    {
                        var lvi = new ListViewItem(key.Name);
                        lvi.SubItems.Add(key.Path);
                        listView1.Items.Add(lvi);

                        if (textBox1.Text == "")
                        {
                            using (Image image = Image.FromFile(key.Path))
                            {
                                textBox1.Text = image.Width.ToString();
                            }
                        }
                        if (textBox2.Text == "")
                        {
                            using (Image image = Image.FromFile(key.Path))
                            {
                                textBox2.Text = image.Height.ToString();
                            }
                        }
                    }
                }
            }
            items.Clear();
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        public ArrayList getSelectedFiles()
        {
            IntPtr handle = GetForegroundWindow();
            string filename;
            var selected = new ArrayList();
            //var shell = new Shell32.Shell();
            foreach (InternetExplorer window in new ShellWindows())
            {
                filename = Path.GetFileNameWithoutExtension(window.FullName).ToLower();
                if (filename.ToLowerInvariant() == "explorer")
                {
                    if (window.HWND == handle.ToInt32())
                    {
                        try
                        {
                            FolderItems items = ((IShellFolderViewDual2) window.Document).SelectedItems();
                            foreach (FolderItem item in items)
                            {
                                // MessageBox.Show(item.Type.ToString());
                                if (item.Type.ToLower().Contains("jpg") || item.Type.ToLower().Contains("jpeg") ||
                                    item.Type.ToLower().Contains("png") || item.Type.ToLower().Contains("bmp") ||
                                    item.Type.ToLower().Contains("gif"))
                                {
                                    selected.Add(item);
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
            return selected;
        }

        public void addToList(string filename, string path)
        {
            timer1.Stop();
            //checkWindowTimer.Stop();
            var item = new ListViewItem(filename);
            item.SubItems.Add(path);
            listView1.Items.Add(item);
        }

        public int getX()
        {
            return int.Parse(textBox1.Text);
        }

        public int getY()
        {
            return int.Parse(textBox2.Text);
        }

        private void selectedItems_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void label4_Click(object sender, EventArgs e)
        {
        }

        private void label3_Click(object sender, EventArgs e)
        {
        }

        private void X_changed(object sender, EventArgs e)
        {
            if (textBox1.Focused)
            {
                if (checkBox1.Checked)
                {
                    try
                    {
                        textBox2.Text = (Math.Round(int.Parse(textBox1.Text)/aspectRatio)).ToString();
                    }
                    catch (Exception)
                    {
                        textBox2.Text = "0";
                    }
                }
                if (comboBox2.SelectedIndex == 0)
                {
                    Settings.Default["width"] = textBox1.Text;
                    Settings.Default["height"] = textBox2.Text;
                }
            }
        }

        private void resizeMethod(object sender, EventArgs e)
        {
        }

        private void Y_changed(object sender, EventArgs e)
        {
            if (textBox2.Focused)
            {
                if (checkBox1.Checked)
                {
                    try
                    {
                        textBox1.Text = (Math.Round(int.Parse(textBox2.Text)*aspectRatio)).ToString();
                    }
                    catch (Exception)
                    {
                        textBox1.Text = "0";
                    }
                }
                if (comboBox2.SelectedIndex == 0)
                {
                    Settings.Default["width"] = textBox1.Text;
                    Settings.Default["height"] = textBox2.Text;
                }
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {
        }

        private static Bitmap ChangePixelFormat(Bitmap oldBmp, PixelFormat NewFormat)
        {
            return (oldBmp.Clone(new Rectangle(0, 0, oldBmp.Width, oldBmp.Height), NewFormat));
        }

        public async void button1_Click(object sender, EventArgs e)
        {
            int listcount = 0;

            foreach (ListViewItem key in listView1.Items)
            {
                listcount++;
            }

            hideControls(this);

            progressBar1.Maximum = listcount;

            var pics = new ArrayList();
            var lvi = new ArrayList();

            foreach (ListViewItem key in listView1.Items)
            {
                lvi.Add(key);
            }
            foreach (ListViewItem key in lvi)
            {
                using (var bitmap = new Bitmap(Image.FromFile(key.SubItems[1].Text, true)))
                {
                    String path = key.SubItems[1].Text;
                    String name = key.SubItems[0].Text;
                    String format = ".jpg";
                    String suffix = "_resized";
                    if (checkBox2.Checked)
                    {
                        suffix = "";
                    }

                    ImageFormat imgFormat = ImageFormat.Jpeg;

                    if (comboBox4.Text.Equals("Automatic"))
                    {
                        if (key.SubItems[1].Text.EndsWith(".jpg") || key.SubItems[1].Text.EndsWith(".jpeg"))
                        {
                            path = path.Replace(".jpg", "");
                            format = ".jpg";
                            imgFormat = ImageFormat.Jpeg;
                        }
                        if (key.SubItems[1].Text.EndsWith(".png"))
                        {
                            path = path.Replace(".png", "");

                            format = ".png";
                            imgFormat = ImageFormat.Png;
                        }
                        if (key.SubItems[1].Text.EndsWith(".bmp"))
                        {
                            path = path.Replace(".bmp", "");

                            format = ".bmp";
                            imgFormat = ImageFormat.Bmp;
                        }
                        if (key.SubItems[1].Text.EndsWith(".gif"))
                        {
                            path = path.Replace(".gif", "");
                            format = ".gif";
                            imgFormat = ImageFormat.Gif;
                        }
                    }
                    else
                    {
                        if (comboBox4.Text.Equals("PNG"))
                        {
                            format = ".png";
                            imgFormat = ImageFormat.Png;
                        }
                        if (comboBox4.Text.Equals("JPEG"))
                        {
                            format = ".jpg";
                            imgFormat = ImageFormat.Jpeg;
                        }
                        if (comboBox4.Text.Equals("BMP"))
                        {
                            format = ".bmp";
                            imgFormat = ImageFormat.Bmp;
                        }
                        if (comboBox4.Text.Equals("GIF"))
                        {
                            format = ".gif";
                            imgFormat = ImageFormat.Gif;
                        }
                    }

                    String fileName = key.SubItems[0].Text;
                    //Bitmap img;


                    string outputFileName = path + suffix + format;

                    using (var memory = new MemoryStream())
                    {
                        bitmap.Save(memory, imgFormat);

                        using (Bitmap img = ResizeImage(memory, new Size(getX(), getY())))
                        {
                            if (checkBox2.Checked)
                            {
                                if (File.Exists(outputFileName))
                                {
                                    using (
                                        var fs = new FileStream(outputFileName, FileMode.OpenOrCreate,
                                            FileAccess.ReadWrite,
                                            FileShare.ReadWrite))
                                    {
                                        lvi = null;


                                        MessageBox.Show("File exists");
                                        File.Delete(outputFileName);
                                        byte[] bytes = memory.ToArray();
                                        fs.Write(bytes, 0, bytes.Length);
                                    }
                                }
                            }

                            long qualitySlider = trackBar1.Value;

                            var qualityParam = new EncoderParameter(Encoder.Quality, qualitySlider);

                            ImageCodecInfo codec = getCodec(format);

                            

                            var encoderParams = new EncoderParameters(1);
                            encoderParams.Param[0] = qualityParam;


                            img.Save(outputFileName, codec, encoderParams);
                            img.Dispose();
                        }
                        memory.Close();
                        memory.Dispose();
                    }
                    bitmap.Dispose();
                }
                progressBar1.Increment(1);
            }

            _shouldStop = false;
            showControls(this);
            progressBar1.Value = 0;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Stop();
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            // MessageBox.Show("activated");
            if (windowActive)
            {
                timer1.Stop();
                windowActive = false;
            }
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            //  timer1.Stop();
            //timer1.Start();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                checkBox1.BackColor = Color.Gray;
                try
                {
                    aspectRatio = getX()/(float) getY();
                }
                catch (Exception)
                {
                    MessageBox.Show("Define a size first");
                    checkBox1.Checked = false;
                }
            }
            else
            {
                checkBox1.BackColor = Color.Transparent;
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                using (var bitmap = new Bitmap(Image.FromFile(listView1.SelectedItems[0].SubItems[1].Text)))
                {
                    if (textBox1.Text == "")
                    {
                        textBox1.Text = bitmap.Width.ToString();
                    }
                    if (textBox2.Text == "")
                    {
                        textBox2.Text = bitmap.Height.ToString();
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        public static ImageCodecInfo getCodec(string codecString)

        {
            codecString.ToLower();
            ImageCodecInfo codec = null;

            foreach (ImageCodecInfo key in ImageCodecInfo.GetImageEncoders())
            {
                if (key.FilenameExtension.ToLower().Contains(codecString))
                {
                    codec = key;
                }
            }
            return codec;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            getCodec(".jpg");
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            label9.Text = trackBar1.Value + "%";
        }


        public void Application_Exit(object sender, EventArgs e)
        {
            saveSettings();
        }

        public void hideControls(Control control)
        {
            if (control.HasChildren)
            {
                foreach (Control child in control.Controls)
                {
                    child.Hide();
                    progressBar1.Show();
                }
            }
        }

        public void showControls(Control control)
        {
            if (control.HasChildren)
            {
                foreach (Control child in control.Controls)
                {
                    child.Show();
                    progressBar1.Hide();
                }
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
        }

        #region Resize Image

        public Bitmap ResizeImage(MemoryStream ms, Size size)
        {
            //MessageBox.Show("W: " + Wratio + " H: " + Hratio);
            Bitmap img = null;
            using (var original = new Bitmap(ms, true))
            {
                if (comboBox2.Text == "Pixel")
                {
                    img = new Bitmap(original, size.Width, size.Height);
                }
                if (comboBox2.Text == "Percent")
                {
                    var percentSize = new Size((int) (original.Width*(float) size.Width/100f),
                        (int) (original.Height*(float) size.Height/100f));
                    img = new Bitmap(original, percentSize.Width, percentSize.Height);
                }


                using (Graphics graphic = Graphics.FromImage(img))
                {
                    graphic.CompositingQuality = CompositingQuality.HighQuality;


                    if (comboBox1.Text == "Pixelated")
                    {
                        graphic.InterpolationMode = InterpolationMode.NearestNeighbor;
                    }
                    if (comboBox1.Text == "Bilinear")
                    {
                        graphic.InterpolationMode = InterpolationMode.Bilinear;
                    }
                    if (comboBox1.Text == "Bicubic")
                    {
                        graphic.InterpolationMode = InterpolationMode.Bicubic;
                    }
                    if (comboBox1.Text == "High Quality Bicubic")
                    {
                        graphic.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    }
                    if (comboBox1.Text == "High Quality Bilinear")
                    {
                        graphic.InterpolationMode = InterpolationMode.HighQualityBilinear;
                    }

                    if (comboBox2.Text == "Pixel")
                    {
                        float Wratio = original.Width/(float) size.Width;
                        float Hratio = original.Height/(float) size.Height;
                        graphic.DrawImage(original,
                            new Rectangle(0, 0, size.Width, size.Height),
                            0,
                            0,
                            size.Width*Wratio,
                            size.Height*Hratio,
                            GraphicsUnit.Pixel);
                    }
                    if (comboBox2.Text.Equals("Percent"))
                    {
                        var newWidth = (int) (original.Width*(size.Width/100f));
                        var newHeight = (int) (original.Height*(size.Height/100f));
                        float Wratio = original.Width/(float) newWidth;
                        float Hratio = original.Height/(float) newHeight;
                        graphic.DrawImage(original,
                            new Rectangle(0, 0, (int) (original.Width*(size.Width/100f)),
                                (int) (original.Height*((float) size.Height/100))),
                            0,
                            0,
                            (int) (original.Width*(size.Width/100f))*Wratio,
                            (int) (original.Height*(size.Height/100f))*Hratio,
                            GraphicsUnit.Pixel);
                    }
                    graphic.Dispose();
                }
                original.Dispose();
                ms.Dispose();
            }

            //long ram = (System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / 1024) / 1024;

            //while (ram >= 600)
            //{
            //MessageBox.Show("sleeping due to overload: " + ram.ToString() + "MB");
            //   ram = (System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / 1024) / 1024;
            //}
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Thread.SpinWait(5000);
            return img;
        }

        #endregion

        private void button2_Click_2(object sender, EventArgs e)
        {
            foreach (ImageCodecInfo key in ImageCodecInfo.GetImageEncoders())
            {
                MessageBox.Show(key.CodecName);
            }
        }
    }
}