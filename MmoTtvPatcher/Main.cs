using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Security.Cryptography;
using System.Diagnostics;
using Octokit;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;

namespace MmoTtvPatcher
{
    class Application
    {
        Thread notifyThread;
        GitHubClient client;
        string lastTagName = "";
        string asarPath;
        NotifyIcon notify;
        CancellationTokenSource quitApp = new CancellationTokenSource();
        AutoResetEvent notifyReady = new AutoResetEvent(false);

        public Application()
        {
            client = new GitHubClient(new ProductHeaderValue("MmoTtvPatcher", "1.1.0"));
            notifyThread = new Thread(NotifyThreadMain);
            notifyThread.Start();
        }

        public static void Main(string[] args)
        {
            new Application().AppMain();
        }

        private void NotifyThreadMain()
        {
            notify = new NotifyIcon();
            notify.Icon = System.Drawing.SystemIcons.Application;
            notify.Visible = true;
            notify.Text = "MmoTtvPatcher";
            notify.ContextMenu = new ContextMenu();
            notify.ContextMenu.MenuItems.Add(new MenuItem("Quit", (a, b) => quitApp.Cancel()));

            notifyReady.Set();

            while (!quitApp.IsCancellationRequested)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
            }

            notify.Dispose();
        }

        private bool SleepChecked(int milliseconds, int intervalMs)
        {
            for (double time = 0; time < milliseconds;)
            {
                var start = DateTime.Now;

                if (quitApp.IsCancellationRequested)
                {
                    return true;
                }

                Thread.Sleep(intervalMs);
                time += (DateTime.Now - start).TotalMilliseconds;
            }

            return false;
        }

        private void AppMain()
        {
            notifyReady.WaitOne();

            asarPath = LocateMattermostAsar();
            if (asarPath == null)
            {
                notify.ShowBalloonTip(5000, "mmottv", "Couldn't locate where your Mattermost is installed!", ToolTipIcon.Error);
                quitApp.Cancel();
            }

            notify.ShowBalloonTip(5000, "mmottv", "Found your Mattermost. I'm running in the background.", ToolTipIcon.Info);

            while (!quitApp.IsCancellationRequested)
            {
                using (var task = CheckAndInstallLatest())
                {
                    task.Wait();
                }

                // Wait 30 minutes before next version check.
                if (SleepChecked(1000 * 60 * 30, 1000)) break;
            }

            if (notifyThread.IsAlive)
            {
                notifyThread.Join(1000 * 5);
            }
        }

        private async Task CheckAndInstallLatest()
        {
            var release = await client.Repository.Release.GetLatest("Lisoph", "mmottv");
            if (release.TagName == lastTagName) return; // Already installed

            var asar = release.Assets.FirstOrDefault(asset => asset.Name == "app.asar.7z" || asset.Name == "app.asar.zip");
            if (asar == null) return;
            var url = asar.BrowserDownloadUrl;

            if (quitApp.IsCancellationRequested) return;

            var web = new WebClient();
            var downloadPath = Path.GetTempFileName();
            await web.DownloadFileTaskAsync(new Uri(url, UriKind.Absolute), downloadPath);
            downloadPath = Path.Combine(ExtractFile7Zip(downloadPath), "app.asar");

            if (quitApp.IsCancellationRequested) return;
            
            if (!HashFile(downloadPath).SequenceEqual(HashFile(asarPath)))
            {
                string desc = $"Version {release.TagName} is available and ready for installation.";
                desc += "\nUpdate will be applied in 5 minutes.";
                notify.ShowBalloonTip(5000, "mmottv - New Version", desc, ToolTipIcon.Warning);

                if (SleepChecked(1000 * 60 * 4, 500)) return;
                notify.ShowBalloonTip(5000, "mmottv - New Version", "Update will be applied in 1 minute.", ToolTipIcon.Warning);

                if (SleepChecked(1000 * 60 * 1, 500)) return;
                notify.ShowBalloonTip(5000, "mmottv - New Version", "Applying update...", ToolTipIcon.Warning);

                QuitMattermost();
                Thread.Sleep(5000);

                var dt = DateTime.Now;
                string oldFileName = $"{dt.Year}-{dt.Month}-{dt.Day} {dt.Hour}_{dt.Minute}_{dt.Second}.asar";
                File.Move(asarPath, Path.Combine(asarPath, "..\\", oldFileName));
                File.Move(downloadPath, asarPath);

                StartMattermost();

                notify.ShowBalloonTip(5000, "mmottv - New Version", $"Updated mmottv from {lastTagName} to {release.TagName}.", ToolTipIcon.Info);

                lastTagName = release.TagName;
            }
        }

        private string LocateMattermostAsar()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            var mmoRootPath = Path.Combine(appData, "Programs", "mattermost-desktop");
            if (!Directory.Exists(mmoRootPath)) return null;

            var asarPath = Path.Combine(mmoRootPath, "resources", "app.asar");
            if (!File.Exists(asarPath)) return null;

            return asarPath;
        }

        private void QuitMattermost()
        {
            foreach (var proc in Process.GetProcessesByName("Mattermost"))
            {
                try
                {
                    proc.Kill();

                    /*if (proc.MainWindowHandle.ToInt32() != 0)
                    {
                        SetForegroundWindow(proc.MainWindowHandle);
                        uint lparam = 0;
                        SendMessage(proc.MainWindowHandle, WM_KEYDOWN, KEY_LCTRL, lparam);
                        SendMessage(proc.MainWindowHandle, WM_KEYDOWN, KEY_Q, lparam);
                    }*/
                } catch (Exception)
                {

                }
            }
        }

        private void StartMattermost()
        {
            Process.Start(Path.Combine(asarPath, "../../Mattermost.exe"));
        }

        private string ExtractFile7Zip(string filePath)
        {
            var outDir = Path.GetTempPath();
            var info = new ProcessStartInfo("C:\\Program Files\\7-Zip\\7z.exe", $"e \"{filePath}\" -o\"{Path.GetFullPath(outDir)}\" -aoa");
            info.WindowStyle = ProcessWindowStyle.Hidden;
            info.CreateNoWindow = true;
            info.UseShellExecute = false;
            var proc = Process.Start(info);
            proc.WaitForExit();
            return outDir;
        }

        private byte[] HashFile(string filePath)
        {
            using (var hasher = SHA256.Create())
            {
                hasher.Initialize();
                using (var stream = new FileStream(filePath, System.IO.FileMode.Open, FileAccess.Read))
                {
                    return hasher.ComputeHash(stream);
                }
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        static extern void PostQuitMessage(int exitCode);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        static uint WM_CLOSE = 0x10;
        static uint WM_QUIT = 0x12;
        static uint WM_KEYDOWN = 0x100;
        static uint WM_KEYUP = 0x0101;
        static uint KEY_LCTRL = 0xA2;
        static uint KEY_Q = 0x51;
    }
}
