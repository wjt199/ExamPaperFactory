//This file is part of ExamPaper Factory
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ExamPaperFactory
{
    internal static class Program
    {

        #region 判断Office 版本
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]

        static extern uint RegOpenKeyEx(UIntPtr hKey, string lpSubKey, uint ulOptions, int samDesired, out int phkResult);
        [DllImport("Advapi32.dll")]
        static extern uint RegCloseKey(int hKey);
        [DllImport("advapi32.dll", EntryPoint = "RegQueryValueEx")]
        static extern int RegQueryValueEx(int hKey, string lpValueName, int lpReserved, ref uint lpType,
            System.Text.StringBuilder lpData, ref uint lpcbData);
        private static UIntPtr HKEY_LOCAL_MACHINE = new UIntPtr(0x80000002u);
        private static UIntPtr HKEY_CURRENT_USER = new UIntPtr(0x80000001u);

        private static Dictionary<string, string> LatestVersions = new Dictionary<string, string>();

        public static void Office()
        {
            LatestVersions.Add("12.0", "Office2007");
            LatestVersions.Add("14.0", "Office2010");
            LatestVersions.Add("15.0", "Office2013");
            LatestVersions.Add("16.0", "Office2017 and higher");
        }

        private static string GetVersionNumberFromRegistry()
        {
            string regVersion;

            regVersion = GetVersionNumberFromRegistry("SOFTWARE\\Microsoft\\Office\\");
            if (regVersion == null)
                regVersion = GetVersionNumberFromRegistry("SOFTWARE\\Wow6432Node\\Microsoft\\Office\\");

            return regVersion;
        }

        private static string GetVersionNumberFromRegistry(string key)
        {
            string version = null;
            foreach (string VerNo in LatestVersions.Keys)
            {
                string offPath = Reg64(HKEY_LOCAL_MACHINE, key + VerNo + "\\Excel\\InstallRoot", "Path");
                if (offPath != null)
                {
                    version = VerNo;
                    break;
                }
            }
            return version;
        }

        public static string GetVersion()
        {
            Office();
            string versionFromReg = GetVersionNumberFromRegistry();
            string versionInstalled = LatestVersions[versionFromReg];

            bool? Office64BitFromReg = Off64Bit("SOFTWARE\\Microsoft\\Office\\", versionFromReg) ?? Off64Bit("SOFTWARE\\Wow6432Node\\Microsoft\\Office\\", versionFromReg);
            if (Office64BitFromReg.HasValue && Office64BitFromReg.Value) versionInstalled += "(64 bit)";
            else if (Office64BitFromReg.HasValue && !Office64BitFromReg.Value) versionInstalled += "(32 bit)";
            else versionInstalled += "(Unknown bit)";


            return versionInstalled;
        }

        private static bool? Off64Bit(string key, string version)
        {
            bool? Office64BitFromReg = null;
            string Bitness = Reg64(HKEY_LOCAL_MACHINE, key + version + "\\Outlook", "Bitness");
            if (Bitness == "x86")
                Office64BitFromReg = false;
            else if ((Bitness == "x64"))
                Office64BitFromReg = true;
            return Office64BitFromReg;
        }

        private static string Reg64(UIntPtr parent, string key, string prop)
        {
            int ikey = 0;
            int bit36_64 = 0x0100;
            int query = 0x0001;
            try
            {
                uint res = RegOpenKeyEx(HKEY_LOCAL_MACHINE, key, 0, query | bit36_64, out ikey);
                if (0 != res) return null;
                uint type = 0;
                uint data = 1024;
                StringBuilder buffer = new StringBuilder(1024);
                RegQueryValueEx(ikey, prop, 0, ref type, buffer, ref data);
                string ver = buffer.ToString();
                return ver;
            }
            finally
            {
                if (0 != ikey) RegCloseKey(ikey);
            }
        }

        #endregion


        #region Win32 API

        [DllImport("user32.dll")]
        static extern IntPtr GetDC(IntPtr ptr);
        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(
        IntPtr hdc, // handle to DC
        int nIndex // index of capability
        );
        [DllImport("user32.dll", EntryPoint = "ReleaseDC")]
        static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDc);

        #endregion

        #region DeviceCaps - 设备属性 常量

        const int HORZRES = 8;
        const int VERTRES = 10;
        const int LOGPIXELSX = 88;
        const int LOGPIXELSY = 90;
        const int DESKTOPVERTRES = 117;
        const int DESKTOPHORZRES = 118;

        #endregion

        #region 属性

        // 获取屏幕分辨率当前物理大小
        public static Size WorkingArea
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                Size size = new Size();
                size.Width = GetDeviceCaps(hdc, HORZRES);
                size.Height = GetDeviceCaps(hdc, VERTRES);
                ReleaseDC(IntPtr.Zero, hdc);
                return size;
            }
        }
        // 当前系统DPI_X 大小 一般为96
        public static int DpiX
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                int DpiX = GetDeviceCaps(hdc, LOGPIXELSX);
                ReleaseDC(IntPtr.Zero, hdc);
                return DpiX;
            }
        }
        // 当前系统DPI_Y 大小 一般为96
        public static int DpiY
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                int DpiX = GetDeviceCaps(hdc, LOGPIXELSY);
                ReleaseDC(IntPtr.Zero, hdc);
                return DpiX;
            }
        }
        // 获取真实设置的桌面分辨率大小
        public static Size DesktopResolution
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                Size size = new Size();
                size.Width = GetDeviceCaps(hdc, DESKTOPHORZRES);
                size.Height = GetDeviceCaps(hdc, DESKTOPVERTRES);
                ReleaseDC(IntPtr.Zero, hdc);
                return size;
            }
        }
        // 获取宽度缩放百分比
        public static float ScaleX
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                //float ScaleX = (float)GetDeviceCaps(hdc, DESKTOPHORZRES) / (float)GetDeviceCaps(hdc, HORZRES);
                float ScaleX = (float)Program.DpiX / 96f;
                ReleaseDC(IntPtr.Zero, hdc);
                return ScaleX;
            }
        }
        // 获取高度缩放百分比
        public static float ScaleY
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                //float ScaleY = (float)GetDeviceCaps(hdc, DESKTOPVERTRES) / (float)GetDeviceCaps(hdc, VERTRES);
                float ScaleY = (float)Program.DpiY / 96f;
                ReleaseDC(IntPtr.Zero, hdc);
                return ScaleY;
            }
        }

        #endregion

        /// <summary>
        /// 判断.Net Framework的Version是否符合需要
        /// (.Net Framework 版本在2.0及以上)
        /// </summary>
        /// <param name="version">需要的版本 version = 4.5</param>
        /// <returns></returns>
        private static bool GetDotNetVersion(string version)
        {
            string oldname = "0";
            using (RegistryKey ndpKey =
                RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, "").
                OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\"))
            {
                foreach (string versionKeyName in ndpKey.GetSubKeyNames())
                {
                    if (versionKeyName.StartsWith("v"))
                    {
                        RegistryKey versionKey = ndpKey.OpenSubKey(versionKeyName);
                        string newname = (string)versionKey.GetValue("Version", "");
                        if (string.Compare(newname, oldname) > 0)
                        {
                            oldname = newname;
                        }
                        if (newname != "")
                        {
                            continue;
                        }
                        foreach (string subKeyName in versionKey.GetSubKeyNames())
                        {
                            RegistryKey subKey = versionKey.OpenSubKey(subKeyName);
                            newname = (string)subKey.GetValue("Version", "");
                            if (string.Compare(newname, oldname) > 0)
                            {
                                oldname = newname;
                            }
                        }
                    }
                }
            }
            return string.Compare(oldname, version) > 0;
        }

        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();//此方法为应用程序启用可视样式
            /**
             * 作用:在应用程序范围内设置控件显示文本的默认方式(可以设为使用新的GDI+ , 还是旧的GDI)
             *   true使用GDI+方式显示文本，
             *   false使用GDI方式显示文本。
             */
            Application.SetCompatibleTextRenderingDefault(false);

            Type officeType = Type.GetTypeFromProgID("Word.Application");
            Office();
            //检查系统环境是否符合要求，缩放比例是否符合要求
            if (!GetDotNetVersion("4.9.2"))
            {
                if (MessageBox.Show("当前缺少\".Net Framework 4.7.2\"运行环境", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error) == DialogResult.OK)
                {
                    SpecialForm specialForm = new SpecialForm();
                    specialForm.ShowDialog();
                    if (specialForm.IsPassedRight) Application.Run(new MainForm());
                }
            }
            else if (ScaleX != 1f)
            {
                float scaling = 100 * ScaleX; if (MessageBox.Show("当前屏幕缩放比例为" + scaling + "%，请设置为100%后重试！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error) == DialogResult.OK)
                {
                    SpecialForm specialForm = new SpecialForm();
                    specialForm.ShowDialog();
                    if (specialForm.IsPassedRight) Application.Run(new MainForm());
                }
            }
            else if (officeType == null)
            {
                if (MessageBox.Show("当前系统未安装Office软件，请安装Office2013及以上版本！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error) == DialogResult.OK)
                {
                    SpecialForm specialForm = new SpecialForm();
                    specialForm.ShowDialog();
                    if (specialForm.IsPassedRight) Application.Run(new MainForm());
                }
            }
            else if (!(GetVersion().Contains("Office2017 and higher") || GetVersion().Contains("Office2013")))
            {
                if (MessageBox.Show("当前系统安装Office版本较低，请升至Office2013及以上！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error) == DialogResult.OK)
                {
                    SpecialForm specialForm = new SpecialForm();
                    specialForm.ShowDialog();
                    if (specialForm.IsPassedRight) Application.Run(new MainForm());
                }
            }
            else Application.Run(new MainForm());
        }
    }
}
