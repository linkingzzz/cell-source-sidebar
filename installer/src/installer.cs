// 表格批注插件 - 单文件安装/卸载程序
// 安装：双击本 exe；卸载：安装目录下的 uninstall.exe /uninstall
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Win32;

[assembly: AssemblyTitle(C.Product)]
[assembly: AssemblyProduct(C.Product)]
[assembly: AssemblyDescription(C.Product + " 安装程序")]
[assembly: AssemblyFileVersion(C.Version + ".0")]
[assembly: AssemblyVersion(C.Version + ".0")]

internal static class C
{
    public const string Product = "表格批注插件";
    public const string Version = "1.0.0";
    public const string AddinId = "685A162B-4215-4F8B-96BF-63C9D19E8A35";
    public const string CatalogId = "{D8F1B2A3-6C4E-4B77-9E51-2A7C0D5B1E34}";
    public const string WpsName = "cell-source-sidebar";
    public const string WpsFolder = "cell-source-sidebar_1.0.0";
    // WPS 只认 ASCII 的 name_version 目录，中文名会让 WPS 拒绝加载（实测），故内部名保持 ASCII；
    // 中文名曾用于过渡版本，安装/卸载时一并清理。
    public const string WpsNameLegacy = "表格批注插件";
    public const string WpsFolderLegacy = "表格批注插件_1.0.0";
    public const string PayloadResource = "payload.zip";

    public static string InstallDir { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Product); } }
    public static string Manifest { get { return Path.Combine(InstallDir, "manifest.xml"); } }
    public static string Workbook { get { return Path.Combine(InstallDir, Product + ".xlsx"); } }
    public static string UninstallExe { get { return Path.Combine(InstallDir, "uninstall.exe"); } }
    public static string WpsAddins { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"kingsoft\wps\jsaddons"); } }
    public static string WpsPlugin { get { return Path.Combine(WpsAddins, WpsFolder); } }
    public static string Shortcut { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), Product + ".lnk"); } }
    public static string UninstallKey { get { return @"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + Product; } }
}

internal static class Detect
{
    public static string Excel()
    {
        string p = RegValue(Registry.LocalMachine, @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", "");
        if (p != null && File.Exists(p)) { return p; }
        p = RegValue(Registry.CurrentUser, @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", "");
        if (p != null && File.Exists(p)) { return p; }
        RegistryKey[] roots = new RegistryKey[] { Registry.LocalMachine, Registry.CurrentUser };
        for (int i = 0; i < roots.Length; i++)
        {
            string inst = RegValue(roots[i], @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "InstallationPath");
            if (inst != null)
            {
                string guess = Path.Combine(inst, Path.Combine("root", Path.Combine("Office16", "EXCEL.EXE")));
                if (File.Exists(guess)) { return guess; }
            }
        }
        string[] fixedPaths = new string[]
        {
            @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            @"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
        };
        for (int i = 0; i < fixedPaths.Length; i++) { if (File.Exists(fixedPaths[i])) { return fixedPaths[i]; } }
        return null;
    }

    public static bool HasWps()
    {
        if (Directory.Exists(C.WpsAddins)) { return true; }
        string[] keys = new string[]
        {
            @"Software\Kingsoft\Office\6.0\common",
            @"Software\WOW6432Node\Kingsoft\Office\6.0\common"
        };
        RegistryKey[] roots = new RegistryKey[] { Registry.CurrentUser, Registry.LocalMachine };
        for (int r = 0; r < roots.Length; r++)
        {
            for (int i = 0; i < keys.Length; i++)
            {
                string root = RegValue(roots[r], keys[i], "InstallRoot");
                if (root != null)
                {
                    if (File.Exists(Path.Combine(root, Path.Combine("office6", "et.exe")))) { return true; }
                }
            }
        }
        return false;
    }

    public static string RegValue(RegistryKey root, string path, string name)
    {
        try
        {
            using (RegistryKey k = root.OpenSubKey(path))
            {
                if (k == null) { return null; }
                object v = k.GetValue(name);
                return v == null ? null : v.ToString();
            }
        }
        catch { return null; }
    }
}

internal static class Payload
{
    public static string Extract()
    {
        string dir = Path.Combine(Path.GetTempPath(), "biaoge_" + Guid.NewGuid().ToString("N").Substring(0, 8));
        Directory.CreateDirectory(dir);
        using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(C.PayloadResource))
        {
            if (s == null) { throw new Exception("安装包内容缺失（payload.zip）"); }
            using (ZipArchive za = new ZipArchive(s, ZipArchiveMode.Read))
            {
                za.ExtractToDirectory(dir);
            }
        }
        return dir;
    }

    public static void Cleanup(string dir)
    {
        try { if (Directory.Exists(dir)) { Directory.Delete(dir, true); } }
        catch { }
    }
}

internal static class Files2
{
    public static void CopyDir(string src, string dst)
    {
        Directory.CreateDirectory(dst);
        string[] files = Directory.GetFiles(src, "*", SearchOption.AllDirectories);
        for (int i = 0; i < files.Length; i++)
        {
            string rel = files[i].Substring(src.Length).TrimStart('\\', '/');
            string target = Path.Combine(dst, rel);
            Directory.CreateDirectory(Path.GetDirectoryName(target));
            File.Copy(files[i], target, true);
        }
    }
}

internal static class ShellLink
{
    public static void Create(string linkPath, string target, string arguments, string iconPath)
    {
        Type t = Type.GetTypeFromProgID("WScript.Shell");
        if (t == null) { throw new Exception("无法创建快捷方式（WScript.Shell 不可用）"); }
        object shell = Activator.CreateInstance(t);
        object link = t.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, shell, new object[] { linkPath });
        Type lt = link.GetType();
        lt.InvokeMember("TargetPath", BindingFlags.SetProperty, null, link, new object[] { target });
        lt.InvokeMember("Arguments", BindingFlags.SetProperty, null, link, new object[] { arguments });
        lt.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, link, new object[] { Path.GetDirectoryName(target) });
        lt.InvokeMember("IconLocation", BindingFlags.SetProperty, null, link, new object[] { iconPath });
        lt.InvokeMember("Save", BindingFlags.InvokeMethod, null, link, null);
    }
}

internal static class Setup
{
    public static void Run(Action<string> log)
    {
        string tmp = Payload.Extract();
        try
        {
            Directory.CreateDirectory(C.InstallDir);
            File.Copy(Path.Combine(tmp, "manifest.xml"), C.Manifest, true);
            log("已安装加载项清单：" + C.Manifest);

            try
            {
                File.Copy(Path.Combine(tmp, Path.Combine("workbook", "annotation-workbook.xlsx")), C.Workbook, true);
                log("已安装批注工作簿：" + C.Workbook);
            }
            catch (Exception ex)
            {
                log("警告：批注工作簿被占用，未能更新（请关闭 Excel 后重跑安装）：" + ex.Message);
            }

            string excel = Detect.Excel();
            if (excel != null)
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\16.0\WEF\Developer"))
                {
                    k.SetValue(C.AddinId, C.Manifest);
                }
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\" + C.CatalogId))
                {
                    k.SetValue("Url", C.InstallDir);
                    k.SetValue("Flags", 1, RegistryValueKind.DWord);
                }
                log("已注册 Excel 加载项（当前用户，无需管理员权限）。");
                ShellLink.Create(C.Shortcut, excel, "\"" + C.Workbook + "\"", excel);
                log("已创建桌面快捷方式：" + C.Shortcut);
            }
            else
            {
                log("未检测到 Excel（Microsoft Office），跳过 Excel 部分。");
            }

            if (Detect.HasWps())
            {
                Directory.CreateDirectory(C.WpsAddins);
                CleanLegacy(log);
                Files2.CopyDir(Path.Combine(tmp, "wps"), C.WpsPlugin);
                log("已复制 WPS 插件：" + C.WpsPlugin);
                MergePublishXml(log);
            }
            else
            {
                log("未检测到 WPS Office，跳过 WPS 部分。");
            }

            File.Copy(Application.ExecutablePath, C.UninstallExe, true);
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(C.UninstallKey))
            {
                k.SetValue("DisplayName", C.Product);
                k.SetValue("DisplayVersion", C.Version);
                k.SetValue("Publisher", "linkingzzz");
                k.SetValue("InstallLocation", C.InstallDir);
                k.SetValue("DisplayIcon", C.UninstallExe);
                k.SetValue("UninstallString", "\"" + C.UninstallExe + "\" /uninstall");
                k.SetValue("NoModify", 1, RegistryValueKind.DWord);
                k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            }
            log("已写入卸载入口（设置 → 应用 → 已安装的应用）。");
            log("");
            log("安装完成。重启 Excel / WPS 后即可在功能区看到「数据源」选项卡。");
        }
        finally
        {
            Payload.Cleanup(tmp);
        }
    }

    private static void CleanLegacy(Action<string> log)
    {
        string old = Path.Combine(C.WpsAddins, C.WpsFolderLegacy);
        try
        {
            if (Directory.Exists(old)) { Directory.Delete(old, true); log("已清理旧版 WPS 插件目录：" + old); }
        }
        catch (Exception ex) { log("清理旧版 WPS 插件目录失败（请先关闭 WPS）：" + ex.Message); }
        ClearAuthRecord(false, log);
    }

    internal static void ClearAuthRecord(bool includeCurrent, Action<string> log)
    {
        string file = Path.Combine(C.WpsAddins, "authaddin.json");
        try
        {
            if (!File.Exists(file)) { return; }
            string text = File.ReadAllText(file, Encoding.UTF8);
            bool legacy = text.IndexOf(C.WpsNameLegacy, StringComparison.Ordinal) >= 0;
            bool current = includeCurrent && text.IndexOf(C.WpsName, StringComparison.Ordinal) >= 0;
            if (legacy || current)
            {
                File.Delete(file);
                log("已清除 WPS 旧授权记录（WPS 下次启动会自动重建）：" + file);
            }
        }
        catch (Exception ex) { log("清除 WPS 授权记录失败：" + ex.Message); }
    }

    private static void MergePublishXml(Action<string> log)
    {
        string file = Path.Combine(C.WpsAddins, "publish.xml");
        XmlDocument doc = new XmlDocument();
        bool ok = false;
        if (File.Exists(file))
        {
            try
            {
                doc.Load(file);
                ok = doc.DocumentElement != null && doc.DocumentElement.Name == "jsplugins";
            }
            catch { ok = false; }
        }
        if (!ok) { doc.LoadXml("<jsplugins></jsplugins>"); }

        XmlElement node = null;
        List<XmlElement> legacyNodes = new List<XmlElement>();
        foreach (XmlNode n in doc.DocumentElement.ChildNodes)
        {
            XmlElement e = n as XmlElement;
            if (e == null || e.Name != "jsplugin") { continue; }
            string nm = e.GetAttribute("name");
            if (nm == C.WpsName) { node = e; }
            else if (nm == C.WpsNameLegacy) { legacyNodes.Add(e); }
        }
        for (int i = 0; i < legacyNodes.Count; i++) { doc.DocumentElement.RemoveChild(legacyNodes[i]); }
        if (node == null)
        {
            node = doc.CreateElement("jsplugin");
            doc.DocumentElement.AppendChild(node);
        }
        node.SetAttribute("name", C.WpsName);
        node.SetAttribute("type", "et");
        node.SetAttribute("url", C.WpsFolder);
        node.SetAttribute("version", C.Version);
        node.SetAttribute("enable", "enable_dev");
        node.SetAttribute("install", "null");
        node.SetAttribute("customDomain", "");
        doc.Save(file);
        log("已更新 WPS 注册文件：" + file);
    }
}

internal static class Remove
{
    public static void Run(Action<string> log)
    {
        try
        {
            Registry.CurrentUser.DeleteSubKeyTree(C.UninstallKey, false);
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\WEF\Developer", true))
            {
                if (k != null) { k.DeleteValue(C.AddinId, false); }
            }
            Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\" + C.CatalogId, false);
            log("已清除注册表项。");
        }
        catch (Exception ex) { log("清除注册表时出错：" + ex.Message); }

        try
        {
            if (File.Exists(C.Shortcut)) { File.Delete(C.Shortcut); log("已删除桌面快捷方式。"); }
        }
        catch (Exception ex) { log("删除快捷方式失败：" + ex.Message); }

        try
        {
            if (Directory.Exists(C.WpsPlugin)) { Directory.Delete(C.WpsPlugin, true); log("已删除 WPS 插件目录。"); }
            string oldWps = Path.Combine(C.WpsAddins, C.WpsFolderLegacy);
            if (Directory.Exists(oldWps)) { Directory.Delete(oldWps, true); log("已删除旧版 WPS 插件目录。"); }
            Setup.ClearAuthRecord(true, log);
        }
        catch (Exception ex) { log("删除 WPS 插件失败（请先关闭 WPS）：" + ex.Message); }

        RemovePublishEntry(log);

        log("");
        log("卸载完成。");
        Process.Start(new ProcessStartInfo("cmd.exe", "/c for /l %i in (1,1,12) do @(rmdir /s /q \"" + C.InstallDir + "\" 2>nul & if not exist \"" + C.InstallDir + "\" exit /b & ping -n 3 127.0.0.1 >nul)")
        {
            WindowStyle = ProcessWindowStyle.Hidden,
            CreateNoWindow = true
        });
    }

    private static void RemovePublishEntry(Action<string> log)
    {
        string file = Path.Combine(C.WpsAddins, "publish.xml");
        if (!File.Exists(file)) { return; }
        try
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(file);
            List<XmlElement> targets = new List<XmlElement>();
            foreach (XmlNode n in doc.DocumentElement.ChildNodes)
            {
                XmlElement e = n as XmlElement;
                if (e == null || e.Name != "jsplugin") { continue; }
                string nm = e.GetAttribute("name");
                if (nm == C.WpsName || nm == C.WpsNameLegacy) { targets.Add(e); }
            }
            for (int i = 0; i < targets.Count; i++) { doc.DocumentElement.RemoveChild(targets[i]); }
            doc.Save(file);
            log("已从 WPS 注册文件中移除插件条目。");
        }
        catch (Exception ex) { log("更新 WPS 注册文件失败：" + ex.Message); }
    }
}

internal class MainForm : Form
{
    private readonly bool _uninstall;
    private readonly TextBox _log;
    private readonly Button _go;
    private readonly Button _close;
    private readonly CheckBox _open;

    public MainForm(bool uninstall)
    {
        _uninstall = uninstall;
        Text = C.Product + (uninstall ? " 卸载" : " 安装");
        ClientSize = new Size(560, 372);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Microsoft YaHei UI", 9F);

        Label title = new Label();
        title.Left = 14; title.Top = 12; title.Width = 520; title.Height = 22;
        title.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold);
        title.Text = C.Product + "  v" + C.Version + (uninstall ? "  卸载程序" : "  安装程序");

        _log = new TextBox();
        _log.Left = 14; _log.Top = 42; _log.Width = 530; _log.Height = 244;
        _log.Multiline = true; _log.ReadOnly = true; _log.ScrollBars = ScrollBars.Vertical;
        _log.BackColor = Color.White; _log.BorderStyle = BorderStyle.FixedSingle;

        _open = new CheckBox();
        _open.Left = 14; _open.Top = 296; _open.Width = 300;
        _open.Text = "完成后打开批注工作簿（Excel）";
        _open.Checked = true;
        _open.Enabled = !uninstall;

        _go = new Button();
        _go.Left = 288; _go.Top = 322; _go.Width = 122; _go.Height = 32;
        _go.Text = uninstall ? "开始卸载" : "开始安装";
        _go.Click += new EventHandler(OnGo);

        _close = new Button();
        _close.Left = 422; _close.Top = 322; _close.Width = 122; _close.Height = 32;
        _close.Text = "关闭";
        _close.Click += new EventHandler(OnClose);

        FormClosing += delegate(object s, FormClosingEventArgs e) { Trace.Write("界面正在关闭，原因=" + e.CloseReason); };
        Shown += delegate(object s, EventArgs e) { Trace.Write("界面已显示"); };
        Controls.Add(title);
        Controls.Add(_open);
        Controls.Add(_log);
        Controls.Add(_go);
        Controls.Add(_close);

        Append("本程序将 " + C.Product + " 安装到当前用户目录（无需管理员权限）：");
        Append("");
        string excel = Detect.Excel();
        Append(excel == null ? "· Excel：未检测到" : "· Excel：已检测到 " + excel);
        Append(Detect.HasWps() ? "· WPS：已检测到" : "· WPS：未检测到");
        Append("");
        if (!uninstall)
        {
            Append("安装内容：");
            Append("1. Excel 加载项（侧边栏批注：附件 / 明细表 / 备注，页签「数据源」）");
            Append("2. WPS 表格插件（同款功能）");
            Append("3. 一个已内嵌加载项的空白批注工作簿 + 桌面快捷方式");
        }
    }

    private void Append(string s)
    {
        _log.AppendText(s + Environment.NewLine);
        _log.SelectionStart = _log.TextLength;
        _log.ScrollToCaret();
        Application.DoEvents();
    }

    private void OnGo(object sender, EventArgs e)
    {
        _go.Enabled = false;
        Append("");
        try
        {
            if (_uninstall) { Remove.Run(new Action<string>(Append)); }
            else
            {
                Setup.Run(new Action<string>(Append));
                if (_open.Checked)
                {
                    string excel = Detect.Excel();
                    if (excel != null && File.Exists(C.Workbook))
                    {
                        Process.Start(new ProcessStartInfo(excel, "\"" + C.Workbook + "\""));
                        Append("已打开批注工作簿。");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Append("出错：" + ex.Message);
        }
        _go.Enabled = false;
        _close.Text = "完成";
        _close.Focus();
    }

    private void OnClose(object sender, EventArgs e)
    {
        Close();
    }
}

internal static class Trace
{
    public static void Write(string s)
    {
        try
        {
            string dir = C.InstallDir;
            if (!Directory.Exists(dir)) { Directory.CreateDirectory(dir); }
            File.AppendAllText(Path.Combine(dir, "setup.log"),
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + s + Environment.NewLine,
                new UTF8Encoding(false));
        }
        catch { }
    }
}

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        bool uninstall = false;
        bool silent = false;
        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i].TrimStart('-', '/').ToLowerInvariant();
            if (a == "uninstall") { uninstall = true; }
            else if (a == "silent" || a == "quiet") { silent = true; }
        }

        AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e)
        {
            Trace.Write("未处理异常：" + e.ExceptionObject);
        };

        Trace.Write("启动：silent=" + silent + " uninstall=" + uninstall);

        try
        {
            if (silent)
            {
                StringBuilder sb = new StringBuilder();
                Action<string> log = delegate(string s) { sb.AppendLine(s); Trace.Write("  " + s); };
                if (uninstall) { Remove.Run(log); } else { Setup.Run(log); }
                File.WriteAllText(Path.Combine(Path.GetTempPath(), "biaoge-setup.log"), sb.ToString(), new UTF8Encoding(false));
                Trace.Write("静默模式结束");
                Environment.Exit(0);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ThreadException += delegate(object s, System.Threading.ThreadExceptionEventArgs e)
            {
                Trace.Write("UI 线程异常：" + e.Exception);
                MessageBox.Show("程序出错：" + e.Exception.Message, C.Product, MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            Application.Run(new MainForm(uninstall));
            Trace.Write("界面已关闭");
        }
        catch (Exception ex)
        {
            Trace.Write("致命错误：" + ex);
            try { MessageBox.Show(ex.Message, C.Product + " 出错", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
        }
    }
}