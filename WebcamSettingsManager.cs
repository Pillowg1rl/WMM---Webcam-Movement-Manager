using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Net;
using System.Threading;
using System.Windows.Forms;

// ============================================================================
// Webcam Settings Manager - DirectShow Camera Configuration Tool
// Single-file C# WinForms application. Compiles with built-in .NET Framework csc.exe.
// ============================================================================

namespace WebcamSettingsManager
{
    static class AppInfo
    {
        public const string Version = "1.1.1";
        public const string GitHubOwner = "Pillowg1rl";
        public const string GitHubRepo = "WMM---Webcam-Movement-Manager";
    }

    // ========================================================================
    // COM Interface Definitions for DirectShow
    // ========================================================================

    #region DirectShow COM Interfaces

    static class DirectShowGuids
    {
        public static readonly Guid CLSID_SystemDeviceEnum = new Guid("62BE5D10-60EB-11D0-BD3B-00A0C911CE86");
        public static readonly Guid CLSID_VideoInputDeviceCategory = new Guid("860BB310-5D01-11D0-BD3B-00A0C911CE86");
    }

    // VideoProcAmpProperty enum (from strmif.h)
    enum VideoProcAmpProperty
    {
        Brightness = 0,
        Contrast = 1,
        Hue = 2,
        Saturation = 3,
        Sharpness = 4,
        Gamma = 5,
        ColorEnable = 6,
        WhiteBalance = 7,
        BacklightCompensation = 8,
        Gain = 9
    }

    // CameraControlProperty enum (from strmif.h)
    enum CameraControlProperty
    {
        Pan = 0,
        Tilt = 1,
        Roll = 2,
        Zoom = 3,
        Exposure = 4,
        Iris = 5,
        Focus = 6
    }

    // Flags for both VideoProcAmp and CameraControl
    [Flags]
    enum PropertyFlags
    {
        Auto = 0x0001,
        Manual = 0x0002
    }

    [ComImport, Guid("29840822-5B84-11D0-BD3B-00A0C911CE86"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICreateDevEnum
    {
        [PreserveSig]
        int CreateClassEnumerator(
            [In] ref Guid clsidDeviceClass,
            [Out] out IEnumMoniker ppEnumMoniker,
            [In] int dwFlags);
    }

    [ComImport, Guid("C6E13360-30AC-11D0-A18C-00A0C9118956"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAMVideoProcAmp
    {
        [PreserveSig]
        int GetRange(
            [In] VideoProcAmpProperty property,
            [Out] out int pMin,
            [Out] out int pMax,
            [Out] out int pSteppingDelta,
            [Out] out int pDefault,
            [Out] out int pCapsFlags);

        [PreserveSig]
        int Set(
            [In] VideoProcAmpProperty property,
            [In] int lValue,
            [In] int flags);

        [PreserveSig]
        int Get(
            [In] VideoProcAmpProperty property,
            [Out] out int lValue,
            [Out] out int flags);
    }

    [ComImport, Guid("C6E13370-30AC-11D0-A18C-00A0C9118956"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAMCameraControl
    {
        [PreserveSig]
        int GetRange(
            [In] CameraControlProperty property,
            [Out] out int pMin,
            [Out] out int pMax,
            [Out] out int pSteppingDelta,
            [Out] out int pDefault,
            [Out] out int pCapsFlags);

        [PreserveSig]
        int Set(
            [In] CameraControlProperty property,
            [In] int lValue,
            [In] int flags);

        [PreserveSig]
        int Get(
            [In] CameraControlProperty property,
            [Out] out int lValue,
            [Out] out int flags);
    }

    // IPropertyBag - for reading device properties (DevicePath, FriendlyName)
    [ComImport, Guid("55272A00-42CB-11CE-8135-00AA004BB851"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyBag
    {
        [PreserveSig]
        int Read(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
            [In, Out] ref object pVar,
            [In] IntPtr pErrorLog);

        [PreserveSig]
        int Write(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
            [In] ref object pVar);
    }

    #endregion

    // ========================================================================
    // Camera Property Data Structures
    // ========================================================================

    #region Data Structures

    class PropertyValue
    {
        public int Value { get; set; }
        public int Flags { get; set; }
        public bool IsAuto { get { return (Flags & (int)PropertyFlags.Auto) != 0; } }
    }

    class PropertyRange
    {
        public int Min { get; set; }
        public int Max { get; set; }
        public int Step { get; set; }
        public int Default { get; set; }
        public int CapsFlags { get; set; }
        public bool SupportsAuto { get { return (CapsFlags & (int)PropertyFlags.Auto) != 0; } }
        public bool SupportsManual { get { return (CapsFlags & (int)PropertyFlags.Manual) != 0; } }
    }

    class PropertyInfo
    {
        public string Category { get; set; } // "VideoProcAmp" or "CameraControl"
        public string Name { get; set; }
        public int PropertyId { get; set; }
        public PropertyValue Current { get; set; }
        public PropertyRange Range { get; set; }
    }

    // JSON-serializable profile structures
    class ProfileProperty
    {
        public int Value { get; set; }
        public int Flags { get; set; }
    }

    class ProfileDevice
    {
        public string FriendlyName { get; set; }
        public Dictionary<string, ProfileProperty> VideoProcAmp { get; set; }
        public Dictionary<string, ProfileProperty> CameraControl { get; set; }

        public ProfileDevice()
        {
            VideoProcAmp = new Dictionary<string, ProfileProperty>();
            CameraControl = new Dictionary<string, ProfileProperty>();
        }
    }

    class Profile
    {
        public string Name { get; set; }
        public string Created { get; set; }
        public string Notes { get; set; }
        public Dictionary<string, ProfileDevice> Devices { get; set; }

        public Profile()
        {
            Notes = "";
            Devices = new Dictionary<string, ProfileDevice>();
        }
    }

    #endregion

    // ========================================================================
    // AppSettings - persisted user preferences (dark mode, etc.)
    // ========================================================================

    #region AppSettings

    class HotkeyBinding
    {
        public int Modifiers { get; set; } // Win32 modifier flags
        public int Key { get; set; }       // Win32 virtual key code

        public string DisplayName
        {
            get
            {
                var parts = new List<string>();
                if ((Modifiers & 0x0002) != 0) parts.Add("Ctrl");
                if ((Modifiers & 0x0001) != 0) parts.Add("Alt");
                if ((Modifiers & 0x0004) != 0) parts.Add("Shift");
                if ((Modifiers & 0x0008) != 0) parts.Add("Win");
                parts.Add(((Keys)Key).ToString());
                return string.Join("+", parts);
            }
        }
    }

    class AppSettings
    {
        public bool DarkMode { get; set; }
        public Dictionary<string, HotkeyBinding> Hotkeys { get; set; }

        private static string SettingsPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "WebcamSettings", "settings.json");
            }
        }

        public AppSettings()
        {
            DarkMode = false;
            Hotkeys = new Dictionary<string, HotkeyBinding>();
        }

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    string json = File.ReadAllText(SettingsPath);
                    var s = new AppSettings();
                    if (json.Contains("\"DarkMode\": true") || json.Contains("\"DarkMode\":true"))
                        s.DarkMode = true;

                    // Parse hotkeys
                    int hkIdx = json.IndexOf("\"Hotkeys\"");
                    if (hkIdx >= 0)
                    {
                        int braceStart = json.IndexOf('{', hkIdx + 9);
                        if (braceStart >= 0)
                        {
                            int depth = 1;
                            int braceEnd = braceStart + 1;
                            while (braceEnd < json.Length && depth > 0)
                            {
                                if (json[braceEnd] == '{') depth++;
                                else if (json[braceEnd] == '}') depth--;
                                braceEnd++;
                            }
                            string hkJson = json.Substring(braceStart, braceEnd - braceStart);
                            ParseHotkeys(s, hkJson);
                        }
                    }
                    return s;
                }
            }
            catch { }
            return new AppSettings();
        }

        private static void ParseHotkeys(AppSettings s, string json)
        {
            // Simple parser for: { "name": {"Modifiers": N, "Key": N}, ... }
            int pos = 1; // skip opening {
            while (pos < json.Length)
            {
                int nameStart = json.IndexOf('"', pos);
                if (nameStart < 0) break;
                int nameEnd = json.IndexOf('"', nameStart + 1);
                if (nameEnd < 0) break;
                string name = json.Substring(nameStart + 1, nameEnd - nameStart - 1);

                int objStart = json.IndexOf('{', nameEnd);
                if (objStart < 0) break;
                int objEnd = json.IndexOf('}', objStart);
                if (objEnd < 0) break;
                string objStr = json.Substring(objStart, objEnd - objStart + 1);

                var hk = new HotkeyBinding();
                int modIdx = objStr.IndexOf("\"Modifiers\"");
                int keyIdx = objStr.IndexOf("\"Key\"");
                if (modIdx >= 0)
                {
                    int colonIdx = objStr.IndexOf(':', modIdx);
                    if (colonIdx >= 0)
                    {
                        string numStr = "";
                        for (int i = colonIdx + 1; i < objStr.Length; i++)
                        {
                            char c = objStr[i];
                            if (char.IsDigit(c) || c == '-') numStr += c;
                            else if (numStr.Length > 0) break;
                        }
                        int val; if (int.TryParse(numStr, out val)) hk.Modifiers = val;
                    }
                }
                if (keyIdx >= 0)
                {
                    int colonIdx = objStr.IndexOf(':', keyIdx);
                    if (colonIdx >= 0)
                    {
                        string numStr = "";
                        for (int i = colonIdx + 1; i < objStr.Length; i++)
                        {
                            char c = objStr[i];
                            if (char.IsDigit(c) || c == '-') numStr += c;
                            else if (numStr.Length > 0) break;
                        }
                        int val; if (int.TryParse(numStr, out val)) hk.Key = val;
                    }
                }

                if (hk.Key != 0)
                    s.Hotkeys[name] = hk;

                pos = objEnd + 1;
            }
        }

        public void Save()
        {
            try
            {
                string dir = Path.GetDirectoryName(SettingsPath);
                Directory.CreateDirectory(dir);
                var sb = new StringBuilder();
                sb.AppendLine("{");
                sb.AppendLine("  \"DarkMode\": " + (DarkMode ? "true" : "false") + ",");
                sb.AppendLine("  \"Hotkeys\": {");
                int idx = 0;
                foreach (var kvp in Hotkeys)
                {
                    string comma = idx < Hotkeys.Count - 1 ? "," : "";
                    sb.AppendLine("    \"" + kvp.Key.Replace("\"", "\\\"") + "\": {\"Modifiers\": " + kvp.Value.Modifiers + ", \"Key\": " + kvp.Value.Key + "}" + comma);
                    idx++;
                }
                sb.AppendLine("  }");
                sb.AppendLine("}");
                File.WriteAllText(SettingsPath, sb.ToString());
            }
            catch { }
        }
    }

    #endregion

    // ========================================================================
    // DarkModeHelper - apply dark/light theme to all controls
    // ========================================================================

    #region DarkModeHelper

    static class DarkModeHelper
    {
        // Dark theme colors
        static readonly Color DarkBg = Color.FromArgb(30, 30, 30);
        static readonly Color DarkControl = Color.FromArgb(45, 45, 45);
        static readonly Color DarkText = Color.FromArgb(220, 220, 220);
        static readonly Color DarkBorder = Color.FromArgb(60, 60, 60);
        static readonly Color DarkGroupPTZ = Color.FromArgb(55, 45, 30);
        static readonly Color DarkGroupCam = Color.FromArgb(35, 50, 35);
        static readonly Color DarkGroupVideo = Color.FromArgb(35, 35, 55);

        // Light theme colors
        static readonly Color LightGroupPTZ = Color.FromArgb(255, 245, 230);
        static readonly Color LightGroupCam = Color.FromArgb(240, 250, 240);
        static readonly Color LightGroupVideo = Color.FromArgb(240, 240, 250);

        public static Color GetGroupColor(string groupTitle, bool dark)
        {
            if (groupTitle == "PTZ Controls") return dark ? DarkGroupPTZ : LightGroupPTZ;
            if (groupTitle == "Camera Control") return dark ? DarkGroupCam : LightGroupCam;
            return dark ? DarkGroupVideo : LightGroupVideo;
        }

        public static void ApplyTheme(Control root, bool dark)
        {
            ApplyToControl(root, dark);
            foreach (Control c in root.Controls)
                ApplyTheme(c, dark);
        }

        private static void ApplyToControl(Control c, bool dark)
        {
            if (c is Form)
            {
                c.BackColor = dark ? DarkBg : SystemColors.Control;
                c.ForeColor = dark ? DarkText : SystemColors.ControlText;
            }
            else if (c is MenuStrip || c is ToolStrip || c is StatusStrip)
            {
                c.BackColor = dark ? DarkControl : SystemColors.Control;
                c.ForeColor = dark ? DarkText : SystemColors.ControlText;
            }
            else if (c is GroupBox)
            {
                var gb = (GroupBox)c;
                gb.ForeColor = dark ? DarkText : SystemColors.ControlText;
                gb.BackColor = GetGroupColor(gb.Text, dark);
            }
            else if (c is TabControl)
            {
                c.BackColor = dark ? DarkBg : SystemColors.Control;
                c.ForeColor = dark ? DarkText : SystemColors.ControlText;
            }
            else if (c is TabPage)
            {
                c.BackColor = dark ? DarkBg : SystemColors.Control;
                c.ForeColor = dark ? DarkText : SystemColors.ControlText;
            }
            else if (c is TextBox || c is NumericUpDown || c is ComboBox || c is RichTextBox)
            {
                c.BackColor = dark ? Color.FromArgb(55, 55, 55) : SystemColors.Window;
                c.ForeColor = dark ? DarkText : SystemColors.WindowText;
            }
            else if (c is Button)
            {
                c.BackColor = dark ? DarkControl : SystemColors.Control;
                c.ForeColor = dark ? DarkText : SystemColors.ControlText;
            }
            else if (c is CheckBox || c is Label)
            {
                c.ForeColor = dark ? DarkText : SystemColors.ControlText;
            }
            else if (c is Panel)
            {
                if (!(c.BackColor != Color.Transparent && c.Parent is GroupBox))
                    c.BackColor = dark ? DarkBg : SystemColors.Control;
                c.ForeColor = dark ? DarkText : SystemColors.ControlText;
            }
            else if (c is TrackBar)
            {
                c.BackColor = dark ? DarkControl : SystemColors.Control;
            }
            else if (c is PictureBox)
            {
                c.BackColor = dark ? Color.FromArgb(20, 20, 20) : Color.FromArgb(200, 200, 200);
            }
        }
    }

    #endregion


    // ========================================================================
    // CameraDevice - wraps a single DirectShow video input device
    // ========================================================================

    #region CameraDevice

    class CameraDevice : IDisposable
    {
        public string DevicePath { get; private set; }
        public string FriendlyName { get; private set; }

        private IMoniker _moniker;
        private IAMVideoProcAmp _procAmp;
        private IAMCameraControl _camCtrl;
        private bool _hasProcAmp;
        private bool _hasCamCtrl;
        private bool _disposed;

        public CameraDevice(string devicePath, string friendlyName, IMoniker moniker)
        {
            DevicePath = devicePath;
            FriendlyName = friendlyName;
            _moniker = moniker;
            BindInterfaces();
        }

        private void BindInterfaces()
        {
            Guid iidProcAmp = typeof(IAMVideoProcAmp).GUID;
            Guid iidCamCtrl = typeof(IAMCameraControl).GUID;
            object obj;

            try
            {
                _moniker.BindToObject(null, null, ref iidProcAmp, out obj);
                _procAmp = (IAMVideoProcAmp)obj;
                _hasProcAmp = true;
            }
            catch { _hasProcAmp = false; }

            try
            {
                _moniker.BindToObject(null, null, ref iidCamCtrl, out obj);
                _camCtrl = (IAMCameraControl)obj;
                _hasCamCtrl = true;
            }
            catch { _hasCamCtrl = false; }
        }

        public PropertyValue GetVideoProcAmpProperty(VideoProcAmpProperty prop)
        {
            if (!_hasProcAmp) return null;
            int val, flags;
            int hr = _procAmp.Get(prop, out val, out flags);
            if (hr != 0) return null;
            return new PropertyValue { Value = val, Flags = flags };
        }

        public PropertyRange GetVideoProcAmpRange(VideoProcAmpProperty prop)
        {
            if (!_hasProcAmp) return null;
            int min, max, step, def, caps;
            int hr = _procAmp.GetRange(prop, out min, out max, out step, out def, out caps);
            if (hr != 0) return null;
            return new PropertyRange { Min = min, Max = max, Step = step, Default = def, CapsFlags = caps };
        }

        public bool SetVideoProcAmpProperty(VideoProcAmpProperty prop, int value, int flags)
        {
            if (!_hasProcAmp) return false;
            return _procAmp.Set(prop, value, flags) == 0;
        }

        public PropertyValue GetCameraControlProperty(CameraControlProperty prop)
        {
            if (!_hasCamCtrl) return null;
            int val, flags;
            int hr = _camCtrl.Get(prop, out val, out flags);
            if (hr != 0) return null;
            return new PropertyValue { Value = val, Flags = flags };
        }

        public PropertyRange GetCameraControlRange(CameraControlProperty prop)
        {
            if (!_hasCamCtrl) return null;
            int min, max, step, def, caps;
            int hr = _camCtrl.GetRange(prop, out min, out max, out step, out def, out caps);
            if (hr != 0) return null;
            return new PropertyRange { Min = min, Max = max, Step = step, Default = def, CapsFlags = caps };
        }

        public bool SetCameraControlProperty(CameraControlProperty prop, int value, int flags)
        {
            if (!_hasCamCtrl) return false;
            return _camCtrl.Set(prop, value, flags) == 0;
        }

        // PTZ cameras have slow motors. Send the absolute target value repeatedly
        // until the motor settles at the correct position.
        public bool SetCameraControlPropertyVerified(CameraControlProperty prop, int targetValue, int flags,
            int maxAttempts, int delayMs, out int finalValue)
        {
            finalValue = targetValue;
            if (!_hasCamCtrl) return false;

            var range = GetCameraControlRange(prop);
            int tolerance = (range != null && range.Step > 0) ? range.Step : 1;

            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                // Always send the absolute target value
                _camCtrl.Set(prop, targetValue, flags);

                // Wait for motor to move
                Thread.Sleep(delayMs);

                // Read back actual position
                int currentVal, currentFlags;
                int hr = _camCtrl.Get(prop, out currentVal, out currentFlags);
                if (hr != 0) return false;

                finalValue = currentVal;

                // Close enough?
                if (Math.Abs(currentVal - targetValue) <= tolerance)
                    return true;
            }

            return false;
        }

        public List<PropertyInfo> GetAllProperties()
        {
            var props = new List<PropertyInfo>();

            foreach (VideoProcAmpProperty p in Enum.GetValues(typeof(VideoProcAmpProperty)))
            {
                var range = GetVideoProcAmpRange(p);
                var current = GetVideoProcAmpProperty(p);
                if (range != null && current != null)
                {
                    props.Add(new PropertyInfo
                    {
                        Category = "VideoProcAmp",
                        Name = p.ToString(),
                        PropertyId = (int)p,
                        Current = current,
                        Range = range
                    });
                }
            }

            foreach (CameraControlProperty p in Enum.GetValues(typeof(CameraControlProperty)))
            {
                var range = GetCameraControlRange(p);
                var current = GetCameraControlProperty(p);
                if (range != null && current != null)
                {
                    // Override Roll range — cameras often report too small a range
                    if (p == CameraControlProperty.Roll)
                    {
                        range.Min = -100;
                        range.Max = 100;
                    }

                    props.Add(new PropertyInfo
                    {
                        Category = "CameraControl",
                        Name = p.ToString(),
                        PropertyId = (int)p,
                        Current = current,
                        Range = range
                    });
                }
            }

            return props;
        }

        public List<string> ApplyProperties(ProfileDevice profileDev)
        {
            var errors = new List<string>();

            foreach (var kvp in profileDev.VideoProcAmp)
            {
                VideoProcAmpProperty prop;
                if (Enum.TryParse(kvp.Key, out prop))
                {
                    if (!SetVideoProcAmpProperty(prop, kvp.Value.Value, kvp.Value.Flags))
                        errors.Add("Failed to set VideoProcAmp." + kvp.Key);
                }
            }

            foreach (var kvp in profileDev.CameraControl)
            {
                CameraControlProperty prop;
                if (Enum.TryParse(kvp.Key, out prop))
                {
                    int finalVal;
                    bool ok = SetCameraControlPropertyVerified(prop, kvp.Value.Value, kvp.Value.Flags,
                        20, 300, out finalVal);
                    if (!ok)
                        errors.Add("CameraControl." + kvp.Key + ": target=" + kvp.Value.Value + " actual=" + finalVal);
                }
            }

            return errors;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (_procAmp != null) { try { Marshal.ReleaseComObject(_procAmp); } catch { } }
            if (_camCtrl != null) { try { Marshal.ReleaseComObject(_camCtrl); } catch { } }
            if (_moniker != null) { try { Marshal.ReleaseComObject(_moniker); } catch { } }
        }
    }

    #endregion

    // ========================================================================
    // DeviceEnumerator - enumerate all DirectShow video input devices
    // ========================================================================

    #region DeviceEnumerator

    static class DeviceEnumerator
    {
        public static List<CameraDevice> EnumerateDevices()
        {
            var devices = new List<CameraDevice>();

            object devEnumObj = null;
            ICreateDevEnum devEnum = null;
            IEnumMoniker enumMoniker = null;

            try
            {
                Type sdeType = Type.GetTypeFromCLSID(DirectShowGuids.CLSID_SystemDeviceEnum);
                devEnumObj = Activator.CreateInstance(sdeType);
                devEnum = (ICreateDevEnum)devEnumObj;

                Guid category = DirectShowGuids.CLSID_VideoInputDeviceCategory;
                int hr = devEnum.CreateClassEnumerator(ref category, out enumMoniker, 0);
                if (hr != 0 || enumMoniker == null)
                    return devices;

                IMoniker[] monikers = new IMoniker[1];
                IntPtr fetched = IntPtr.Zero;

                while (enumMoniker.Next(1, monikers, fetched) == 0)
                {
                    IMoniker moniker = monikers[0];
                    string devicePath = "";
                    string friendlyName = "";

                    IPropertyBag propBag = null;
                    try
                    {
                        Guid iidPropBag = typeof(IPropertyBag).GUID;
                        object propBagObj;
                        moniker.BindToStorage(null, null, ref iidPropBag, out propBagObj);
                        propBag = (IPropertyBag)propBagObj;

                        object val = null;
                        if (propBag.Read("DevicePath", ref val, IntPtr.Zero) == 0 && val != null)
                        {
                            devicePath = val.ToString();
                            val = null;
                        }
                        if (propBag.Read("FriendlyName", ref val, IntPtr.Zero) == 0 && val != null)
                        {
                            friendlyName = val.ToString();
                            val = null;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (propBag != null) Marshal.ReleaseComObject(propBag);
                    }

                    if (!string.IsNullOrEmpty(devicePath))
                    {
                        devices.Add(new CameraDevice(devicePath, friendlyName, moniker));
                    }
                    else
                    {
                        Marshal.ReleaseComObject(moniker);
                    }
                }
            }
            catch { }
            finally
            {
                if (enumMoniker != null) Marshal.ReleaseComObject(enumMoniker);
                if (devEnumObj != null) Marshal.ReleaseComObject(devEnumObj);
            }

            return devices;
        }
    }

    #endregion

    // ========================================================================
    // ProfileManager - save/load JSON profiles
    // ========================================================================

    #region ProfileManager

    class ProfileManager
    {
        public string ProfileDir { get; private set; }

        public ProfileManager()
        {
            ProfileDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "WebcamSettings", "profiles");
            Directory.CreateDirectory(ProfileDir);
        }

        public List<string> ListProfiles()
        {
            var profiles = new List<string>();
            if (Directory.Exists(ProfileDir))
            {
                foreach (var f in Directory.GetFiles(ProfileDir, "*.json"))
                    profiles.Add(Path.GetFileNameWithoutExtension(f));
            }
            profiles.Sort();
            return profiles;
        }

        public void SaveProfile(string name, List<CameraDevice> devices, string devicePathFilter = null, string notes = null)
        {
            var profile = new Profile
            {
                Name = name,
                Created = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                Notes = notes ?? ""
            };

            foreach (var cam in devices)
            {
                if (devicePathFilter != null && cam.DevicePath != devicePathFilter)
                    continue;

                var profileDev = new ProfileDevice { FriendlyName = cam.FriendlyName };
                var allProps = cam.GetAllProperties();

                foreach (var prop in allProps)
                {
                    var pp = new ProfileProperty { Value = prop.Current.Value, Flags = prop.Current.Flags };
                    if (prop.Category == "VideoProcAmp")
                        profileDev.VideoProcAmp[prop.Name] = pp;
                    else
                        profileDev.CameraControl[prop.Name] = pp;
                }

                profile.Devices[cam.DevicePath] = profileDev;
            }

            string json = SimpleJson.Serialize(profile);
            File.WriteAllText(Path.Combine(ProfileDir, name + ".json"), json, Encoding.UTF8);
        }

        public Profile LoadProfile(string name)
        {
            string path = Path.Combine(ProfileDir, name + ".json");
            if (!File.Exists(path)) return null;
            string json = File.ReadAllText(path, Encoding.UTF8);
            return SimpleJson.Deserialize(json);
        }

        public List<string> ApplyProfile(string name, List<CameraDevice> devices, string devicePathFilter = null)
        {
            var errors = new List<string>();
            var profile = LoadProfile(name);
            if (profile == null) { errors.Add("Profile not found: " + name); return errors; }

            foreach (var cam in devices)
            {
                if (devicePathFilter != null && cam.DevicePath != devicePathFilter)
                    continue;

                ProfileDevice profileDev;
                if (profile.Devices.TryGetValue(cam.DevicePath, out profileDev))
                {
                    errors.AddRange(cam.ApplyProperties(profileDev));
                }
            }

            return errors;
        }

        public void DeleteProfile(string name)
        {
            string path = Path.Combine(ProfileDir, name + ".json");
            if (File.Exists(path)) File.Delete(path);
        }

        public void ExportProfile(string name, string destPath)
        {
            string src = Path.Combine(ProfileDir, name + ".json");
            if (File.Exists(src)) File.Copy(src, destPath, true);
        }

        public string ImportProfile(string srcPath)
        {
            string name = Path.GetFileNameWithoutExtension(srcPath);
            string dest = Path.Combine(ProfileDir, name + ".json");
            File.Copy(srcPath, dest, true);
            return name;
        }
    }

    #endregion

    // ========================================================================
    // SimpleJson - minimal JSON serializer/deserializer (no external deps)
    // ========================================================================

    #region SimpleJson

    static class SimpleJson
    {
        public static string Serialize(Profile profile)
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine("  \"Name\": " + JsonStr(profile.Name) + ",");
            sb.AppendLine("  \"Created\": " + JsonStr(profile.Created) + ",");
            sb.AppendLine("  \"Notes\": " + JsonStr(profile.Notes) + ",");
            sb.AppendLine("  \"Devices\": {");

            int devIdx = 0;
            foreach (var kvp in profile.Devices)
            {
                sb.AppendLine("    " + JsonStr(kvp.Key) + ": {");
                sb.AppendLine("      \"FriendlyName\": " + JsonStr(kvp.Value.FriendlyName) + ",");

                sb.AppendLine("      \"VideoProcAmp\": {");
                SerializeProps(sb, kvp.Value.VideoProcAmp);
                sb.AppendLine("      },");

                sb.AppendLine("      \"CameraControl\": {");
                SerializeProps(sb, kvp.Value.CameraControl);
                sb.AppendLine("      }");

                sb.AppendLine("    }" + (devIdx < profile.Devices.Count - 1 ? "," : ""));
                devIdx++;
            }

            sb.AppendLine("  }");
            sb.AppendLine("}");
            return sb.ToString();
        }

        private static void SerializeProps(StringBuilder sb, Dictionary<string, ProfileProperty> props)
        {
            int idx = 0;
            foreach (var kvp in props)
            {
                string comma = idx < props.Count - 1 ? "," : "";
                sb.AppendLine("        " + JsonStr(kvp.Key) + ": {\"Value\": " + kvp.Value.Value + ", \"Flags\": " + kvp.Value.Flags + "}" + comma);
                idx++;
            }
        }

        private static string JsonStr(string s)
        {
            if (s == null) return "null";
            return "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r") + "\"";
        }

        public static Profile Deserialize(string json)
        {
            var profile = new Profile();
            var tokens = Tokenize(json);
            int pos = 0;
            var obj = ParseObject(tokens, ref pos);

            if (obj.ContainsKey("Name")) profile.Name = obj["Name"] as string;
            if (obj.ContainsKey("Created")) profile.Created = obj["Created"] as string;
            if (obj.ContainsKey("Notes")) profile.Notes = (obj["Notes"] as string) ?? "";

            if (obj.ContainsKey("Devices"))
            {
                var devicesObj = obj["Devices"] as Dictionary<string, object>;
                if (devicesObj != null)
                {
                    foreach (var devKvp in devicesObj)
                    {
                        var devObj = devKvp.Value as Dictionary<string, object>;
                        if (devObj == null) continue;

                        var profileDev = new ProfileDevice();
                        if (devObj.ContainsKey("FriendlyName"))
                            profileDev.FriendlyName = devObj["FriendlyName"] as string;

                        ParsePropCategory(devObj, "VideoProcAmp", profileDev.VideoProcAmp);
                        ParsePropCategory(devObj, "CameraControl", profileDev.CameraControl);

                        profile.Devices[devKvp.Key] = profileDev;
                    }
                }
            }

            return profile;
        }

        private static void ParsePropCategory(Dictionary<string, object> devObj, string category,
            Dictionary<string, ProfileProperty> target)
        {
            if (!devObj.ContainsKey(category)) return;
            var catObj = devObj[category] as Dictionary<string, object>;
            if (catObj == null) return;

            foreach (var kvp in catObj)
            {
                var propObj = kvp.Value as Dictionary<string, object>;
                if (propObj == null) continue;
                var pp = new ProfileProperty();
                if (propObj.ContainsKey("Value")) pp.Value = ToInt(propObj["Value"]);
                if (propObj.ContainsKey("Flags")) pp.Flags = ToInt(propObj["Flags"]);
                target[kvp.Key] = pp;
            }
        }

        private static int ToInt(object o)
        {
            if (o is int) return (int)o;
            if (o is long) return (int)(long)o;
            if (o is double) return (int)(double)o;
            if (o is string) { int v; int.TryParse((string)o, out v); return v; }
            return 0;
        }

        // Minimal JSON tokenizer
        enum TokenType { LBrace, RBrace, LBracket, RBracket, Colon, Comma, StringVal, NumberVal, BoolVal, NullVal }

        class Token
        {
            public TokenType Type;
            public string Value;
        }

        static List<Token> Tokenize(string json)
        {
            var tokens = new List<Token>();
            int i = 0;
            while (i < json.Length)
            {
                char c = json[i];
                if (char.IsWhiteSpace(c)) { i++; continue; }

                if (c == '{') { tokens.Add(new Token { Type = TokenType.LBrace }); i++; }
                else if (c == '}') { tokens.Add(new Token { Type = TokenType.RBrace }); i++; }
                else if (c == '[') { tokens.Add(new Token { Type = TokenType.LBracket }); i++; }
                else if (c == ']') { tokens.Add(new Token { Type = TokenType.RBracket }); i++; }
                else if (c == ':') { tokens.Add(new Token { Type = TokenType.Colon }); i++; }
                else if (c == ',') { tokens.Add(new Token { Type = TokenType.Comma }); i++; }
                else if (c == '"')
                {
                    i++;
                    var sb = new StringBuilder();
                    while (i < json.Length)
                    {
                        if (json[i] == '\\' && i + 1 < json.Length)
                        {
                            char next = json[i + 1];
                            if (next == '"') { sb.Append('"'); i += 2; }
                            else if (next == '\\') { sb.Append('\\'); i += 2; }
                            else if (next == 'n') { sb.Append('\n'); i += 2; }
                            else if (next == 'r') { sb.Append('\r'); i += 2; }
                            else if (next == 't') { sb.Append('\t'); i += 2; }
                            else { sb.Append(next); i += 2; }
                        }
                        else if (json[i] == '"') { i++; break; }
                        else { sb.Append(json[i]); i++; }
                    }
                    tokens.Add(new Token { Type = TokenType.StringVal, Value = sb.ToString() });
                }
                else if (c == '-' || char.IsDigit(c))
                {
                    int start = i;
                    if (c == '-') i++;
                    while (i < json.Length && (char.IsDigit(json[i]) || json[i] == '.')) i++;
                    tokens.Add(new Token { Type = TokenType.NumberVal, Value = json.Substring(start, i - start) });
                }
                else if (json.Substring(i).StartsWith("true")) { tokens.Add(new Token { Type = TokenType.BoolVal, Value = "true" }); i += 4; }
                else if (json.Substring(i).StartsWith("false")) { tokens.Add(new Token { Type = TokenType.BoolVal, Value = "false" }); i += 5; }
                else if (json.Substring(i).StartsWith("null")) { tokens.Add(new Token { Type = TokenType.NullVal, Value = "null" }); i += 4; }
                else { i++; }
            }
            return tokens;
        }

        static object ParseValue(List<Token> tokens, ref int pos)
        {
            if (pos >= tokens.Count) return null;
            var t = tokens[pos];
            switch (t.Type)
            {
                case TokenType.LBrace: return ParseObject(tokens, ref pos);
                case TokenType.LBracket: return ParseArray(tokens, ref pos);
                case TokenType.StringVal: pos++; return t.Value;
                case TokenType.NumberVal:
                    pos++;
                    if (t.Value.Contains(".")) { double d; double.TryParse(t.Value, out d); return d; }
                    else { long l; long.TryParse(t.Value, out l); return (int)l; }
                case TokenType.BoolVal: pos++; return t.Value == "true";
                case TokenType.NullVal: pos++; return null;
                default: pos++; return null;
            }
        }

        static Dictionary<string, object> ParseObject(List<Token> tokens, ref int pos)
        {
            var obj = new Dictionary<string, object>();
            pos++; // skip {
            while (pos < tokens.Count && tokens[pos].Type != TokenType.RBrace)
            {
                if (tokens[pos].Type == TokenType.Comma) { pos++; continue; }
                string key = tokens[pos].Value; pos++; // key
                if (pos < tokens.Count && tokens[pos].Type == TokenType.Colon) pos++; // :
                obj[key] = ParseValue(tokens, ref pos);
            }
            if (pos < tokens.Count) pos++; // skip }
            return obj;
        }

        static List<object> ParseArray(List<Token> tokens, ref int pos)
        {
            var arr = new List<object>();
            pos++; // skip [
            while (pos < tokens.Count && tokens[pos].Type != TokenType.RBracket)
            {
                if (tokens[pos].Type == TokenType.Comma) { pos++; continue; }
                arr.Add(ParseValue(tokens, ref pos));
            }
            if (pos < tokens.Count) pos++; // skip ]
            return arr;
        }
    }

    #endregion

    // ========================================================================
    // JoystickControl - 2D Pan/Tilt joystick widget
    // ========================================================================

    #region JoystickControl

    // Velocity-style joystick: push to move, snap back to center on release.
    // The further you push, the faster the camera moves in that direction.
    class JoystickControl : Panel
    {
        private int _panMin, _panMax, _tiltMin, _tiltMax;
        private bool _dragging;
        private const int DotSize = 24;

        // Dot position in pixels relative to control's top-left
        private int _dotX, _dotY;
        // Center of control
        private int CenterX { get { return Width / 2; } }
        private int CenterY { get { return Height / 2; } }
        // Max distance from center
        private int MaxRadius { get { return Math.Min(Width, Height) / 2 - DotSize / 2 - 4; } }

        // Direction vector in [-1, 1] (set during drag, zero when idle)
        private double _dirX, _dirY;

        private System.Windows.Forms.Timer _moveTimer;
        // Each tick we send EITHER a pan delta OR a tilt delta (alternating)
        // to avoid the slow PTZ motor cancellation issue
        private bool _alternateTickIsPan;

        // Tunable parameters (exposed for live testing)
        // Speed: at full deflection, move this fraction of the range per tick
        public double MaxSpeedFraction { get; set; }
        public int DiagonalMinGapMs { get; set; }
        public int TickIntervalMs
        {
            get { return _moveTimer != null ? _moveTimer.Interval : 50; }
            set { if (_moveTimer != null && value >= 10) _moveTimer.Interval = value; }
        }
        private DateTime _lastFireTime = DateTime.MinValue;

        // Fired each tick during drag with deltas to apply (panDelta, tiltDelta)
        // Either panDelta or tiltDelta is non-zero per call (alternating)
        public event Action<int, int> MoveStep;

        public JoystickControl(int panMin, int panMax, int tiltMin, int tiltMax)
        {
            _panMin = panMin; _panMax = panMax;
            _tiltMin = tiltMin; _tiltMax = tiltMax;

            Width = 180;
            Height = 180;
            BackColor = Color.FromArgb(245, 235, 220);
            BorderStyle = BorderStyle.FixedSingle;
            DoubleBuffered = true;

            _dotX = CenterX - DotSize / 2;
            _dotY = CenterY - DotSize / 2;

            MaxSpeedFraction = 0.020; // 2.0% per tick
            DiagonalMinGapMs = 20;

            _moveTimer = new System.Windows.Forms.Timer { Interval = 40 };
            _moveTimer.Tick += OnMoveTick;

            MouseDown += OnMouseDownHandler;
            MouseMove += OnMouseMoveHandler;
            MouseUp += OnMouseUpHandler;
        }

        // No-op for compatibility (joystick is velocity-based, not position-based)
        public void SetPosition(int pan, int tilt) { }

        private void OnMouseDownHandler(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            _dragging = true;
            Capture = true;
            UpdateDotFromMouse(e.X, e.Y);
            _moveTimer.Start();
        }

        private void OnMouseMoveHandler(object sender, MouseEventArgs e)
        {
            if (!_dragging) return;
            UpdateDotFromMouse(e.X, e.Y);
        }

        private void OnMouseUpHandler(object sender, MouseEventArgs e)
        {
            if (!_dragging) return;
            _dragging = false;
            Capture = false;
            _moveTimer.Stop();

            // Snap dot back to center
            _dirX = 0;
            _dirY = 0;
            _dotX = CenterX - DotSize / 2;
            _dotY = CenterY - DotSize / 2;
            Invalidate();
        }

        private void UpdateDotFromMouse(int mx, int my)
        {
            // Vector from center to mouse
            double dx = mx - CenterX;
            double dy = my - CenterY;

            // CARDINAL-ONLY MODE: snap to dominant axis (no diagonal movement)
            // because diagonal jiggles the slow PTZ motors
            if (Math.Abs(dx) >= Math.Abs(dy))
                dy = 0;
            else
                dx = 0;

            // Clamp to max radius
            double dist = Math.Sqrt(dx * dx + dy * dy);
            if (dist > MaxRadius)
            {
                double scale = MaxRadius / dist;
                dx *= scale;
                dy *= scale;
            }

            // Normalize direction to [-1, 1]
            _dirX = dx / MaxRadius;
            _dirY = -dy / MaxRadius; // invert Y (screen Y grows downward, but tilt+ = up)

            // Update dot position visually
            _dotX = CenterX + (int)dx - DotSize / 2;
            _dotY = CenterY + (int)dy - DotSize / 2;
            Invalidate();
        }

        private void OnMoveTick(object sender, EventArgs e)
        {
            if (_dirX == 0 && _dirY == 0) return;
            if (MoveStep == null) return;

            // Detect diagonal movement (both axes have meaningful deflection)
            bool isDiagonal = Math.Abs(_dirX) > 0.1 && Math.Abs(_dirY) > 0.1;

            // Throttle diagonal movement so consecutive Pan/Tilt commands have
            // a gap big enough for the motor to settle between them
            if (isDiagonal && (DateTime.Now - _lastFireTime).TotalMilliseconds < DiagonalMinGapMs)
                return;

            int panDelta = 0, tiltDelta = 0;

            // Apply non-linear curve so small deflections give precise control
            double curveX = _dirX * Math.Abs(_dirX);
            double curveY = _dirY * Math.Abs(_dirY);

            if (_alternateTickIsPan)
            {
                panDelta = (int)Math.Round((_panMax - _panMin) * MaxSpeedFraction * curveX);
            }
            else
            {
                tiltDelta = (int)Math.Round((_tiltMax - _tiltMin) * MaxSpeedFraction * curveY);
            }

            _alternateTickIsPan = !_alternateTickIsPan;

            if (panDelta != 0 || tiltDelta != 0)
            {
                _lastFireTime = DateTime.Now;
                MoveStep(panDelta, tiltDelta);
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            int w = Width - 1;
            int h = Height - 1;
            int cx = CenterX;
            int cy = CenterY;
            int r = MaxRadius;

            // Outer ring (boundary)
            using (var pen = new Pen(Color.FromArgb(180, 160, 140), 2))
                g.DrawEllipse(pen, cx - r, cy - r, r * 2, r * 2);

            // Crosshair through center
            using (var pen = new Pen(Color.FromArgb(200, 180, 160)))
            {
                g.DrawLine(pen, cx - r, cy, cx + r, cy);
                g.DrawLine(pen, cx, cy - r, cx, cy + r);
            }

            // Center mark
            using (var brush = new SolidBrush(Color.FromArgb(160, 140, 120)))
                g.FillEllipse(brush, cx - 3, cy - 3, 6, 6);

            // Line from center to dot (when dragging)
            if (_dragging)
            {
                using (var pen = new Pen(Color.FromArgb(220, 100, 50), 2))
                    g.DrawLine(pen, cx, cy, _dotX + DotSize / 2, _dotY + DotSize / 2);
            }

            // Joystick dot
            using (var brush = new SolidBrush(_dragging ? Color.FromArgb(220, 100, 50) : Color.FromArgb(180, 130, 90)))
                g.FillEllipse(brush, _dotX, _dotY, DotSize, DotSize);
            using (var pen = new Pen(Color.FromArgb(140, 60, 30), 2))
                g.DrawEllipse(pen, _dotX, _dotY, DotSize, DotSize);

            // Direction labels
            using (var font = new Font(Font.FontFamily, 8, FontStyle.Bold))
            using (var brush = new SolidBrush(Color.FromArgb(140, 100, 80)))
            {
                g.DrawString("Up", font, brush, cx - 8, 2);
                g.DrawString("Dn", font, brush, cx - 8, h - 14);
                g.DrawString("L", font, brush, 4, cy - 7);
                g.DrawString("R", font, brush, w - 12, cy - 7);
            }
        }
    }

    #endregion

    // ========================================================================
    // PropertyRowControl - reusable slider row for one camera property
    // ========================================================================

    #region PropertyRowControl

    class PropertyRowControl : Panel
    {
        private Label _label;
        private TrackBar _slider;
        private NumericUpDown _valueBox;
        private CheckBox _autoCheck;
        private Button _defaultBtn;

        private PropertyInfo _propInfo;
        private bool _suppressEvents;

        public event Action<PropertyRowControl> PropertyChanged;

        public string Category { get { return _propInfo.Category; } }
        public string PropertyName { get { return _propInfo.Name; } }
        public int PropertyId { get { return _propInfo.PropertyId; } }
        public int CurrentValue { get { return (int)_valueBox.Value; } }
        public int CurrentFlags { get { return _autoCheck.Checked ? (int)PropertyFlags.Auto : (int)PropertyFlags.Manual; } }

        public PropertyRowControl(PropertyInfo propInfo)
        {
            _propInfo = propInfo;
            Height = 32;
            Dock = DockStyle.Top;
            BuildControls();
            LoadValues();
        }

        private void BuildControls()
        {
            _label = new Label
            {
                Text = _propInfo.Name,
                Width = 150,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Left
            };

            _defaultBtn = new Button
            {
                Text = "Def",
                Width = 40,
                Dock = DockStyle.Right,
                FlatStyle = FlatStyle.Flat
            };
            _defaultBtn.Click += (s, e) => ResetToDefault();

            _autoCheck = new CheckBox
            {
                Text = "Auto",
                Width = 55,
                Dock = DockStyle.Right,
                Checked = _propInfo.Current.IsAuto
            };
            _autoCheck.CheckedChanged += OnAutoChanged;

            _valueBox = new NumericUpDown
            {
                Width = 70,
                Dock = DockStyle.Right,
                Minimum = _propInfo.Range.Min,
                Maximum = _propInfo.Range.Max,
                Increment = Math.Max(1, _propInfo.Range.Step),
                DecimalPlaces = 0
            };
            _valueBox.ValueChanged += OnValueBoxChanged;

            _slider = new TrackBar
            {
                Minimum = _propInfo.Range.Min,
                Maximum = _propInfo.Range.Max,
                SmallChange = Math.Max(1, _propInfo.Range.Step),
                LargeChange = Math.Max(1, _propInfo.Range.Step * 5),
                TickFrequency = Math.Max(1, (_propInfo.Range.Max - _propInfo.Range.Min) / 10),
                Dock = DockStyle.Fill
            };
            _slider.ValueChanged += OnSliderChanged;

            // Add in reverse order due to Dock layout
            Controls.Add(_slider);
            Controls.Add(_label);
            Controls.Add(_defaultBtn);
            Controls.Add(_autoCheck);
            Controls.Add(_valueBox);

            // Disable auto check if not supported
            if (!_propInfo.Range.SupportsAuto)
            {
                _autoCheck.Enabled = false;
                _autoCheck.Checked = false;
            }
        }

        private void LoadValues()
        {
            _suppressEvents = true;
            int val = Clamp(_propInfo.Current.Value, _propInfo.Range.Min, _propInfo.Range.Max);
            _slider.Value = val;
            _valueBox.Value = val;
            _autoCheck.Checked = _propInfo.Current.IsAuto;
            UpdateEnabledState();
            _suppressEvents = false;
        }

        public void UpdateFromDevice(PropertyValue pv)
        {
            if (pv == null) return;
            _suppressEvents = true;
            int val = Clamp(pv.Value, _propInfo.Range.Min, _propInfo.Range.Max);
            _slider.Value = val;
            _valueBox.Value = val;
            _autoCheck.Checked = pv.IsAuto;
            UpdateEnabledState();
            _suppressEvents = false;
        }

        // Set value externally and trigger PropertyChanged (used by joystick)
        public void SetValueAndApply(int value)
        {
            int val = Clamp(value, _propInfo.Range.Min, _propInfo.Range.Max);
            _suppressEvents = true;
            _slider.Value = val;
            _valueBox.Value = val;
            _suppressEvents = false;
            if (PropertyChanged != null) PropertyChanged(this);
        }

        public int Min { get { return _propInfo.Range.Min; } }
        public int Max { get { return _propInfo.Range.Max; } }

        private void OnSliderChanged(object sender, EventArgs e)
        {
            if (_suppressEvents) return;
            _suppressEvents = true;
            _valueBox.Value = _slider.Value;
            _suppressEvents = false;
            if (PropertyChanged != null) PropertyChanged(this);
        }

        private void OnValueBoxChanged(object sender, EventArgs e)
        {
            if (_suppressEvents) return;
            _suppressEvents = true;
            int val = Clamp((int)_valueBox.Value, _slider.Minimum, _slider.Maximum);
            _slider.Value = val;
            _suppressEvents = false;
            if (PropertyChanged != null) PropertyChanged(this);
        }

        private void OnAutoChanged(object sender, EventArgs e)
        {
            UpdateEnabledState();
            if (!_suppressEvents && PropertyChanged != null)
                PropertyChanged(this);
        }

        private void UpdateEnabledState()
        {
            bool manual = !_autoCheck.Checked;
            _slider.Enabled = manual;
            _valueBox.Enabled = manual;
            _defaultBtn.Enabled = manual;
        }

        private void ResetToDefault()
        {
            _suppressEvents = true;
            int val = Clamp(_propInfo.Range.Default, _propInfo.Range.Min, _propInfo.Range.Max);
            _slider.Value = val;
            _valueBox.Value = val;
            _suppressEvents = false;
            if (PropertyChanged != null) PropertyChanged(this);
        }

        private int Clamp(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }
    }

    #endregion

    // ========================================================================
    // CameraTabPage - one tab per camera with grouped property controls
    // ========================================================================

    #region CameraTabPage

    class CameraTabPage : TabPage
    {
        private CameraDevice _camera;
        private List<PropertyRowControl> _rows = new List<PropertyRowControl>();
        private Action<string> _statusCallback;
        private Panel _scrollPanel;

        public CameraDevice Camera { get { return _camera; } }

        public CameraTabPage(CameraDevice camera, Action<string> statusCallback)
        {
            _camera = camera;
            _statusCallback = statusCallback;
            Text = camera.FriendlyName;
            BuildUI();
        }

        private void BuildUI()
        {
            _scrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(5)
            };
            Controls.Add(_scrollPanel);

            var allProps = _camera.GetAllProperties();

            // Create groups in reverse order (because Dock=Top stacks top-down)
            var videoProcProps = allProps.Where(p => p.Category == "VideoProcAmp").ToList();
            var cameraCtrlNonPTZ = allProps.Where(p =>
                p.Category == "CameraControl" &&
                p.Name != "Pan" && p.Name != "Tilt" && p.Name != "Zoom").ToList();
            var ptzProps = allProps.Where(p =>
                p.Category == "CameraControl" &&
                (p.Name == "Pan" || p.Name == "Tilt" || p.Name == "Zoom")).ToList();

            // Add in reverse order for Dock=Top stacking
            AddGroup("Video Processing", videoProcProps, Color.FromArgb(240, 240, 250));
            AddGroup("Camera Control", cameraCtrlNonPTZ, Color.FromArgb(240, 250, 240));
            AddGroup("PTZ Controls", ptzProps, Color.FromArgb(255, 245, 230));

            // Add joystick to the PTZ group if Pan AND Tilt are present
            AddJoystickToPTZ(ptzProps);
        }

        private void AddJoystickToPTZ(List<PropertyInfo> ptzProps)
        {
            var panInfo = ptzProps.FirstOrDefault(p => p.Name == "Pan");
            var tiltInfo = ptzProps.FirstOrDefault(p => p.Name == "Tilt");
            if (panInfo == null || tiltInfo == null) return;

            // Find the PTZ groupbox we just added (last GroupBox in scrollPanel)
            GroupBox ptzGroup = null;
            foreach (Control c in _scrollPanel.Controls)
            {
                if (c is GroupBox && ((GroupBox)c).Text == "PTZ Controls")
                {
                    ptzGroup = (GroupBox)c;
                    break;
                }
            }
            if (ptzGroup == null) return;

            var panRow = _rows.FirstOrDefault(r => r.PropertyName == "Pan");
            var tiltRow = _rows.FirstOrDefault(r => r.PropertyName == "Tilt");
            if (panRow == null || tiltRow == null) return;

            // Container panel for the joystick (centered)
            var joyPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 200,
                BackColor = Color.Transparent
            };

            var joystick = new JoystickControl(panRow.Min, panRow.Max, tiltRow.Min, tiltRow.Max);
            joystick.Anchor = AnchorStyles.Top;

            joystick.MoveStep += (panDelta, tiltDelta) =>
            {
                if (panDelta != 0)
                    panRow.SetValueAndApply(panRow.CurrentValue + panDelta);
                if (tiltDelta != 0)
                    tiltRow.SetValueAndApply(tiltRow.CurrentValue + tiltDelta);
            };

            joyPanel.Controls.Add(joystick);

            // Center horizontally
            joyPanel.Resize += (s, e) =>
                joystick.Left = (joyPanel.Width - joystick.Width) / 2;
            joystick.Top = 5;
            joystick.Left = (joyPanel.Width - joystick.Width) / 2;

            // Insert at the top of the PTZ group's inner panel
            if (ptzGroup.Controls.Count > 0)
            {
                var innerPanel = ptzGroup.Controls[0] as Panel;
                if (innerPanel != null)
                    innerPanel.Controls.Add(joyPanel);
            }

            _ptzJoystick = joystick;
            _ptzPanRow = panRow;
            _ptzTiltRow = tiltRow;
        }

        private JoystickControl _ptzJoystick;
        private PropertyRowControl _ptzPanRow;
        private PropertyRowControl _ptzTiltRow;

        private void AddGroup(string title, List<PropertyInfo> props, Color bgColor)
        {
            if (props.Count == 0) return;

            var group = new GroupBox
            {
                Text = title,
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(5, 15, 5, 5),
                BackColor = bgColor
            };

            var innerPanel = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true
            };

            // Add rows in reverse so Dock=Top stacks correctly
            for (int i = props.Count - 1; i >= 0; i--)
            {
                var row = new PropertyRowControl(props[i]);
                row.PropertyChanged += OnPropertyChanged;
                _rows.Add(row);
                innerPanel.Controls.Add(row);
            }

            group.Controls.Add(innerPanel);
            _scrollPanel.Controls.Add(group);

            // Add spacing
            var spacer = new Panel { Height = 5, Dock = DockStyle.Top };
            _scrollPanel.Controls.Add(spacer);
        }

        private void OnPropertyChanged(PropertyRowControl row)
        {
            bool success;
            if (row.Category == "VideoProcAmp")
            {
                VideoProcAmpProperty prop = (VideoProcAmpProperty)row.PropertyId;
                success = _camera.SetVideoProcAmpProperty(prop, row.CurrentValue, row.CurrentFlags);
            }
            else
            {
                CameraControlProperty prop = (CameraControlProperty)row.PropertyId;
                success = _camera.SetCameraControlProperty(prop, row.CurrentValue, row.CurrentFlags);
            }

            _statusCallback(success
                ? "Set " + row.PropertyName + " = " + row.CurrentValue + (row.CurrentFlags == (int)PropertyFlags.Auto ? " [Auto]" : " [Manual]")
                : "Failed to set " + row.PropertyName);
        }

        public void RefreshFromDevice()
        {
            foreach (var row in _rows)
            {
                PropertyValue pv;
                if (row.Category == "VideoProcAmp")
                    pv = _camera.GetVideoProcAmpProperty((VideoProcAmpProperty)row.PropertyId);
                else
                    pv = _camera.GetCameraControlProperty((CameraControlProperty)row.PropertyId);

                row.UpdateFromDevice(pv);
            }

            // Sync joystick position with current Pan/Tilt
            if (_ptzJoystick != null && _ptzPanRow != null && _ptzTiltRow != null)
                _ptzJoystick.SetPosition(_ptzPanRow.CurrentValue, _ptzTiltRow.CurrentValue);
        }
    }

    #endregion

    // ========================================================================
    // MainForm - main application window
    // ========================================================================

    #region MainForm

    class MainForm : Form
    {
        private TabControl _tabControl;
        private ToolStrip _toolbar;
        private StatusStrip _statusBar;
        private ToolStripStatusLabel _statusLabel;
        private ComboBox _profileCombo;
        private MenuStrip _menuStrip;
        private ProfileManager _profileMgr;
        private List<CameraDevice> _cameras = new List<CameraDevice>();
        private List<CameraTabPage> _cameraTabs = new List<CameraTabPage>();

        // Dark mode
        private AppSettings _settings;
        private ToolStripMenuItem _darkModeItem;

        // Profile notes
        private Label _notesLabel;
        private Panel _profileInfoPanel;
        private bool _suppressProfileChange;

        // Global hotkeys
        private List<string> _hotkeyProfiles = new List<string>();
        const int WM_HOTKEY = 0x0312;
        const int HOTKEY_BASE_ID = 9000;

        [DllImport("user32.dll")]
        static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
        [DllImport("user32.dll")]
        static extern bool UnregisterHotKey(IntPtr hWnd, int id);


        public MainForm()
        {
            Text = "Webcam Settings Manager";
            Size = new Size(900, 700);
            MinimumSize = new Size(700, 500);
            StartPosition = FormStartPosition.CenterScreen;

            _settings = AppSettings.Load();
            _profileMgr = new ProfileManager();

            BuildMenu();
            BuildToolbar();
            BuildProfileInfoPanel();
            BuildTabControl();
            BuildStatusBar();
            EnumerateAndPopulate();

            if (_settings.DarkMode)
            {
                _darkModeItem.Checked = true;
                DarkModeHelper.ApplyTheme(this, true);
            }
        }

        private void BuildMenu()
        {
            _menuStrip = new MenuStrip();

            var fileMenu = new ToolStripMenuItem("File");
            fileMenu.DropDownItems.Add("Import Profile...", null, (s, e) => ImportProfile());
            fileMenu.DropDownItems.Add("Export Profile...", null, (s, e) => ExportProfile());
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add("Open Profiles Folder", null, (s, e) => OpenProfilesFolder());
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add("Exit", null, (s, e) => Close());

            var viewMenu = new ToolStripMenuItem("View");
            _darkModeItem = new ToolStripMenuItem("Dark Mode");
            _darkModeItem.CheckOnClick = true;
            _darkModeItem.Click += (s, e) => ToggleDarkMode();
            viewMenu.DropDownItems.Add(_darkModeItem);

            var devMenu = new ToolStripMenuItem("Devices");
            devMenu.DropDownItems.Add("Refresh Devices", null, (s, e) => RefreshDevices());

            var helpMenu = new ToolStripMenuItem("Help");
            helpMenu.DropDownItems.Add("Check for Updates", null, (s, e) => CheckForUpdates());
            helpMenu.DropDownItems.Add(new ToolStripSeparator());
            helpMenu.DropDownItems.Add("Hotkeys", null, (s, e) => ShowHotkeyHelp());
            helpMenu.DropDownItems.Add(new ToolStripSeparator());
            helpMenu.DropDownItems.Add("About", null, (s, e) =>
                MessageBox.Show("Webcam Settings Manager v" + AppInfo.Version + "\n\nSave and restore DirectShow webcam settings.\nSupports PTZ preset positions.\n\nBuilt with .NET Framework WinForms.",
                    "About", MessageBoxButtons.OK, MessageBoxIcon.Information));

            _menuStrip.Items.Add(fileMenu);
            _menuStrip.Items.Add(viewMenu);
            _menuStrip.Items.Add(devMenu);
            _menuStrip.Items.Add(helpMenu);
            MainMenuStrip = _menuStrip;
            Controls.Add(_menuStrip);
        }

        private void BuildToolbar()
        {
            _toolbar = new ToolStrip();
            _toolbar.GripStyle = ToolStripGripStyle.Hidden;
            _toolbar.RenderMode = ToolStripRenderMode.System;

            _toolbar.Items.Add(new ToolStripButton("Refresh", null, (s, e) => RefreshDevices()) { ToolTipText = "Re-enumerate cameras" });
            _toolbar.Items.Add(new ToolStripSeparator());
            _toolbar.Items.Add(new ToolStripButton("Save", null, (s, e) => SaveProfile(false)) { ToolTipText = "Save current settings to profile" });
            _toolbar.Items.Add(new ToolStripSeparator());

            _profileCombo = new ComboBox
            {
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _profileCombo.SelectedIndexChanged += (s, e) => OnProfileSelected();
            _toolbar.Items.Add(new ToolStripControlHost(_profileCombo));

            _toolbar.Items.Add(new ToolStripButton("Delete", null, (s, e) => DeleteProfile()) { ToolTipText = "Delete selected profile" });
            _toolbar.Items.Add(new ToolStripSeparator());
            _toolbar.Items.Add(new ToolStripButton("Generate .bat", null, (s, e) => GenerateBat()) { ToolTipText = "Generate a .bat file to restore this profile" });
            _toolbar.Items.Add(new ToolStripButton("Set Hotkey", null, (s, e) => SetHotkeyForProfile()) { ToolTipText = "Assign a global hotkey to restore this profile" });
            Controls.Add(_toolbar);
        }

        private void BuildProfileInfoPanel()
        {
            _profileInfoPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 180,
                Padding = new Padding(5)
            };

            _notesLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.TopLeft,
                Text = "",
                Padding = new Padding(2)
            };

            _profileInfoPanel.Controls.Add(_notesLabel);

            var titleLabel = new Label
            {
                Text = "Profile Notes",
                Dock = DockStyle.Top,
                Height = 20,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font(Font.FontFamily, 9, FontStyle.Bold)
            };
            _profileInfoPanel.Controls.Add(titleLabel);

            Controls.Add(_profileInfoPanel);

        }

        private void BuildTabControl()
        {
            _tabControl = new TabControl
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(_tabControl);

            // Ensure proper z-order (tab control fills remaining space)
            _tabControl.BringToFront();
        }

        private void BuildStatusBar()
        {
            _statusBar = new StatusStrip();
            _statusLabel = new ToolStripStatusLabel("Ready");
            _statusLabel.Spring = true;
            _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            _statusBar.Items.Add(_statusLabel);
            Controls.Add(_statusBar);
        }

        private void SetStatus(string msg)
        {
            _statusLabel.Text = msg;
        }

        private void EnumerateAndPopulate()
        {
            // Clean up existing
            foreach (var tab in _cameraTabs)
                _tabControl.TabPages.Remove(tab);
            _cameraTabs.Clear();

            foreach (var cam in _cameras)
                cam.Dispose();
            _cameras.Clear();

            // Enumerate
            try
            {
                _cameras = DeviceEnumerator.EnumerateDevices();
            }
            catch (Exception ex)
            {
                SetStatus("Error enumerating devices: " + ex.Message);
                return;
            }

            if (_cameras.Count == 0)
            {
                var noDevTab = new TabPage("No cameras found");
                var lbl = new Label
                {
                    Text = "No DirectShow video input devices detected.\n\nPlug in a webcam and click 'Refresh'.",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font(Font.FontFamily, 12)
                };
                noDevTab.Controls.Add(lbl);
                _tabControl.TabPages.Add(noDevTab);
                SetStatus("No cameras found");
                return;
            }

            foreach (var cam in _cameras)
            {
                var tab = new CameraTabPage(cam, SetStatus);
                _cameraTabs.Add(tab);
                _tabControl.TabPages.Add(tab);
            }

            SetStatus("Found " + _cameras.Count + " camera(s)");
            RefreshProfileList();
        }

        private void RefreshDevices()
        {
            EnumerateAndPopulate();
        }

        private void RefreshProfileList()
        {
            _suppressProfileChange = true;
            _profileCombo.Items.Clear();
            foreach (var name in _profileMgr.ListProfiles())
                _profileCombo.Items.Add(name);
            if (_profileCombo.Items.Count > 0)
                _profileCombo.SelectedIndex = 0;
            _suppressProfileChange = false;
            UpdateProfileInfo();
        }

        private void SaveProfile(bool allDevices)
        {
            if (_cameras.Count == 0)
            {
                MessageBox.Show("No cameras available.", "Save Profile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string name = null;
            string notes = "";

            // If a profile is selected, offer to overwrite it
            if (_profileCombo.SelectedItem != null)
            {
                string existing = _profileCombo.SelectedItem.ToString();
                var result = MessageBox.Show(
                    "Overwrite existing profile '" + existing + "'?\n\nYes = overwrite, No = save as new",
                    "Save Profile", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result == DialogResult.Cancel) return;
                if (result == DialogResult.Yes)
                {
                    name = existing;
                    var existingProfile = _profileMgr.LoadProfile(name);
                    if (existingProfile != null) notes = existingProfile.Notes ?? "";
                }
            }

            if (name == null)
            {
                name = PromptProfileName(out notes);
                if (string.IsNullOrEmpty(name)) return;
            }

            try
            {
                _profileMgr.SaveProfile(name, _cameras, null, notes);
                _suppressProfileChange = true;
                RefreshProfileList();
                int idx = _profileCombo.Items.IndexOf(name);
                if (idx >= 0) _profileCombo.SelectedIndex = idx;
                _suppressProfileChange = false;
                RegisterProfileHotkeys();

                SetStatus("Saved profile: " + name);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving profile: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteProfile()
        {
            if (_profileCombo.SelectedItem == null)
            {
                MessageBox.Show("Please select a profile first.", "Delete Profile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string name = _profileCombo.SelectedItem.ToString();
            if (MessageBox.Show("Delete profile '" + name + "'?", "Confirm Delete",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _profileMgr.DeleteProfile(name);
                RefreshProfileList();
                RegisterProfileHotkeys();
                SetStatus("Deleted profile: " + name);
            }
        }

        private void GenerateBat()
        {
            if (_profileCombo.SelectedItem == null)
            {
                MessageBox.Show("Please select a profile first.", "Generate .bat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string name = _profileCombo.SelectedItem.ToString();
            using (var dlg = new SaveFileDialog())
            {
                dlg.Filter = "Batch file|*.bat";
                dlg.FileName = "restore_" + name + ".bat";
                dlg.Title = "Save .bat file";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    string exePath = Application.ExecutablePath;
                    string content = "@echo off\r\n\"" + exePath + "\" --restore \"" + name + "\"\r\npause\r\n";
                    File.WriteAllText(dlg.FileName, content, Encoding.Default);
                    SetStatus("Generated: " + dlg.FileName);
                }
            }
        }

        private void CheckForUpdates()
        {
            SetStatus("Checking for updates...");
            Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            try
            {
                string apiUrl = "https://api.github.com/repos/" + AppInfo.GitHubOwner + "/" + AppInfo.GitHubRepo + "/releases/latest";

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                var request = (HttpWebRequest)WebRequest.Create(apiUrl);
                request.UserAgent = "WebcamSettingsManager/" + AppInfo.Version;
                request.Accept = "application/json";
                request.Timeout = 10000;

                string json;
                using (var response = (HttpWebResponse)request.GetResponse())
                using (var reader = new StreamReader(response.GetResponseStream()))
                {
                    json = reader.ReadToEnd();
                }

                // Parse tag_name, body, and browser_download_url for .exe
                string tagName = ExtractJsonString(json, "tag_name");
                string body = ExtractJsonString(json, "body");
                string htmlUrl = ExtractJsonString(json, "html_url");

                // Find .exe download URL from assets
                string exeUrl = null;
                string exeName = null;
                int assetsIdx = json.IndexOf("\"assets\"");
                if (assetsIdx > 0)
                {
                    string assetsSection = json.Substring(assetsIdx);
                    // Find all browser_download_url entries
                    int searchFrom = 0;
                    while (true)
                    {
                        int dlIdx = assetsSection.IndexOf("\"browser_download_url\"", searchFrom);
                        if (dlIdx < 0) break;
                        string url = ExtractJsonString(assetsSection.Substring(dlIdx - 1), "browser_download_url");
                        if (url != null && url.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        {
                            exeUrl = url;
                            exeName = url.Substring(url.LastIndexOf('/') + 1);
                            break;
                        }
                        searchFrom = dlIdx + 20;
                    }
                }

                Cursor = Cursors.Default;

                // Clean version strings for comparison
                string remoteVer = (tagName ?? "").TrimStart('v', 'V');
                string localVer = AppInfo.Version;

                if (remoteVer == localVer)
                {
                    SetStatus("You are on the latest version (v" + localVer + ").");
                    MessageBox.Show("You are already on the latest version.\n\nInstalled: v" + localVer,
                        "No Update Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Show release notes and ask user
                string message = "A new version is available!\n\n"
                    + "Installed: v" + localVer + "\n"
                    + "Latest: " + (tagName ?? "unknown") + "\n\n"
                    + "--- Release Notes ---\n"
                    + (body ?? "(no release notes)") + "\n\n";

                if (exeUrl != null)
                    message += "Do you want to download and install the update?";
                else
                    message += "No .exe found in this release. Visit the releases page?";

                var result = MessageBox.Show(message, "Update Available",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result != DialogResult.Yes) return;

                if (exeUrl != null)
                {
                    // Download the new exe
                    SetStatus("Downloading update...");
                    Cursor = Cursors.WaitCursor;
                    Application.DoEvents();

                    string currentExe = Application.ExecutablePath;
                    string newExe = currentExe + ".update";
                    string oldExe = currentExe + ".old";

                    using (var wc = new WebClient())
                    {
                        wc.Headers.Add("User-Agent", "WebcamSettingsManager/" + AppInfo.Version);
                        wc.DownloadFile(exeUrl, newExe);
                    }

                    Cursor = Cursors.Default;

                    // Create a batch script that waits for this process to exit,
                    // swaps the files, and launches the new version
                    string batPath = Path.Combine(Path.GetTempPath(), "wmm_update.bat");
                    string bat = "@echo off\r\n"
                        + "echo Updating Webcam Settings Manager...\r\n"
                        + "timeout /t 2 /nobreak >nul\r\n"
                        + "if exist \"" + oldExe + "\" del \"" + oldExe + "\"\r\n"
                        + "move \"" + currentExe + "\" \"" + oldExe + "\"\r\n"
                        + "move \"" + newExe + "\" \"" + currentExe + "\"\r\n"
                        + "start \"\" \"" + currentExe + "\"\r\n"
                        + "del \"" + oldExe + "\"\r\n"
                        + "del \"%~f0\"\r\n";
                    File.WriteAllText(batPath, bat);

                    // Launch the updater and exit
                    var psi = new System.Diagnostics.ProcessStartInfo();
                    psi.FileName = batPath;
                    psi.CreateNoWindow = true;
                    psi.UseShellExecute = false;
                    System.Diagnostics.Process.Start(psi);
                    Application.Exit();
                }
                else
                {
                    // Open releases page in browser
                    System.Diagnostics.Process.Start(htmlUrl ?? ("https://github.com/" + AppInfo.GitHubOwner + "/" + AppInfo.GitHubRepo + "/releases"));
                }
            }
            catch (WebException wex)
            {
                Cursor = Cursors.Default;
                var httpResp = wex.Response as HttpWebResponse;
                if (httpResp != null && httpResp.StatusCode == HttpStatusCode.NotFound)
                {
                    SetStatus("No releases found.");
                    MessageBox.Show("No releases found on GitHub yet.", "Check for Updates", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    SetStatus("Update check failed.");
                    MessageBox.Show("Could not check for updates:\n" + wex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                SetStatus("Update check failed.");
                MessageBox.Show("Could not check for updates:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Simple helper to extract a string value from JSON by key
        private string ExtractJsonString(string json, string key)
        {
            string search = "\"" + key + "\"";
            int idx = json.IndexOf(search);
            if (idx < 0) return null;
            int colonIdx = json.IndexOf(':', idx + search.Length);
            if (colonIdx < 0) return null;
            // Skip whitespace after colon
            int start = colonIdx + 1;
            while (start < json.Length && (json[start] == ' ' || json[start] == '\t' || json[start] == '\r' || json[start] == '\n'))
                start++;
            if (start >= json.Length || json[start] != '"') return null;
            start++; // skip opening quote
            var sb = new StringBuilder();
            while (start < json.Length)
            {
                if (json[start] == '\\' && start + 1 < json.Length)
                {
                    char next = json[start + 1];
                    if (next == '"') { sb.Append('"'); start += 2; }
                    else if (next == '\\') { sb.Append('\\'); start += 2; }
                    else if (next == 'n') { sb.Append('\n'); start += 2; }
                    else if (next == 'r') { sb.Append('\r'); start += 2; }
                    else if (next == 't') { sb.Append('\t'); start += 2; }
                    else { sb.Append(next); start += 2; }
                }
                else if (json[start] == '"') break;
                else { sb.Append(json[start]); start++; }
            }
            return sb.ToString();
        }

        private void ImportProfile()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "JSON profiles|*.json|All files|*.*";
                dlg.Title = "Import Profile";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string name = _profileMgr.ImportProfile(dlg.FileName);
                        RefreshProfileList();
                        SetStatus("Imported profile: " + name);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error importing: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ExportProfile()
        {
            if (_profileCombo.SelectedItem == null)
            {
                MessageBox.Show("Please select a profile first.", "Export Profile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string name = _profileCombo.SelectedItem.ToString();
            using (var dlg = new SaveFileDialog())
            {
                dlg.Filter = "JSON profiles|*.json";
                dlg.FileName = name + ".json";
                dlg.Title = "Export Profile";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _profileMgr.ExportProfile(name, dlg.FileName);
                        SetStatus("Exported profile: " + name);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error exporting: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void OpenProfilesFolder()
        {
            try { System.Diagnostics.Process.Start("explorer.exe", _profileMgr.ProfileDir); }
            catch { }
        }

        private string PromptProfileName(out string notes)
        {
            notes = "";
            using (var dlg = new Form())
            {
                dlg.Text = "Save Profile";
                dlg.Size = new Size(380, 250);
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;

                var lbl = new Label { Text = "Profile name:", Location = new Point(15, 15), AutoSize = true };
                var txt = new TextBox { Location = new Point(15, 35), Width = 330, Text = "MyProfile" };
                var lblNotes = new Label { Text = "Notes (optional):", Location = new Point(15, 65), AutoSize = true };
                var txtNotes = new TextBox { Location = new Point(15, 85), Width = 330, Height = 70, Multiline = true, ScrollBars = ScrollBars.Vertical };
                var btnOk = new Button { Text = "Save", DialogResult = DialogResult.OK, Location = new Point(185, 170), Width = 75 };
                var btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(270, 170), Width = 75 };

                if (_settings.DarkMode)
                    DarkModeHelper.ApplyTheme(dlg, true);

                dlg.AcceptButton = btnOk;
                dlg.CancelButton = btnCancel;
                dlg.Controls.AddRange(new Control[] { lbl, txt, lblNotes, txtNotes, btnOk, btnCancel });

                txt.SelectAll();

                if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(txt.Text))
                {
                    notes = txtNotes.Text;
                    // Sanitize filename
                    string name = txt.Text.Trim();
                    foreach (char c in Path.GetInvalidFileNameChars())
                        name = name.Replace(c, '_');
                    return name;
                }
                return null;
            }
        }

        // ---- Dark Mode ----

        private void ToggleDarkMode()
        {
            _settings.DarkMode = _darkModeItem.Checked;
            _settings.Save();
            DarkModeHelper.ApplyTheme(this, _settings.DarkMode);
        }

        // ---- Profile Info Panel ----

        private void OnProfileSelected()
        {
            if (_suppressProfileChange) return;
            UpdateProfileInfo();

            // Auto-restore when switching profiles
            if (_profileCombo.SelectedItem == null || _cameras.Count == 0) return;

            string name = _profileCombo.SelectedItem.ToString();
            try
            {
                SetStatus("Switching to profile '" + name + "'...");
                Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                var errors = _profileMgr.ApplyProfile(name, _cameras);

                Cursor = Cursors.Default;

                foreach (var tab in _cameraTabs)
                    tab.RefreshFromDevice();

                if (errors.Count > 0)
                    SetStatus("Profile '" + name + "': " + errors.Count + " warning(s)");
                else
                    SetStatus("Switched to profile: " + name);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                SetStatus("Error: " + ex.Message);
            }
        }

        private void UpdateProfileInfo()
        {
            if (_profileCombo.SelectedItem == null)
            {
                _notesLabel.Text = "";
                return;
            }

            string name = _profileCombo.SelectedItem.ToString();
            var profile = _profileMgr.LoadProfile(name);
            if (profile != null && !string.IsNullOrEmpty(profile.Notes))
                _notesLabel.Text = profile.Notes;
            else
                _notesLabel.Text = "(no notes)";
        }

        // ---- Global Hotkeys ----

        private void RegisterProfileHotkeys()
        {
            UnregisterAllHotkeys();
            _hotkeyProfiles.Clear();

            int id = 0;
            foreach (var kvp in _settings.Hotkeys)
            {
                _hotkeyProfiles.Add(kvp.Key);
                RegisterHotKey(Handle, HOTKEY_BASE_ID + id, kvp.Value.Modifiers, kvp.Value.Key);
                id++;
            }
        }

        private void UnregisterAllHotkeys()
        {
            for (int i = 0; i < _hotkeyProfiles.Count + 10; i++)
                UnregisterHotKey(Handle, HOTKEY_BASE_ID + i);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                int id = m.WParam.ToInt32() - HOTKEY_BASE_ID;
                if (id >= 0 && id < _hotkeyProfiles.Count)
                {
                    string profileName = _hotkeyProfiles[id];
                    var errors = _profileMgr.ApplyProfile(profileName, _cameras);
                    foreach (var tab in _cameraTabs) tab.RefreshFromDevice();

                    HotkeyBinding hk;
                    string hkName = _settings.Hotkeys.TryGetValue(profileName, out hk) ? hk.DisplayName : "?";
                    if (errors.Count == 0)
                        SetStatus("Hotkey " + hkName + ": restored '" + profileName + "'");
                    else
                        SetStatus("Hotkey " + hkName + ": restored '" + profileName + "' with " + errors.Count + " warning(s)");
                }
            }
            base.WndProc(ref m);
        }

        private void SetHotkeyForProfile()
        {
            if (_profileCombo.SelectedItem == null)
            {
                MessageBox.Show("Select a profile first.", "Set Hotkey", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string profileName = _profileCombo.SelectedItem.ToString();

            using (var dlg = new Form())
            {
                dlg.Text = "Set Hotkey for '" + profileName + "'";
                dlg.Size = new Size(350, 180);
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;
                dlg.KeyPreview = true;

                var lbl = new Label
                {
                    Text = "Press any key combination (e.g. Ctrl+F1, Alt+Shift+1)...\n\nPress Escape to cancel, Delete to remove hotkey.",
                    Location = new Point(15, 15),
                    Size = new Size(300, 60)
                };

                var resultLabel = new Label
                {
                    Text = "",
                    Location = new Point(15, 80),
                    Size = new Size(300, 25),
                    Font = new Font(Font.FontFamily, 11, FontStyle.Bold)
                };

                // Show current hotkey if set
                HotkeyBinding existing;
                if (_settings.Hotkeys.TryGetValue(profileName, out existing))
                    resultLabel.Text = "Current: " + existing.DisplayName;

                var btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Location = new Point(170, 110), Width = 75, Enabled = false };
                var btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(255, 110), Width = 75 };
                var btnRemove = new Button { Text = "Remove", Location = new Point(15, 110), Width = 75 };

                HotkeyBinding captured = null;

                btnRemove.Click += (s, e) =>
                {
                    _settings.Hotkeys.Remove(profileName);
                    _settings.Save();
                    RegisterProfileHotkeys();
                    SetStatus("Hotkey removed for '" + profileName + "'.");
                    dlg.DialogResult = DialogResult.Cancel;
                    dlg.Close();
                };

                if (_settings.DarkMode)
                    DarkModeHelper.ApplyTheme(dlg, true);

                dlg.AcceptButton = btnOk;
                dlg.CancelButton = btnCancel;
                dlg.Controls.AddRange(new Control[] { lbl, resultLabel, btnOk, btnCancel, btnRemove });

                dlg.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.Escape) return;
                    if (e.KeyCode == Keys.Delete)
                    {
                        _settings.Hotkeys.Remove(profileName);
                        _settings.Save();
                        RegisterProfileHotkeys();
                        SetStatus("Hotkey removed for '" + profileName + "'.");
                        dlg.DialogResult = DialogResult.Cancel;
                        dlg.Close();
                        return;
                    }

                    // Ignore lone modifier keys
                    if (e.KeyCode == Keys.ControlKey || e.KeyCode == Keys.ShiftKey ||
                        e.KeyCode == Keys.Menu || e.KeyCode == Keys.LWin || e.KeyCode == Keys.RWin)
                        return;

                    int mods = 0;
                    if (e.Control) mods |= 0x0002;
                    if (e.Alt) mods |= 0x0001;
                    if (e.Shift) mods |= 0x0004;

                    captured = new HotkeyBinding { Modifiers = mods, Key = (int)e.KeyCode };
                    resultLabel.Text = captured.DisplayName;
                    btnOk.Enabled = true;
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                };

                if (dlg.ShowDialog(this) == DialogResult.OK && captured != null)
                {
                    _settings.Hotkeys[profileName] = captured;
                    _settings.Save();
                    RegisterProfileHotkeys();
                    SetStatus("Hotkey " + captured.DisplayName + " set for '" + profileName + "'.");
                }
            }
        }

        private void ShowHotkeyHelp()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Global Hotkeys (work even when minimized):");
            sb.AppendLine();
            if (_settings.Hotkeys.Count == 0)
            {
                sb.AppendLine("  No hotkeys configured.");
                sb.AppendLine();
                sb.AppendLine("  Use the 'Set Hotkey' button in the toolbar");
                sb.AppendLine("  to assign a key combo to a profile.");
            }
            else
            {
                foreach (var kvp in _settings.Hotkeys)
                    sb.AppendLine("  " + kvp.Value.DisplayName + "  ->  " + kvp.Key);
            }
            MessageBox.Show(sb.ToString(), "Hotkeys", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ---- Lifecycle ----

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            RegisterProfileHotkeys();
            CheckForUpdatesSilent();
        }

        private void CheckForUpdatesSilent()
        {
            // Run on a background thread so startup isn't blocked
            System.Threading.ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    string apiUrl = "https://api.github.com/repos/" + AppInfo.GitHubOwner + "/" + AppInfo.GitHubRepo + "/releases/latest";
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    var request = (HttpWebRequest)WebRequest.Create(apiUrl);
                    request.UserAgent = "WebcamSettingsManager/" + AppInfo.Version;
                    request.Accept = "application/json";
                    request.Timeout = 5000;

                    string json;
                    using (var response = (HttpWebResponse)request.GetResponse())
                    using (var reader = new StreamReader(response.GetResponseStream()))
                        json = reader.ReadToEnd();

                    string tagName = ExtractJsonString(json, "tag_name");
                    string remoteVer = (tagName ?? "").TrimStart('v', 'V');
                    if (!string.IsNullOrEmpty(remoteVer) && remoteVer != AppInfo.Version)
                    {
                        BeginInvoke(new Action(delegate
                        {
                            SetStatus("Update available: " + tagName);
                            var result = MessageBox.Show(
                                "A new version (" + tagName + ") is available.\n\nWould you like to open the update dialog?",
                                "Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                            if (result == DialogResult.Yes)
                                CheckForUpdates();
                        }));
                    }
                }
                catch { }
            });
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            UnregisterAllHotkeys();
            foreach (var cam in _cameras)
                cam.Dispose();
            base.OnFormClosing(e);
        }
    }

    #endregion


    // ========================================================================
    // Entry Point
    // ========================================================================

    static class Program
    {
        [DllImport("kernel32.dll")]
        static extern bool AttachConsole(int dwProcessId);

        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                // CLI mode: attach to parent console for output
                AttachConsole(-1); // ATTACH_PARENT_PROCESS

                string profileName = null;
                string deviceName = null;
                bool showHelp = false;
                bool listProfiles = false;
                bool listDevices = false;

                for (int i = 0; i < args.Length; i++)
                {
                    string arg = args[i].ToLower();
                    if (arg == "--restore" || arg == "-r")
                    {
                        if (i + 1 < args.Length) profileName = args[++i];
                    }
                    else if (arg == "--device" || arg == "-d")
                    {
                        if (i + 1 < args.Length) deviceName = args[++i];
                    }
                    else if (arg == "--list-profiles" || arg == "-lp")
                    {
                        listProfiles = true;
                    }
                    else if (arg == "--list-devices" || arg == "-ld")
                    {
                        listDevices = true;
                    }
                    else if (arg == "--help" || arg == "-h" || arg == "/?")
                    {
                        showHelp = true;
                    }
                }

                if (showHelp)
                {
                    Console.WriteLine();
                    Console.WriteLine("WebcamSettingsManager - CLI Usage:");
                    Console.WriteLine();
                    Console.WriteLine("  No arguments          Launch GUI");
                    Console.WriteLine("  --restore <name>      Restore a saved profile");
                    Console.WriteLine("  --device <name>       Only restore to this camera (partial match)");
                    Console.WriteLine("  --list-profiles       List all saved profiles");
                    Console.WriteLine("  --list-devices        List all connected cameras");
                    Console.WriteLine("  --help                Show this help");
                    Console.WriteLine();
                    Console.WriteLine("Examples:");
                    Console.WriteLine("  WebcamSettingsManager.exe --restore dji");
                    Console.WriteLine("  WebcamSettingsManager.exe --restore dji --device OsmoPocket3");
                    Console.WriteLine();
                    return;
                }

                if (listProfiles)
                {
                    var mgr = new ProfileManager();
                    Console.WriteLine();
                    Console.WriteLine("Saved profiles:");
                    foreach (var p in mgr.ListProfiles())
                        Console.WriteLine("  " + p);
                    Console.WriteLine();
                    return;
                }

                if (listDevices)
                {
                    var devices = DeviceEnumerator.EnumerateDevices();
                    Console.WriteLine();
                    Console.WriteLine("Connected cameras (" + devices.Count + "):");
                    foreach (var d in devices)
                    {
                        Console.WriteLine("  " + d.FriendlyName);
                        Console.WriteLine("    Path: " + d.DevicePath);
                    }
                    Console.WriteLine();
                    foreach (var d in devices) d.Dispose();
                    return;
                }

                if (profileName != null)
                {
                    var mgr = new ProfileManager();
                    var profile = mgr.LoadProfile(profileName);
                    if (profile == null)
                    {
                        Console.WriteLine();
                        Console.WriteLine("ERROR: Profile '" + profileName + "' not found.");
                        Console.WriteLine();
                        return;
                    }

                    var devices = DeviceEnumerator.EnumerateDevices();
                    if (devices.Count == 0)
                    {
                        Console.WriteLine();
                        Console.WriteLine("ERROR: No cameras found.");
                        Console.WriteLine();
                        return;
                    }

                    // Find device filter by friendly name (partial, case-insensitive)
                    string devicePathFilter = null;
                    if (deviceName != null)
                    {
                        string lower = deviceName.ToLower();
                        foreach (var d in devices)
                        {
                            if (d.FriendlyName.ToLower().Contains(lower))
                            {
                                devicePathFilter = d.DevicePath;
                                break;
                            }
                        }
                        if (devicePathFilter == null)
                        {
                            Console.WriteLine();
                            Console.WriteLine("ERROR: No camera matching '" + deviceName + "' found.");
                            Console.WriteLine();
                            foreach (var d in devices) d.Dispose();
                            return;
                        }
                    }

                    Console.WriteLine();
                    Console.WriteLine("Restoring profile '" + profileName + "'...");
                    var errors = mgr.ApplyProfile(profileName, devices, devicePathFilter);

                    if (errors.Count == 0)
                        Console.WriteLine("OK - all settings restored.");
                    else
                    {
                        Console.WriteLine("Restored with " + errors.Count + " warning(s):");
                        foreach (var err in errors)
                            Console.WriteLine("  " + err);
                    }
                    Console.WriteLine();

                    foreach (var d in devices) d.Dispose();
                    return;
                }

                // No valid CLI action
                Console.WriteLine();
                Console.WriteLine("Unknown arguments. Use --help for usage.");
                Console.WriteLine();
                return;
            }

            // No args: launch GUI
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
