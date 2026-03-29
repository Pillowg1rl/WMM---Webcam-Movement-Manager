using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Windows.Forms;

// ============================================================================
// Webcam Settings Manager - DirectShow Camera Configuration Tool
// Single-file C# WinForms application. Compiles with built-in .NET Framework csc.exe.
// ============================================================================

namespace WebcamSettingsManager
{
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
        public Dictionary<string, ProfileDevice> Devices { get; set; }

        public Profile()
        {
            Devices = new Dictionary<string, ProfileDevice>();
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

        // PTZ cameras have motors that take time to move. This method retries
        // until the value is verified or max attempts reached.
        public bool SetCameraControlPropertyVerified(CameraControlProperty prop, int targetValue, int flags,
            int maxAttempts, int delayMs, out int finalValue)
        {
            finalValue = targetValue;
            if (!_hasCamCtrl) return false;

            // Get the step size to determine acceptable tolerance
            var range = GetCameraControlRange(prop);
            int tolerance = (range != null && range.Step > 0) ? range.Step : 1;

            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                // Set the target value
                int hr = _camCtrl.Set(prop, targetValue, flags);
                if (hr != 0) return false;

                // Wait for motor to move
                Thread.Sleep(delayMs);

                // Read back the actual value
                int currentVal, currentFlags;
                hr = _camCtrl.Get(prop, out currentVal, out currentFlags);
                if (hr != 0) return false;

                finalValue = currentVal;

                // Check if close enough (within step tolerance)
                if (Math.Abs(currentVal - targetValue) <= tolerance)
                    return true;
            }

            // Didn't reach target after all attempts
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
                        10, 100, out finalVal);
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

        public void SaveProfile(string name, List<CameraDevice> devices, string devicePathFilter = null)
        {
            var profile = new Profile
            {
                Name = name,
                Created = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss")
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
        }

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

        public MainForm()
        {
            Text = "Webcam Settings Manager";
            Size = new Size(900, 700);
            MinimumSize = new Size(700, 500);
            StartPosition = FormStartPosition.CenterScreen;

            _profileMgr = new ProfileManager();

            BuildMenu();
            BuildToolbar();
            BuildTabControl();
            BuildStatusBar();
            EnumerateAndPopulate();
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

            var devMenu = new ToolStripMenuItem("Devices");
            devMenu.DropDownItems.Add("Refresh Devices", null, (s, e) => RefreshDevices());

            var helpMenu = new ToolStripMenuItem("Help");
            helpMenu.DropDownItems.Add("About", null, (s, e) =>
                MessageBox.Show("Webcam Settings Manager\n\nSave and restore DirectShow webcam settings.\nSupports PTZ preset positions.\n\nBuilt with .NET Framework WinForms.",
                    "About", MessageBoxButtons.OK, MessageBoxIcon.Information));

            _menuStrip.Items.Add(fileMenu);
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
            _toolbar.Items.Add(new ToolStripButton("Save All", null, (s, e) => SaveProfile(true)) { ToolTipText = "Save all cameras to profile" });
            _toolbar.Items.Add(new ToolStripButton("Save Selected", null, (s, e) => SaveProfile(false)) { ToolTipText = "Save current camera to profile" });
            _toolbar.Items.Add(new ToolStripSeparator());
            _toolbar.Items.Add(new ToolStripButton("Restore All", null, (s, e) => RestoreProfile(true)) { ToolTipText = "Restore all cameras from profile" });
            _toolbar.Items.Add(new ToolStripButton("Restore Selected", null, (s, e) => RestoreProfile(false)) { ToolTipText = "Restore current camera from profile" });
            _toolbar.Items.Add(new ToolStripSeparator());

            var profileLabel = new ToolStripLabel("Profile:");
            _toolbar.Items.Add(profileLabel);

            _profileCombo = new ComboBox
            {
                Width = 180,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _toolbar.Items.Add(new ToolStripControlHost(_profileCombo));

            _toolbar.Items.Add(new ToolStripButton("Delete", null, (s, e) => DeleteProfile()) { ToolTipText = "Delete selected profile" });
            _toolbar.Items.Add(new ToolStripButton("Generate .bat", null, (s, e) => GenerateBat()) { ToolTipText = "Generate a .bat file to restore this profile" });

            Controls.Add(_toolbar);
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
            _profileCombo.Items.Clear();
            foreach (var name in _profileMgr.ListProfiles())
                _profileCombo.Items.Add(name);
            if (_profileCombo.Items.Count > 0)
                _profileCombo.SelectedIndex = 0;
        }

        private void SaveProfile(bool allDevices)
        {
            if (_cameras.Count == 0)
            {
                MessageBox.Show("No cameras available.", "Save Profile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string name = PromptProfileName();
            if (string.IsNullOrEmpty(name)) return;

            try
            {
                string filter = null;
                if (!allDevices)
                {
                    var selectedTab = _tabControl.SelectedTab as CameraTabPage;
                    if (selectedTab != null)
                        filter = selectedTab.Camera.DevicePath;
                }

                _profileMgr.SaveProfile(name, _cameras, filter);
                RefreshProfileList();

                // Select the newly saved profile
                int idx = _profileCombo.Items.IndexOf(name);
                if (idx >= 0) _profileCombo.SelectedIndex = idx;

                SetStatus("Saved profile: " + name + (allDevices ? " (all cameras)" : " (selected camera)"));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving profile: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RestoreProfile(bool allDevices)
        {
            if (_cameras.Count == 0)
            {
                MessageBox.Show("No cameras available.", "Restore Profile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_profileCombo.SelectedItem == null)
            {
                MessageBox.Show("Please select a profile first.", "Restore Profile", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string name = _profileCombo.SelectedItem.ToString();

            try
            {
                string filter = null;
                if (!allDevices)
                {
                    var selectedTab = _tabControl.SelectedTab as CameraTabPage;
                    if (selectedTab != null)
                        filter = selectedTab.Camera.DevicePath;
                }

                SetStatus("Restoring profile '" + name + "'... (PTZ motors may take a moment)");
                Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                var errors = _profileMgr.ApplyProfile(name, _cameras, filter);

                Cursor = Cursors.Default;

                // Refresh all slider positions
                foreach (var tab in _cameraTabs)
                    tab.RefreshFromDevice();

                if (errors.Count > 0)
                    SetStatus("Restored with " + errors.Count + " warning(s): " + string.Join(", ", errors.Take(3)));
                else
                    SetStatus("Restored profile: " + name + (allDevices ? " (all cameras)" : " (selected camera)"));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error restoring profile: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private string PromptProfileName()
        {
            using (var dlg = new Form())
            {
                dlg.Text = "Save Profile";
                dlg.Size = new Size(350, 150);
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;

                var lbl = new Label { Text = "Profile name:", Location = new Point(15, 20), AutoSize = true };
                var txt = new TextBox { Location = new Point(15, 45), Width = 300, Text = "MyProfile" };
                var btnOk = new Button { Text = "Save", DialogResult = DialogResult.OK, Location = new Point(155, 80), Width = 75 };
                var btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(240, 80), Width = 75 };

                dlg.AcceptButton = btnOk;
                dlg.CancelButton = btnCancel;
                dlg.Controls.AddRange(new Control[] { lbl, txt, btnOk, btnCancel });

                txt.SelectAll();

                if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(txt.Text))
                {
                    // Sanitize filename
                    string name = txt.Text.Trim();
                    foreach (char c in Path.GetInvalidFileNameChars())
                        name = name.Replace(c, '_');
                    return name;
                }
                return null;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
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
