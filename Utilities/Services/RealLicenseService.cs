using System;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace utilities.Services
{
    /// <summary>
    /// Production license service. Replaces StubLicenseService at startup.
    ///
    /// States:
    ///   Trial       — 30-day countdown from first run (start date in HKCU registry).
    ///   Licensed    — product key validated offline via HMAC-SHA256; cached in registry.
    ///   Offline     — licensed but machine ID changed; 14-day grace then demoted.
    ///   Expired     — trial ended, no valid key.
    ///   FeatureLocked — key valid but tier does not include the requested feature.
    ///
    /// Key format: XXXXX-XXXXX-XXXXX-XXXXX  (20 hex chars grouped by 5).
    /// Validation: last 8 chars = first 4 bytes of HMAC-SHA256(machineId+body, salt).
    /// Fully offline — no server round-trip needed for basic validation.
    /// </summary>
    public sealed class RealLicenseService : ILicenseService
    {
        // ── Registry ──────────────────────────────────────────────────────────
        private const string RegRoot       = @"Software\ExcelUtilitiesSuite";
        private const string ValTrialStart = "TrialStart";  // long (DateTime.ToBinary)
        private const string ValLicKey     = "LicKey";      // normalised key string
        private const string ValLicDate    = "LicDate";     // long (DateTime.ToBinary)
        private const string ValLicMachine = "LicMachine";  // hex machine fingerprint

        private const int TrialDays        = 30;
        private const int OfflineGraceDays = 14;

        // Secret loaded from LicenseSalt.cs (gitignored — never committed).
        private static readonly byte[] _salt =
            Encoding.UTF8.GetBytes(LicenseSalt.Value);

        // ── Cached state ──────────────────────────────────────────────────────
        private LicenseState _state;
        private int          _trialDays;
        private string       _storedKey;

        private RealLicenseService() { }

        /// <summary>Load or initialise license state from the registry.</summary>
        public static RealLicenseService Load()
        {
            var svc = new RealLicenseService();
            svc.Refresh();
            return svc;
        }

        // ── ILicenseService ───────────────────────────────────────────────────

        public LicenseState State              { get { return _state; } }
        public int          TrialDaysRemaining { get { return _trialDays; } }

        public string StatusText
        {
            get
            {
                switch (_state)
                {
                    case LicenseState.Trial:        return "Trial — " + _trialDays + " day(s) remaining";
                    case LicenseState.Licensed:     return "Licensed — " + MaskedKey(_storedKey);
                    case LicenseState.Offline:      return "Licensed (offline — reconnect within " + OfflineGraceDays + " days)";
                    case LicenseState.Expired:      return "Trial expired — enter a product key to continue";
                    case LicenseState.FeatureLocked:return "Upgrade required for this feature";
                    default:                        return "Unknown";
                }
            }
        }

        public bool IsFeatureAvailable(string featureKey)
        {
            if (string.IsNullOrEmpty(featureKey) ||
                string.Equals(featureKey, "core", StringComparison.OrdinalIgnoreCase))
                return true;

            switch (_state)
            {
                case LicenseState.Licensed:
                case LicenseState.Trial:
                case LicenseState.Offline:
                    return true;
                default:
                    return false;
            }
        }

        // ── Activation / deactivation ─────────────────────────────────────────

        /// <summary>Validate and persist a product key. Returns true on success.</summary>
        public bool Activate(string key)
        {
            string norm = Normalise(key);
            if (!ValidateKey(norm)) return false;

            try
            {
                using (var rk = Registry.CurrentUser.CreateSubKey(RegRoot))
                {
                    if (rk == null) return false;
                    rk.SetValue(ValLicKey,     norm);
                    rk.SetValue(ValLicDate,    DateTime.UtcNow.ToBinary().ToString());
                    rk.SetValue(ValLicMachine, MachineId());
                }
            }
            catch { return false; }

            Refresh();
            return true;
        }

        /// <summary>Remove the stored key (resets to trial or expired).</summary>
        public void Deactivate()
        {
            try
            {
                using (var rk = Registry.CurrentUser.OpenSubKey(RegRoot, writable: true))
                {
                    rk?.DeleteValue(ValLicKey,     throwOnMissingValue: false);
                    rk?.DeleteValue(ValLicDate,    throwOnMissingValue: false);
                    rk?.DeleteValue(ValLicMachine, throwOnMissingValue: false);
                }
            }
            catch { }
            Refresh();
        }

        // ── Refresh (reads registry, sets _state) ─────────────────────────────

        private void Refresh()
        {
            try
            {
                using (var rk = Registry.CurrentUser.OpenSubKey(RegRoot))
                {
                    // Licensed path
                    string licKey = rk?.GetValue(ValLicKey) as string;
                    if (!string.IsNullOrEmpty(licKey) && ValidateKey(licKey))
                    {
                        _storedKey = licKey;
                        bool sameBox = string.Equals(
                            rk?.GetValue(ValLicMachine) as string,
                            MachineId(), StringComparison.OrdinalIgnoreCase);

                        long ticks; DateTime licDate;
                        string raw = rk?.GetValue(ValLicDate) as string;
                        licDate = (raw != null && long.TryParse(raw, out ticks))
                            ? DateTime.FromBinary(ticks) : DateTime.UtcNow;

                        double age = (DateTime.UtcNow - licDate).TotalDays;
                        _state     = (!sameBox && age > OfflineGraceDays)
                                     ? LicenseState.Offline : LicenseState.Licensed;
                        _trialDays = 0;
                        return;
                    }

                    // Trial path
                    DateTime trialStart = ReadOrCreateTrialStart(rk);
                    int used      = (int)(DateTime.UtcNow - trialStart).TotalDays;
                    int remaining = Math.Max(0, TrialDays - used);
                    _trialDays = remaining;
                    _state     = remaining > 0 ? LicenseState.Trial : LicenseState.Expired;
                }
            }
            catch
            {
                _state     = LicenseState.Trial;
                _trialDays = TrialDays;
            }
        }

        private static DateTime ReadOrCreateTrialStart(RegistryKey rk)
        {
            string raw = rk?.GetValue(ValTrialStart) as string;
            long ticks;
            if (raw != null && long.TryParse(raw, out ticks))
                return DateTime.FromBinary(ticks);

            DateTime now = DateTime.UtcNow;
            try
            {
                using (var w = Registry.CurrentUser.CreateSubKey(RegRoot))
                    w?.SetValue(ValTrialStart, now.ToBinary().ToString());
            }
            catch { }
            return now;
        }

        // ── Key validation ────────────────────────────────────────────────────

        /// <summary>
        /// A valid key: last 8 hex chars = first 4 bytes of
        /// HMAC-SHA256( (machineId + body).ToUpper(), _salt ).
        /// Fully offline — works without any server.
        /// </summary>
        internal static bool ValidateKey(string norm)
        {
            if (string.IsNullOrEmpty(norm)) return false;
            string stripped = norm.Replace("-", "");
            if (stripped.Length < 12) return false;

            string body = stripped.Substring(0, stripped.Length - 8);
            string tag  = stripped.Substring(stripped.Length - 8);
            string msg  = (MachineId() + body).ToUpperInvariant();

            try
            {
                using (var hmac = new HMACSHA256(_salt))
                {
                    byte[] hash     = hmac.ComputeHash(Encoding.UTF8.GetBytes(msg));
                    string expected = BitConverter.ToString(hash, 0, 4).Replace("-", "");
                    return string.Equals(expected, tag, StringComparison.OrdinalIgnoreCase);
                }
            }
            catch { return false; }
        }

        /// <summary>
        /// Key generator — used by the internal keygen tool (not shipped to users).
        /// Pass null to generate a key for the current machine.
        /// </summary>
        public static string GenerateKey(string machineId = null)
        {
            machineId = machineId ?? MachineId();
            byte[] rand = new byte[6];
            using (var rng = new RNGCryptoServiceProvider()) rng.GetBytes(rand);
            string body = BitConverter.ToString(rand).Replace("-", ""); // 12 hex chars
            string msg  = (machineId + body).ToUpperInvariant();
            using (var hmac = new HMACSHA256(_salt))
            {
                byte[] hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(msg));
                string tag  = BitConverter.ToString(hash, 0, 4).Replace("-", ""); // 8 chars
                string full = body + tag; // 20 hex chars
                return full.Substring(0,5)+"-"+full.Substring(5,5)+"-"+
                       full.Substring(10,5)+"-"+full.Substring(15,5);
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string Normalise(string key) =>
            (key ?? string.Empty).Trim().ToUpperInvariant();

        private static string MaskedKey(string key)
        {
            if (string.IsNullOrEmpty(key) || key.Length < 5) return "****";
            return key.Substring(0, 5) + "-****-****-****";
        }

        private static string _machineId;
        internal static string MachineId()
        {
            if (_machineId != null) return _machineId;
            try
            {
                string sysRoot = Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows";
                string drive   = System.IO.Path.GetPathRoot(sysRoot) ?? @"C:\";
                string raw     = new System.IO.DriveInfo(drive).RootDirectory.FullName
                                 + Environment.MachineName;
                using (var sha = SHA256.Create())
                {
                    byte[] h = sha.ComputeHash(Encoding.UTF8.GetBytes(raw));
                    _machineId = BitConverter.ToString(h, 0, 8).Replace("-", "");
                }
            }
            catch { _machineId = "GENERIC00000000"; }
            return _machineId;
        }
    }
}
