using System;
using System.Collections.Generic;
using System.IO;

namespace utilities.Services
{
    /// <summary>
    /// Dependency-free per-user preference store used by dialogs to remember last-used
    /// options. Persisted as simple "key=value" lines under %APPDATA% (no JSON library
    /// dependency on .NET Framework). Keys are namespaced by command id, e.g.
    /// "ChangeCase.LastMode=Upper".
    /// </summary>
    public static class SettingsService
    {
        private static readonly object _lock = new object();
        private static Dictionary<string, string> _cache;

        private static string FilePath
        {
            get { return Path.Combine(ErrorService.LogDirectory, "settings.txt"); }
        }

        private static Dictionary<string, string> Cache
        {
            get
            {
                if (_cache == null) Load();
                return _cache;
            }
        }

        public static string Get(string key, string fallback = null)
        {
            lock (_lock)
            {
                string value;
                return Cache.TryGetValue(key, out value) ? value : fallback;
            }
        }

        public static int GetInt(string key, int fallback)
        {
            int v;
            return int.TryParse(Get(key), out v) ? v : fallback;
        }

        public static bool GetBool(string key, bool fallback)
        {
            bool v;
            return bool.TryParse(Get(key), out v) ? v : fallback;
        }

        public static void Set(string key, string value)
        {
            lock (_lock)
            {
                Cache[key] = value ?? string.Empty;
                Save();
            }
        }

        public static void Set(string key, int value) { Set(key, value.ToString()); }
        public static void Set(string key, bool value) { Set(key, value.ToString()); }

        private static void Load()
        {
            _cache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (!File.Exists(FilePath)) return;
                foreach (string line in File.ReadAllLines(FilePath))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string k = line.Substring(0, eq).Trim();
                    string v = line.Substring(eq + 1);
                    _cache[k] = v;
                }
            }
            catch
            {
                // Corrupt/locked settings should never break startup.
            }
        }

        private static void Save()
        {
            try
            {
                Directory.CreateDirectory(ErrorService.LogDirectory);
                var lines = new List<string>(_cache.Count);
                foreach (var kv in _cache)
                {
                    // Drop newlines so each entry stays on one line.
                    string v = (kv.Value ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
                    lines.Add(kv.Key + "=" + v);
                }
                File.WriteAllLines(FilePath, lines.ToArray());
            }
            catch
            {
            }
        }
    }
}
