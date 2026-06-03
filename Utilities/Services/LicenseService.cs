using System;

namespace utilities.Services
{
    /// <summary>
    /// The user's licensing state. Modelled in full now even though real key/trial
    /// logic arrives in a later phase, so the ribbon and command model never need a
    /// second redesign.
    /// </summary>
    public enum LicenseState
    {
        Trial,
        Licensed,
        Expired,
        Offline,
        FeatureLocked
    }

    public interface ILicenseService
    {
        LicenseState State { get; }

        /// <summary>Days left in trial (0 when not in trial).</summary>
        int TrialDaysRemaining { get; }

        /// <summary>True when the given feature key is usable in the current state/tier.</summary>
        bool IsFeatureAvailable(string featureKey);

        /// <summary>Short status line for the Help/About area.</summary>
        string StatusText { get; }
    }

    /// <summary>
    /// Stub implementation: everything is licensed. The state machine, feature gate and
    /// status text are wired so the real implementation (key validation, trial countdown,
    /// offline grace) can be dropped in later without touching any command or ribbon code.
    /// During development, set <see cref="ForcedState"/> to exercise locked/expired UI.
    /// </summary>
    public sealed class StubLicenseService : ILicenseService
    {
        /// <summary>Dev/QA override to preview non-licensed UI. Null = behave as Licensed.</summary>
        public LicenseState? ForcedState { get; set; }

        public LicenseState State
        {
            get { return ForcedState ?? LicenseState.Licensed; }
        }

        public int TrialDaysRemaining
        {
            get { return State == LicenseState.Trial ? 30 : 0; }
        }

        public bool IsFeatureAvailable(string featureKey)
        {
            // "core" features are always available regardless of tier.
            if (string.IsNullOrEmpty(featureKey) ||
                string.Equals(featureKey, "core", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            switch (State)
            {
                case LicenseState.Licensed:
                case LicenseState.Trial:
                case LicenseState.Offline:
                    return true;
                case LicenseState.Expired:
                case LicenseState.FeatureLocked:
                    return false;
                default:
                    return false;
            }
        }

        public string StatusText
        {
            get
            {
                switch (State)
                {
                    case LicenseState.Trial:
                        return "Trial — " + TrialDaysRemaining + " day(s) remaining";
                    case LicenseState.Licensed:
                        return "Licensed";
                    case LicenseState.Expired:
                        return "Trial expired — upgrade to continue";
                    case LicenseState.Offline:
                        return "Licensed (offline)";
                    case LicenseState.FeatureLocked:
                        return "Some tools require an upgrade";
                    default:
                        return "Unknown";
                }
            }
        }
    }

    /// <summary>Ambient accessor so commands can gate without dependency wiring.</summary>
    public static class License
    {
        private static ILicenseService _current = new StubLicenseService();

        public static ILicenseService Current
        {
            get { return _current; }
            set { _current = value ?? new StubLicenseService(); }
        }
    }
}
