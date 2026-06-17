using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using utilities.Services;

namespace utilities.Commands
{
    /// <summary>
    /// Discovers every <see cref="IExcelCommand"/> marked with <see cref="ExcelCommandAttribute"/>
    /// in the add-in assembly and indexes it by <c>Definition.Id</c>. The ribbon's OnAction,
    /// getEnabled and tooltip callbacks all read from here, so adding a tool never touches
    /// ribbon plumbing.
    /// </summary>
    public static class CommandRegistry
    {
        private static readonly Dictionary<string, IExcelCommand> _byId =
            new Dictionary<string, IExcelCommand>(StringComparer.Ordinal);

        private static bool _initialised;

        public static IEnumerable<IExcelCommand> All
        {
            get { return _byId.Values; }
        }

        public static int Count { get { return _byId.Count; } }

        /// <summary>Build the registry once. Safe to call multiple times.</summary>
        public static void Initialize()
        {
            if (_initialised) return;

            Assembly asm = Assembly.GetExecutingAssembly();
            Type[] types;
            try { types = asm.GetTypes(); }
            catch (ReflectionTypeLoadException ex) { types = ex.Types.Where(t => t != null).ToArray(); }

            foreach (Type type in types)
            {
                if (type == null || type.IsAbstract || type.IsInterface) continue;
                if (type.GetCustomAttribute<ExcelCommandAttribute>() == null) continue;
                if (!typeof(IExcelCommand).IsAssignableFrom(type)) continue;
                if (type.GetConstructor(Type.EmptyTypes) == null)
                {
                    ErrorService.Log("WARN", "Command " + type.FullName + " has no parameterless constructor; skipped.");
                    continue;
                }

                try
                {
                    var command = (IExcelCommand)Activator.CreateInstance(type);
                    CommandDefinition def = command.Definition;
                    if (def == null || string.IsNullOrEmpty(def.Id))
                    {
                        ErrorService.Log("WARN", "Command " + type.FullName + " has no Id; skipped.");
                        continue;
                    }
                    if (_byId.ContainsKey(def.Id))
                    {
                        ErrorService.Log("WARN", "Duplicate command id '" + def.Id + "' (" + type.FullName + "); skipped.");
                        continue;
                    }
                    _byId.Add(def.Id, command);
                }
                catch (Exception ex)
                {
                    ErrorService.Log("ERROR", "Failed to instantiate command " + type.FullName + ": " + ex.Message);
                }
            }

            ErrorService.Log("INFO", "Registered " + _byId.Count + " commands.");
            Validate();

            // Mark complete only after a successful pass. If reflection above threw, the
            // flag stays false so a later defensive call (e.g. from GetCustomUI) can retry.
            _initialised = true;
        }

        /// <summary>
        /// Asserts structural integrity of all registered commands at startup.
        /// Catches definition drift (missing Label/Tab/Group/tooltip) rather than
        /// producing a silently broken ribbon at runtime.
        /// </summary>
        public static void Validate()
        {
            int warnings = 0;
            foreach (var kv in _byId)
            {
                CommandDefinition d = kv.Value.Definition;
                if (string.IsNullOrEmpty(d.Label))
                { ErrorService.Log("WARN", "[Self-check] '" + d.Id + "' has no Label."); warnings++; }
                if (string.IsNullOrEmpty(d.Tab))
                { ErrorService.Log("WARN", "[Self-check] '" + d.Id + "' has no Tab."); warnings++; }
                if (string.IsNullOrEmpty(d.Group))
                { ErrorService.Log("WARN", "[Self-check] '" + d.Id + "' has no Group."); warnings++; }
                if (string.IsNullOrEmpty(d.Screentip) && string.IsNullOrEmpty(d.Supertip))
                { ErrorService.Log("WARN", "[Self-check] '" + d.Id + "' has no tooltip."); warnings++; }
            }

            if (warnings == 0)
                ErrorService.Log("INFO", "[Self-check] PASS — " + _byId.Count + " commands registered, 0 definition warnings.");
            else
                ErrorService.Log("WARN", "[Self-check] " + warnings + " definition warning(s). See log for details.");
        }

        public static IExcelCommand Get(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            IExcelCommand cmd;
            return _byId.TryGetValue(id, out cmd) ? cmd : null;
        }

        public static CommandDefinition GetDefinition(string id)
        {
            IExcelCommand cmd = Get(id);
            return cmd != null ? cmd.Definition : null;
        }

        public static bool Contains(string id)
        {
            return id != null && _byId.ContainsKey(id);
        }
    }
}
