namespace RegistryTools
{
    using System;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Runtime.CompilerServices;
    using System.Security;
    using Microsoft.Win32;
    using Tracing;

    /// <summary>
    /// Represents settings stored in the registry.
    /// </summary>
    public class Settings : IDisposable
    {
        #region Init
        /// <summary>
        /// Initializes a new instance of the <see cref="Settings"/> class.
        /// </summary>
        /// <param name="rootName">A root name to use for sets of products key, for example the company name.</param>
        /// <param name="productName">A product name to use for all settings.</param>
        public Settings(string rootName, string productName)
        {
            RootName = rootName;
            ProductName = productName;
            Tracer = null;

            OpenKey();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Settings"/> class.
        /// </summary>
        /// <param name="rootName">A root name to use for sets of products key, for example the company name.</param>
        /// <param name="productName">A product name to use for all settings.</param>
        /// <param name="tracer">A tracer for diagnostic purpose.</param>
        public Settings(string rootName, string productName, ITracer tracer)
        {
            RootName = rootName;
            ProductName = productName;
            Tracer = tracer;

            TraceStart();

            OpenKey();

            TraceEnd();
        }
        #endregion

        #region Properties
        /// <summary>
        /// Gets the root name.
        /// </summary>
        public string RootName { get; }

        /// <summary>
        /// Gets the product name.
        /// </summary>
        public string ProductName { get; }

        /// <summary>
        /// Gets the tracer.
        /// </summary>
        public ITracer? Tracer { get; }
        #endregion

        #region Client Interface
        /// <summary>
        /// Renews the registry key in case it's been kept for a long time. Windows might have decided to close it on its own.
        /// Use this method after a system wake up for instance.
        /// </summary>
        public void RenewKey()
        {
            TraceStart();

            CloseKey();
            OpenKey();

            TraceEnd();
        }

        /// <summary>
        /// Checks is a value is set.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <returns>True if the value is set; otherwise, false.</returns>
        public bool IsValueSet(string valueName)
        {
            TraceStart();

            bool Result = GetSettingKey(valueName, isExistQuery: true) != null;

            TraceEnd();

            return Result;
        }

        /// <summary>
        /// Gets a <see cref="bool"/> value, using a default value if not found in the registry settings.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="defaultValue">The default to use if not found in the registry.</param>
        /// <param name="value">The value upon return. It will be <paramref name="defaultValue"/> if the method returns false.</param>
        /// <returns>True if found in the registry; otherwise, false.</returns>
        public bool GetBool(string valueName, bool defaultValue, out bool value)
        {
            TraceStart();

            bool Result;
            int? KeyValue = GetSettingKey(valueName) as int?;

            if (KeyValue.HasValue)
            {
                value = KeyValue.Value != 0;
                Result = true;
            }
            else
            {
                value = defaultValue;
                Result = false;
            }

            TraceEnd();

            return Result;
        }

        /// <summary>
        /// Sets a <see cref="bool"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetBool(string valueName, bool value)
        {
            TraceStart();

            SetSettingKey(valueName, value ? 1 : 0, RegistryValueKind.DWord);

            TraceEnd();
        }

        /// <summary>
        /// Gets a <see cref="int"/> value, using a default value if not found in the registry settings.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="defaultValue">The default to use if not found in the registry.</param>
        /// <param name="value">The value upon return. It will be <paramref name="defaultValue"/> if the method returns false.</param>
        /// <returns>True if found in the registry; otherwise, false.</returns>
        public bool GetInt(string valueName, int defaultValue, out int value)
        {
            TraceStart();

            bool Result;
            int? KeyValue = GetSettingKey(valueName) as int?;

            if (KeyValue.HasValue)
            {
                value = KeyValue.Value;
                Result = true;
            }
            else
            {
                value = defaultValue;
                Result = false;
            }

            TraceEnd();

            return Result;
        }

        /// <summary>
        /// Sets a <see cref="int"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetInt(string valueName, int value)
        {
            TraceStart();

            SetSettingKey(valueName, value, RegistryValueKind.DWord);

            TraceEnd();
        }

        /// <summary>
        /// Gets a <see cref="string"/> value, using a default value if not found in the registry settings.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="defaultValue">The default to use if not found in the registry.</param>
        /// <param name="value">The value upon return. It will be <paramref name="defaultValue"/> if the method returns false.</param>
        /// <returns>True if found in the registry; otherwise, false.</returns>
        public bool GetString(string valueName, string defaultValue, out string value)
        {
            TraceStart();

            bool Result;
            string? KeyValue = GetSettingKey(valueName) as string;

            if (KeyValue != null)
            {
                value = KeyValue;
                Result = true;
            }
            else
            {
                value = defaultValue;
                Result = false;
            }

            TraceEnd();

            return Result;
        }

        /// <summary>
        /// Sets a <see cref="string"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetString(string valueName, string value)
        {
            TraceStart();

            if (value == null)
                DeleteSetting(valueName);
            else
                SetSettingKey(valueName, value, RegistryValueKind.String);

            TraceEnd();
        }

        /// <summary>
        /// Gets a <see cref="float"/> value, using a default value if not found in the registry settings.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="defaultValue">The default to use if not found in the registry.</param>
        /// <param name="value">The value upon return. It will be <paramref name="defaultValue"/> if the method returns false.</param>
        /// <returns>True if found in the registry; otherwise, false.</returns>
        public bool GetFloat(string valueName, float defaultValue, out float value)
        {
            TraceStart();

            bool Result;
            string? KeyValue = GetSettingKey(valueName) as string;

            if (KeyValue != null && float.TryParse(KeyValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                Result = true;
            else
            {
                value = defaultValue;
                Result = false;
            }

            TraceEnd();

            return Result;
        }

        /// <summary>
        /// Sets a <see cref="float"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetFloat(string valueName, float value)
        {
            TraceStart();

            string StringValue = value.ToString(CultureInfo.InvariantCulture);
            SetSettingKey(valueName, StringValue, RegistryValueKind.String);

            TraceEnd();
        }

        /// <summary>
        /// Gets a <see cref="double"/> value, using a default value if not found in the registry settings.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="defaultValue">The default to use if not found in the registry.</param>
        /// <param name="value">The value upon return. It will be <paramref name="defaultValue"/> if the method returns false.</param>
        /// <returns>True if found in the registry; otherwise, false.</returns>
        public bool GetDouble(string valueName, double defaultValue, out double value)
        {
            TraceStart();

            bool Result;
            string? KeyValue = GetSettingKey(valueName) as string;

            if (KeyValue != null && double.TryParse(KeyValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                Result = true;
            else
            {
                value = defaultValue;
                Result = false;
            }

            TraceEnd();

            return Result;
        }

        /// <summary>
        /// Sets a <see cref="double"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetDouble(string valueName, double value)
        {
            TraceStart();

            string StringValue = value.ToString(CultureInfo.InvariantCulture);
            SetSettingKey(valueName, StringValue, RegistryValueKind.String);

            TraceEnd();
        }

        /// <summary>
        /// Gets a <see cref="Guid"/> value, using a default value if not found in the registry settings.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="defaultValue">The default to use if not found in the registry.</param>
        /// <param name="value">The value upon return. It will be <paramref name="defaultValue"/> if the method returns false.</param>
        /// <returns>True if found in the registry; otherwise, false.</returns>
        public bool GetGuid(string valueName, Guid defaultValue, out Guid value)
        {
            TraceStart();

            bool Result;
            string? KeyValue = GetSettingKey(valueName) as string;

            if (KeyValue != null && Guid.TryParse(KeyValue, out value))
                Result = true;
            else
            {
                value = defaultValue;
                Result = false;
            }

            TraceEnd();

            return Result;
        }

        /// <summary>
        /// Sets a <see cref="Guid"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetGuid(string valueName, Guid value)
        {
            TraceStart();

            string StringValue = value.ToString();
            SetSettingKey(valueName, StringValue, RegistryValueKind.String);

            TraceEnd();
        }
        #endregion

        #region Implementation
        private void OpenKey()
        {
            TraceStart();

            RegistryKey Key = Registry.CurrentUser.OpenSubKey(@"Software", true);
            Key = Key.CreateSubKey(RootName);
            SettingKey = Key.CreateSubKey(ProductName);

            TraceEnd();
        }

        private void CloseKey()
        {
            TraceStart();

            using (RegistryKey? Key = SettingKey)
            {
                SettingKey = null;
            }

            TraceEnd();
        }

        private object? GetSettingKey(string valueName, bool isExistQuery = false)
        {
            TraceStart();

            object? Result;

            try
            {
                Result = SettingKey?.GetValue(valueName);
            }
            catch (Exception e) when (e is SecurityException || e is ObjectDisposedException || e is IOException || e is UnauthorizedAccessException)
            {
                if (!isExistQuery)
                    Tracer?.Write(Category.Debug, e, "Failed to query value");

                Result = null;
            }

            TraceEnd();

            return Result;
        }

        private void SetSettingKey(string valueName, object value, RegistryValueKind kind)
        {
            TraceStart();

            try
            {
                SettingKey?.SetValue(valueName, value, kind);
            }
            catch (Exception e) when (e is SecurityException || e is ObjectDisposedException || e is IOException || e is UnauthorizedAccessException)
            {
                Tracer?.Write(Category.Debug, e, "Failed to set value");
            }

            TraceEnd();
        }

        private void DeleteSetting(string valueName)
        {
            TraceStart();

            try
            {
                SettingKey?.DeleteValue(valueName, false);
            }
            catch (Exception e) when (e is SecurityException || e is ObjectDisposedException || e is IOException || e is UnauthorizedAccessException)
            {
                Tracer?.Write(Category.Debug, e, "Failed to delete value");
            }

            TraceEnd();
        }

        private RegistryKey? SettingKey;
        #endregion

        #region Miscellaneous
        private void TraceStart([CallerMemberName] string methodName = "")
        {
            Debug.Assert(methodName.Length > 0, "Unexpected empty method name");

            Tracer?.Write(Category.Debug, $"{methodName} starting");
        }

        private void TraceEnd([CallerMemberName] string methodName = "")
        {
            Debug.Assert(methodName.Length > 0, "Unexpected empty method name");

            Tracer?.Write(Category.Debug, $"{methodName} done");
        }
        #endregion

        #region Implementation of IDisposable
        /// <summary>
        /// Called when an object should release its resources.
        /// </summary>
        /// <param name="isDisposing">Indicates if resources must be disposed now.</param>
        protected virtual void Dispose(bool isDisposing)
        {
            if (!IsDisposed)
            {
                IsDisposed = true;

                if (isDisposing)
                    DisposeNow();
            }
        }

        /// <summary>
        /// Called when an object should release its resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Finalizes an instance of the <see cref="Settings"/> class.
        /// </summary>
        ~Settings()
        {
            Dispose(false);
        }

        /// <summary>
        /// True after <see cref="Dispose(bool)"/> has been invoked.
        /// </summary>
        private bool IsDisposed;

        /// <summary>
        /// Disposes of every reference that must be cleaned up.
        /// </summary>
        private void DisposeNow()
        {
            CloseKey();
        }
        #endregion
    }
}
