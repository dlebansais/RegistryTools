namespace RegistryTools
{
    using System;
    using System.Globalization;
    using System.IO;
    using System.Security;
    using Microsoft.Win32;

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

            OpenKey();
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
        #endregion

        #region Client Interface
        /// <summary>
        /// Renews the registry key in case it's been kept for a long time. Windows might have decided to close it on its own.
        /// Use this method after a system wake up for instance.
        /// </summary>
        public void RenewKey()
        {
            CloseKey();
            OpenKey();
        }

        /// <summary>
        /// Checks is a value is set.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <returns>True if the value is set; otherwise, false.</returns>
        public bool IsValueSet(string valueName)
        {
            return GetSettingKey(valueName) != null;
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
            int? KeyValue = GetSettingKey(valueName) as int?;

            if (KeyValue.HasValue)
            {
                value = KeyValue.Value != 0;
                return true;
            }
            else
            {
                value = defaultValue;
                return false;
            }
        }

        /// <summary>
        /// Sets a <see cref="bool"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetBool(string valueName, bool value)
        {
            SetSettingKey(valueName, value ? 1 : 0, RegistryValueKind.DWord);
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
            int? KeyValue = GetSettingKey(valueName) as int?;

            if (KeyValue.HasValue)
            {
                value = KeyValue.Value;
                return true;
            }
            else
            {
                value = defaultValue;
                return false;
            }
        }

        /// <summary>
        /// Sets a <see cref="int"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetInt(string valueName, int value)
        {
            SetSettingKey(valueName, value, RegistryValueKind.DWord);
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
            string? KeyValue = GetSettingKey(valueName) as string;

            if (KeyValue != null)
            {
                value = KeyValue;
                return true;
            }
            else
            {
                value = defaultValue;
                return false;
            }
        }

        /// <summary>
        /// Sets a <see cref="string"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetString(string valueName, string value)
        {
            if (value == null)
                DeleteSetting(valueName);
            else
                SetSettingKey(valueName, value, RegistryValueKind.String);
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
            string? KeyValue = GetSettingKey(valueName) as string;

            if (KeyValue != null && float.TryParse(KeyValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                return true;
            else
            {
                value = defaultValue;
                return false;
            }
        }

        /// <summary>
        /// Sets a <see cref="float"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetFloat(string valueName, float value)
        {
            string StringValue = value.ToString(CultureInfo.InvariantCulture);
            SetSettingKey(valueName, StringValue, RegistryValueKind.String);
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
            string? KeyValue = GetSettingKey(valueName) as string;

            if (KeyValue != null && double.TryParse(KeyValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                return true;
            else
            {
                value = defaultValue;
                return false;
            }
        }

        /// <summary>
        /// Sets a <see cref="double"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetDouble(string valueName, double value)
        {
            string StringValue = value.ToString(CultureInfo.InvariantCulture);
            SetSettingKey(valueName, StringValue, RegistryValueKind.String);
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
            string? KeyValue = GetSettingKey(valueName) as string;

            if (KeyValue != null && Guid.TryParse(KeyValue, out value))
                return true;
            else
            {
                value = defaultValue;
                return false;
            }
        }

        /// <summary>
        /// Sets a <see cref="Guid"/> value in the registry, creating the entry if necessary.
        /// Upon return, <see cref="IsValueSet"/> will return true if called.
        /// </summary>
        /// <param name="valueName">The value name.</param>
        /// <param name="value">The value to set.</param>
        public void SetGuid(string valueName, Guid value)
        {
            string StringValue = value.ToString();
            SetSettingKey(valueName, StringValue, RegistryValueKind.String);
        }
        #endregion

        #region Implementation
        private void OpenKey()
        {
            RegistryKey Key = Registry.CurrentUser.OpenSubKey(@"Software", true);
            Key = Key.CreateSubKey(RootName);
            SettingKey = Key.CreateSubKey(ProductName);
        }

        private void CloseKey()
        {
            using (RegistryKey? Key = SettingKey)
            {
                SettingKey = null;
            }
        }

        private object? GetSettingKey(string valueName)
        {
            try
            {
                return SettingKey?.GetValue(valueName);
            }
            catch (Exception e) when (e is SecurityException || e is ObjectDisposedException || e is IOException || e is UnauthorizedAccessException)
            {
                return null;
            }
        }

        private void SetSettingKey(string valueName, object value, RegistryValueKind kind)
        {
            try
            {
                SettingKey?.SetValue(valueName, value, kind);
            }
            catch (Exception e) when (e is SecurityException || e is ObjectDisposedException || e is IOException || e is UnauthorizedAccessException)
            {
            }
        }

        private void DeleteSetting(string valueName)
        {
            try
            {
                SettingKey?.DeleteValue(valueName, false);
            }
            catch (Exception e) when (e is SecurityException || e is ObjectDisposedException || e is IOException || e is UnauthorizedAccessException)
            {
            }
        }

        private RegistryKey? SettingKey;
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
