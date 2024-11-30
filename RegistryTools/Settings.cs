namespace RegistryTools;

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Security;
using Microsoft.Extensions.Logging;
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
        Logger = null;

        OpenKey();
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Settings"/> class.
    /// </summary>
    /// <param name="rootName">A root name to use for sets of products key, for example the company name.</param>
    /// <param name="productName">A product name to use for all settings.</param>
    /// <param name="tracer">A tracer for diagnostic purpose.</param>
    public Settings(string rootName, string productName, ILogger tracer)
    {
        RootName = rootName;
        ProductName = productName;
        Logger = tracer;

        TraceStart();

        OpenKey();

        TraceEnd();
    }
    #endregion

    #region Properties
    /// <summary>
    /// Gets the root name.
    /// </summary>
    public string RootName { get; init; }

    /// <summary>
    /// Gets the product name.
    /// </summary>
    public string ProductName { get; init; }

    /// <summary>
    /// Gets the tracer.
    /// </summary>
    public ILogger? Logger { get; init; }
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

        bool Result = GetSettingKey(valueName, isExistQuery: true) is not null;

        TraceEnd();

        return Result;
    }

    /// <summary>
    /// Gets a <see cref="bool"/> value, using a default value if not found in the registry settings.
    /// </summary>
    /// <param name="valueName">The value name.</param>
    /// <param name="defaultValue">The default to use if not found in the registry.</param>
    /// <returns>The value read if found in the registry; otherwise, the default value in <paramref name="defaultValue"/>.</returns>
    public bool GetBool(string valueName, bool defaultValue)
    {
        TraceStart();

        bool Result;
        int? KeyValue = GetSettingKey(valueName) as int?;

        Result = KeyValue.HasValue ? KeyValue.Value != 0 : defaultValue;

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
    /// <returns>The value read if found in the registry; otherwise, the default value in <paramref name="defaultValue"/>.</returns>
    public int GetInt(string valueName, int defaultValue)
    {
        TraceStart();

        int Result;
        int? KeyValue = GetSettingKey(valueName) as int?;

        Result = KeyValue ?? defaultValue;

        TraceEnd();

        return Result;
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
    /// <returns>The value read if found in the registry; otherwise, the default value in <paramref name="defaultValue"/>.</returns>
    public string GetString(string valueName, string defaultValue)
    {
        TraceStart();

        string Result;
        string? KeyValue = GetSettingKey(valueName) as string;

        Result = KeyValue ?? defaultValue;

        TraceEnd();

        return Result;
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

        if (KeyValue is not null)
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

        if (value is null)
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
    /// <returns>The value read if found in the registry; otherwise, the default value in <paramref name="defaultValue"/>.</returns>
    public float GetFloat(string valueName, float defaultValue)
    {
        TraceStart();

        float Result;

        if (GetSettingKey(valueName) is not string KeyValue)
            Result = defaultValue;
        else if (!float.TryParse(KeyValue, NumberStyles.Float, CultureInfo.InvariantCulture, out Result))
            Result = defaultValue;

        TraceEnd();

        return Result;
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

        if (KeyValue is not null && float.TryParse(KeyValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
        {
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
    /// <returns>The value read if found in the registry; otherwise, the default value in <paramref name="defaultValue"/>.</returns>
    public double GetDouble(string valueName, double defaultValue)
    {
        TraceStart();

        double Result;

        if (GetSettingKey(valueName) is not string KeyValue)
            Result = defaultValue;
        else if (!double.TryParse(KeyValue, NumberStyles.Float, CultureInfo.InvariantCulture, out Result))
            Result = defaultValue;

        TraceEnd();

        return Result;
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

        if (KeyValue is not null && double.TryParse(KeyValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
        {
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
    /// <returns>The value read if found in the registry; otherwise, the default value in <paramref name="defaultValue"/>.</returns>
    public Guid GetGuid(string valueName, Guid defaultValue)
    {
        TraceStart();

        Guid Result;

        if (GetSettingKey(valueName) is not string KeyValue)
            Result = defaultValue;
        else if (!Guid.TryParse(KeyValue, out Result))
            Result = defaultValue;

        TraceEnd();

        return Result;
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

        if (KeyValue is not null && Guid.TryParse(KeyValue, out value))
        {
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

        using RegistryKey? SoftwareKey = Registry.CurrentUser.OpenSubKey(@"Software", true);
        using RegistryKey? RootKey = SoftwareKey?.CreateSubKey(RootName);

        SettingKey?.Dispose();
        SettingKey = RootKey?.CreateSubKey(ProductName);

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
        catch (Exception e) when (e is SecurityException or ObjectDisposedException or IOException or UnauthorizedAccessException)
        {
            if (!isExistQuery && Logger is not null)
                LoggerMessage.Define(LogLevel.Debug, 0, "Failed to query value")(Logger, e);

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
        catch (Exception e) when (e is SecurityException or ObjectDisposedException or IOException or UnauthorizedAccessException)
        {
            if (Logger is not null)
                LoggerMessage.Define(LogLevel.Debug, 0, "Failed to set value")(Logger, e);
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
        catch (Exception e) when (e is SecurityException or ObjectDisposedException or IOException or UnauthorizedAccessException)
        {
            if (Logger is not null)
                LoggerMessage.Define(LogLevel.Debug, 0, "Failed to delete value")(Logger, e);
        }

        TraceEnd();
    }

    private RegistryKey? SettingKey;
    #endregion

    #region Miscellaneous
    private void TraceStart([CallerMemberName] string methodName = "")
    {
        Debug.Assert(methodName.Length > 0, "Unexpected empty method name");

        if (Logger is not null)
            LoggerMessage.Define(LogLevel.Debug, 0, $"{methodName} starting")(Logger, null);
    }

    private void TraceEnd([CallerMemberName] string methodName = "")
    {
        Debug.Assert(methodName.Length > 0, "Unexpected empty method name");

        if (Logger is not null)
            LoggerMessage.Define(LogLevel.Debug, 0, $"{methodName} done")(Logger, null);
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
