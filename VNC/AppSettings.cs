using System;
using System.IO;
using System.Threading.Tasks;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace VNC;

/// <summary>
/// Application settings model
/// </summary>
public class AppSettings
{
    public string Intelligence { get; set; } = "managed";
    public string Model { get; set; } = "Image";
    public string TogetherApiKey { get; set; } = string.Empty;
    public string SelectedTogetherModel { get; set; } = "google/gemma-3n-E4B-it";
    public string OllamaEndpoint { get; set; } = "http://localhost:11434";
    public string SelectedOllamaModel { get; set; } = string.Empty;
    public string ManagedServiceUrl { get; set; } = "http://localhost:8080";
}

/// <summary>
/// Settings manager for persisting app settings to YAML file in AppData
/// </summary>
public static class SettingsManager
{
    private static readonly string SettingsDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
        "VNCompanion");
    
    private static readonly string SettingsFilePath = Path.Combine(SettingsDirectory, "settings.yml");
    
    private static readonly ISerializer YamlSerializer = new SerializerBuilder()
        .WithNamingConvention(CamelCaseNamingConvention.Instance)
        .Build();
    
    private static readonly IDeserializer YamlDeserializer = new DeserializerBuilder()
        .WithNamingConvention(CamelCaseNamingConvention.Instance)
        .Build();

    /// <summary>
    /// Loads settings from the YAML file, or returns default settings if file doesn't exist
    /// </summary>
    public static async Task<AppSettings> LoadSettingsAsync()
    {
        try
        {
            if (!File.Exists(SettingsFilePath))
            {
                return new AppSettings();
            }

            var yamlContent = await File.ReadAllTextAsync(SettingsFilePath);
            return YamlDeserializer.Deserialize<AppSettings>(yamlContent) ?? new AppSettings();
        }
        catch (Exception)
        {
            return new AppSettings();
        }
    }

    /// <summary>
    /// Saves settings to the YAML file
    /// </summary>
    public static async Task SaveSettingsAsync(AppSettings settings)
    {
        try
        {
            Directory.CreateDirectory(SettingsDirectory);
            
            var yamlContent = YamlSerializer.Serialize(settings);
            await File.WriteAllTextAsync(SettingsFilePath, yamlContent);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to save settings: {ex.Message}", ex);
        }
    }
}
