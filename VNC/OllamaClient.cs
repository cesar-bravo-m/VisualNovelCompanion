using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace VNC;

/// <summary>
/// Client for communicating with local Ollama instance
/// </summary>
public class OllamaClient : IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly string _baseUrl;

    public OllamaClient(string baseUrl = "http://localhost:11434")
    {
        _baseUrl = baseUrl.TrimEnd('/');
        _httpClient = new HttpClient();
        _httpClient.Timeout = TimeSpan.FromMinutes(5);
    }

    /// <summary>
    /// Gets the list of available models from Ollama
    /// </summary>
    /// <returns>List of available model names</returns>
    public async Task<List<string>> GetAvailableModelsAsync()
    {
        try
        {
            var response = await _httpClient.GetAsync($"{_baseUrl}/api/tags");
            
            if (!response.IsSuccessStatusCode)
            {
                throw new HttpRequestException($"Failed to get models: {response.StatusCode}");
            }

            var jsonContent = await response.Content.ReadAsStringAsync();
            
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };
            
            var modelsResponse = JsonSerializer.Deserialize<OllamaModelsResponse>(jsonContent, options);

            var modelNames = new List<string>();
            if (modelsResponse?.Models != null)
            {
                foreach (var model in modelsResponse.Models)
                {
                    if (!string.IsNullOrEmpty(model.Name))
                    {
                        modelNames.Add(model.Name);
                    }
                }
            }

            return modelNames;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to retrieve Ollama models: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Checks if Ollama is available and running
    /// </summary>
    /// <returns>True if Ollama is accessible, false otherwise</returns>
    public async Task<bool> IsAvailableAsync()
    {
        try
        {
            var response = await _httpClient.GetAsync($"{_baseUrl}/api/tags");
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Performs OCR using Ollama with vision-capable models
    /// </summary>
    /// <param name="bitmap">The image to process</param>
    /// <param name="modelName">The Ollama model to use (must support vision)</param>
    /// <returns>Extracted text from the image</returns>
    public async Task<string> ExtractTextFromImageAsync(Bitmap bitmap, string modelName)
    {
        if (bitmap == null)
            throw new ArgumentNullException(nameof(bitmap));
        
        if (string.IsNullOrEmpty(modelName))
            throw new ArgumentException("Model name cannot be empty", nameof(modelName));

        try
        {
            var base64Image = ConvertBitmapToBase64(bitmap);

            var request = new OllamaGenerateRequest
            {
                Model = modelName,
                Prompt = "Extract all text from this image. Return only the text content without any additional commentary or formatting. If there is no text, return 'No text found'.",
                Images = new[] { base64Image },
                Stream = false,
                Options = new OllamaOptions
                {
                    Temperature = 0.1f,
                    TopP = 0.9f
                }
            };

            var requestJson = JsonSerializer.Serialize(request, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            var response = await _httpClient.PostAsync($"{_baseUrl}/api/generate", content);

            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"Ollama request failed ({response.StatusCode}): {errorContent}");
            }

            var responseJson = await response.Content.ReadAsStringAsync();
            var ollamaResponse = JsonSerializer.Deserialize<OllamaGenerateResponse>(responseJson, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            return ollamaResponse?.Response?.Trim() ?? string.Empty;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Ollama OCR failed: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Processes text using Ollama for translation or analysis
    /// </summary>
    /// <param name="text">The extracted text to process</param>
    /// <param name="modelName">The Ollama model to use</param>
    /// <returns>Processed/translated text</returns>
    public async Task<string> ProcessTextAsync(string text, string modelName)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("Text cannot be empty", nameof(text));
        
        if (string.IsNullOrEmpty(modelName))
            throw new ArgumentException("Model name cannot be empty", nameof(modelName));

        try
        {
            var request = new OllamaGenerateRequest
            {
                Model = modelName,
                Prompt = $"Translate the following text following the system prompt closely:\n\n{text}",
                Stream = false,
                Options = new OllamaOptions
                {
                    Temperature = 0.3f,
                    TopP = 0.9f
                }
            };

            var requestJson = JsonSerializer.Serialize(request, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            var response = await _httpClient.PostAsync($"{_baseUrl}/api/generate", content);

            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"Ollama request failed ({response.StatusCode}): {errorContent}");
            }

            var responseJson = await response.Content.ReadAsStringAsync();
            var ollamaResponse = JsonSerializer.Deserialize<OllamaGenerateResponse>(responseJson, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            return ollamaResponse?.Response?.Trim() ?? string.Empty;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Ollama text processing failed: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Converts a bitmap to base64 string for Ollama API
    /// </summary>
    private static string ConvertBitmapToBase64(Bitmap bitmap)
    {
        using var memoryStream = new MemoryStream();
        bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
        var imageBytes = memoryStream.ToArray();
        return Convert.ToBase64String(imageBytes);
    }

    /// <summary>
    /// Auto-detects common Ollama endpoints
    /// </summary>
    /// <returns>First available Ollama endpoint, or null if none found</returns>
    public static async Task<string?> AutoDetectEndpointAsync()
    {
        var commonEndpoints = new[]
        {
            "http://localhost:11434",
            "http://127.0.0.1:11434",
            "http://localhost:11435",
            "http://127.0.0.1:11435"
        };

        foreach (var endpoint in commonEndpoints)
        {
            try
            {
                using var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(3);
                
                var response = await client.GetAsync($"{endpoint}/api/tags");
                if (response.IsSuccessStatusCode)
                {
                    return endpoint;
                }
            }
            catch
            {
            }
        }

        return null;
    }

    public void Dispose()
    {
        _httpClient?.Dispose();
    }
}

/// <summary>
/// Response model for Ollama models list API
/// </summary>
public class OllamaModelsResponse
{
    [JsonPropertyName("models")]
    public OllamaModel[]? Models { get; set; }
}

/// <summary>
/// Ollama model information
/// </summary>
public class OllamaModel
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;
    
    [JsonPropertyName("model")]
    public string Model { get; set; } = string.Empty;
    
    [JsonPropertyName("size")]
    public long Size { get; set; }
    
    [JsonPropertyName("digest")]
    public string Digest { get; set; } = string.Empty;
    
    [JsonPropertyName("modified_at")]
    public string ModifiedAt { get; set; } = string.Empty;
    
    [JsonPropertyName("details")]
    public OllamaModelDetails? Details { get; set; }
}

/// <summary>
/// Ollama model details
/// </summary>
public class OllamaModelDetails
{
    [JsonPropertyName("parent_model")]
    public string ParentModel { get; set; } = string.Empty;
    
    [JsonPropertyName("format")]
    public string Format { get; set; } = string.Empty;
    
    [JsonPropertyName("family")]
    public string Family { get; set; } = string.Empty;
    
    [JsonPropertyName("families")]
    public string[]? Families { get; set; }
    
    [JsonPropertyName("parameter_size")]
    public string ParameterSize { get; set; } = string.Empty;
    
    [JsonPropertyName("quantization_level")]
    public string QuantizationLevel { get; set; } = string.Empty;
}

/// <summary>
/// Request model for Ollama generate API
/// </summary>
public class OllamaGenerateRequest
{
    public string Model { get; set; } = string.Empty;
    public string Prompt { get; set; } = string.Empty;
    public string[]? Images { get; set; }
    public bool Stream { get; set; } = false;
    public OllamaOptions? Options { get; set; }
}

/// <summary>
/// Ollama generation options
/// </summary>
public class OllamaOptions
{
    public float Temperature { get; set; } = 0.8f;
    public float TopP { get; set; } = 0.9f;
    public int TopK { get; set; } = 40;
}

/// <summary>
/// Response model for Ollama generate API
/// </summary>
public class OllamaGenerateResponse
{
    public string Model { get; set; } = string.Empty;
    public string Response { get; set; } = string.Empty;
    public bool Done { get; set; }
    public long TotalDuration { get; set; }
    public long LoadDuration { get; set; }
    public long PromptEvalCount { get; set; }
    public long PromptEvalDuration { get; set; }
    public long EvalCount { get; set; }
    public long EvalDuration { get; set; }
}
