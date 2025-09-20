using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace VNC;

/// <summary>
/// Client for communicating with the managed analysis service
/// </summary>
public static class ManagedClient
{
    private static readonly HttpClient Http = new HttpClient();
    
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web)
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    /// <summary>
    /// Analyzes text using the managed service OCR endpoint
    /// </summary>
    /// <param name="text">Text to analyze</param>
    /// <param name="baseUrl">Base URL of the managed service</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Analysis result</returns>
    public static async Task<string> AnalyzeTextAsync(string text, string baseUrl, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException("Text cannot be null or empty", nameof(text));

        if (string.IsNullOrWhiteSpace(baseUrl))
            throw new ArgumentException("Base URL cannot be null or empty", nameof(baseUrl));

        var payload = new { text = text };
        var json = JsonSerializer.Serialize(payload, SerializerOptions);
        
        var url = baseUrl.TrimEnd('/') + "/api/analysis/analyze-text";
        using var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json")
        };

        try
        {
            using var response = await Http.SendAsync(request, cancellationToken);
            var responseBody = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                throw new HttpRequestException(
                    $"Managed API error {(int)response.StatusCode} {response.ReasonPhrase}: {responseBody}");
            }

            var result = JsonSerializer.Deserialize<ManagedApiResponse>(responseBody, SerializerOptions);
            return result?.Result ?? string.Empty;
        }
        catch (HttpRequestException)
        {
            throw;
        }
        catch (TaskCanceledException)
        {
            throw new TimeoutException("Request to managed service timed out");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to analyze text with managed service: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Analyzes an image using the managed service image endpoint
    /// </summary>
    /// <param name="base64Image">Base64 encoded image</param>
    /// <param name="baseUrl">Base URL of the managed service</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Analysis result</returns>
    public static async Task<string> AnalyzeImageAsync(string base64Image, string baseUrl, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(base64Image))
            throw new ArgumentException("Base64 image cannot be null or empty", nameof(base64Image));

        if (string.IsNullOrWhiteSpace(baseUrl))
            throw new ArgumentException("Base URL cannot be null or empty", nameof(baseUrl));

        var payload = new { base64image = base64Image };
        var json = JsonSerializer.Serialize(payload, SerializerOptions);
        
        var url = baseUrl.TrimEnd('/') + "/api/analysis/analyze-image";
        using var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json")
        };

        try
        {
            using var response = await Http.SendAsync(request, cancellationToken);
            var responseBody = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                throw new HttpRequestException(
                    $"Managed API error {(int)response.StatusCode} {response.ReasonPhrase}: {responseBody}");
            }

            var result = JsonSerializer.Deserialize<ManagedApiResponse>(responseBody, SerializerOptions);
            return result?.Result ?? string.Empty;
        }
        catch (HttpRequestException)
        {
            throw;
        }
        catch (TaskCanceledException)
        {
            throw new TimeoutException("Request to managed service timed out");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to analyze image with managed service: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Response model for the managed API
    /// </summary>
    private class ManagedApiResponse
    {
        public string Result { get; set; } = string.Empty;
    }
}
