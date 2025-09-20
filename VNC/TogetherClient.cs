using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VNC;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

public class TogetherClient
{
    private sealed class ChatCompletionsRequest
    {
        [JsonPropertyName("model")] public string Model { get; set; } = default!;
        [JsonPropertyName("messages")] public ChatMessage[] Messages { get; set; } = default!;
    }

    private sealed class ChatMessage
    {
        [JsonPropertyName("role")] public string Role { get; set; } = default!;
        [JsonPropertyName("content")] public object Content { get; set; } = default!;
    }

    private sealed class ChatCompletionsResponse
    {
        [JsonPropertyName("choices")] public Choice[]? Choices { get; set; }
    }

    private sealed class Choice
    {
        [JsonPropertyName("message")] public ChoiceMessage? Message { get; set; }
    }

    private sealed class ChoiceMessage
    {
        [JsonPropertyName("role")] public string? Role { get; set; }
        [JsonPropertyName("content")] public string? Content { get; set; }
    }
    private static readonly HttpClient Http = new()
    {
        Timeout = TimeSpan.FromSeconds(60)
    };
    public static async Task<String> AnalyzeScreenshotAsync(
        string base64Image,
        string apiKey = "",
        string modelName = "google/gemma-3n-E4B-it",
        CancellationToken cancellationToken = default)
    {
        return await AnalyzeContentAsync(base64Image, null, apiKey, modelName, cancellationToken);
    }

    public static async Task<String> AnalyzeTextAsync(
        string extractedText,
        string apiKey = "",
        string modelName = "google/gemma-3n-E4B-it",
        CancellationToken cancellationToken = default)
    {
        return await AnalyzeContentAsync(null, extractedText, apiKey, modelName, cancellationToken);
    }

    private static async Task<String> AnalyzeContentAsync(
        string? base64Image,
        string? extractedText,
        string apiKey,
        string modelName,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
            throw new InvalidOperationException("Together API key is required.");

        ChatMessage userMessage;
        
        if (!string.IsNullOrWhiteSpace(base64Image))
        {
            userMessage = new ChatMessage
            {
                Role = "user",
                Content = new object[]
                {
                    new
                    {
                        type = "text",
                        text = $"Translate the text in this image following the system prompt closely."
                    },
                    new
                    {
                        type = "image_url",
                        image_url = new {
                            url = base64Image
                        }
                    }
                }
            };
        }
        else if (!string.IsNullOrWhiteSpace(extractedText))
        {
            userMessage = new ChatMessage
            {
                Role = "user",
                Content = $"Translate the following text following the system prompt closely:\n\n{extractedText}"
            };
        }
        else
        {
            throw new InvalidOperationException("Either base64Image or extractedText must be provided.");
        }

        var payload = new ChatCompletionsRequest
        {
            Model = modelName,
            Messages = new[]
            {
                new ChatMessage
                {
                    Role = "system",
                    Content = SystemPrompt
                },
                userMessage
            }
        };

        var json = JsonSerializer.Serialize(payload, SerializerOptions);
        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.together.xyz/v1/chat/completions")
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json")
        };
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        using var resp = await Http.SendAsync(req, cancellationToken);
        var body = await resp.Content.ReadAsStringAsync(cancellationToken);

        if (!resp.IsSuccessStatusCode)
        {
            throw new HttpRequestException(
                $"Together API error {(int)resp.StatusCode} {resp.ReasonPhrase}: {body}");
        }

        var parsed = JsonSerializer.Deserialize<ChatCompletionsResponse>(body, SerializerOptions);
        var text = parsed?.Choices?[0]?.Message?.Content;
        return text ?? string.Empty;
    }
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    private const string SystemPrompt =
@"You must ONLY return the original text, translations and breakdowns of the text in the attached image.

If the text is in Chinese, Japanese, or Korean, provide a full English translation of the text, and a breakdown of each word in this format:
<word>: (<pronunciation>) <English meaning>
Apply this breakdown line by line for the input text.

DO include the original text in the output, enclosed by <original_text> and </original_text>.
Do NOT include any additional commentary, explanations, or meta-text.
Do NOT prepend phrases like ""The translation is..."". Output must contain ONLY the translation and the breakdowns.

Example:
Input: 今天天气好热
Output:
""Today's weather is very hot""
<original_text>
今天天气好热
</original_text>
今天: (jīn tiān) Today
天气: (tiān qì) weather
好: (hǎo) very
热: (rè) hot
";
}

