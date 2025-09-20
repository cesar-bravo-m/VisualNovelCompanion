using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace VNC;

public sealed partial class SettingsDialog : ContentDialog
{
    private AppSettings _currentSettings = new AppSettings();
    private bool _isLoading = true;
    private OllamaClient? _ollamaClient;

    public SettingsDialog()
    {
        this.InitializeComponent();
        this.Loaded += SettingsDialog_Loaded;
    }

    private async void SettingsDialog_Loaded(object sender, RoutedEventArgs e)
    {
        await LoadSettingsAsync();
    }

    private async Task LoadSettingsAsync()
    {
        try
        {
            _isLoading = true;
            _currentSettings = await SettingsManager.LoadSettingsAsync();
            
            var intelligenceButton = IntelligenceRadioButtons.Items
                .OfType<RadioButton>()
                .FirstOrDefault(rb => rb.Tag?.ToString() == _currentSettings.Intelligence);
            if (intelligenceButton != null)
            {
                intelligenceButton.IsChecked = true;
            }
            else
            {
                var managedButton = IntelligenceRadioButtons.Items
                    .OfType<RadioButton>()
                    .FirstOrDefault(rb => rb.Tag?.ToString() == "managed");
                if (managedButton != null)
                    managedButton.IsChecked = true;
            }

            if (_currentSettings.Model == "Image")
            {
                ImageRadioButton.IsChecked = true;
            }
            else if (_currentSettings.Model == "OCR")
            {
                OCRRadioButton.IsChecked = true;
            }
            else
            {
                ImageRadioButton.IsChecked = true;
            }

            TogetherApiKeyBox.Password = _currentSettings.TogetherApiKey;
            
            var modelItem = TogetherModelComboBox.Items
                .OfType<ComboBoxItem>()
                .FirstOrDefault(item => item.Tag?.ToString() == _currentSettings.SelectedTogetherModel);
            if (modelItem != null)
            {
                TogetherModelComboBox.SelectedItem = modelItem;
            }
            else
            {
                if (TogetherModelComboBox.Items.Count > 0)
                    TogetherModelComboBox.SelectedIndex = 0;
            }

            OllamaEndpointBox.Text = _currentSettings.OllamaEndpoint;
            
            _ollamaClient = new OllamaClient(_currentSettings.OllamaEndpoint);

            if (_currentSettings.Intelligence == "local")
            {
                await LoadOllamaModelsAsync();
            }

            UpdatePanelVisibility();
        }
        finally
        {
            _isLoading = false;
        }
    }

    private void IntelligenceRadioButtons_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoading) return;

        var selectedButton = IntelligenceRadioButtons.SelectedItem as RadioButton;
        if (selectedButton?.Tag is string intelligence)
        {
            _currentSettings.Intelligence = intelligence;
            
            // For managed mode, force Image mode since OCR is not supported
            if (intelligence == "managed")
            {
                _currentSettings.Model = "Image";
                ImageRadioButton.IsChecked = true;
            }
            
            UpdatePanelVisibility();
        }
    }

    private void TogetherApiKeyBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (_isLoading) return;
        _currentSettings.TogetherApiKey = TogetherApiKeyBox.Password;
    }

    private void TogetherModelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoading) return;

        if (TogetherModelComboBox.SelectedItem is ComboBoxItem selectedItem &&
            selectedItem.Tag is string model)
        {
            _currentSettings.SelectedTogetherModel = model;
        }
    }

    private void OllamaEndpointBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isLoading) return;
        _currentSettings.OllamaEndpoint = OllamaEndpointBox.Text;
        
        // Update Ollama client with new endpoint
        _ollamaClient?.Dispose();
        _ollamaClient = new OllamaClient(_currentSettings.OllamaEndpoint);
        
        // Clear existing models since endpoint changed
        OllamaModelComboBox.Items.Clear();
        OllamaStatusText.Text = "Endpoint changed. Click 'Refresh Models' to load models from the new endpoint.";
        OllamaStatusText.Foreground = App.Current.Resources["SystemControlForegroundBaseMediumBrush"] as Microsoft.UI.Xaml.Media.Brush;
    }

    private void UpdatePanelVisibility()
    {
        // Hide all panels first
        ManagedSettingsPanel.Visibility = Visibility.Collapsed;
        BYOKSettingsPanel.Visibility = Visibility.Collapsed;
        LocalSettingsPanel.Visibility = Visibility.Collapsed;
        InputModePanel.Visibility = Visibility.Collapsed;

        // Show the appropriate panels
        switch (_currentSettings.Intelligence)
        {
            case "managed":
                ManagedSettingsPanel.Visibility = Visibility.Visible;
                // Input mode is hidden for managed - only IMAGE mode is supported
                InputModePanel.Visibility = Visibility.Collapsed;
                break;
            case "BYOK":
                BYOKSettingsPanel.Visibility = Visibility.Visible;
                InputModePanel.Visibility = Visibility.Visible;
                break;
            case "local":
                LocalSettingsPanel.Visibility = Visibility.Visible;
                InputModePanel.Visibility = Visibility.Visible;
                // Load Ollama models when switching to local mode
                if (!_isLoading)
                {
                    _ = LoadOllamaModelsAsync();
                }
                break;
        }
    }

    private async void ContentDialog_PrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        // Get a deferral to allow async operations
        var deferral = args.GetDeferral();
        
        try
        {
            // Update Model setting from radio buttons
            if (ImageRadioButton.IsChecked == true)
            {
                _currentSettings.Model = "Image";
            }
            else if (OCRRadioButton.IsChecked == true)
            {
                _currentSettings.Model = "OCR";
            }

            await SettingsManager.SaveSettingsAsync(_currentSettings);
        }
        catch (Exception ex)
        {
            // Cancel the dialog close and show error
            args.Cancel = true;
            
            var errorDialog = new ContentDialog()
            {
                Title = "Save Error",
                Content = $"Failed to save settings: {ex.Message}",
                CloseButtonText = "OK",
                XamlRoot = this.XamlRoot
            };
            await errorDialog.ShowAsync();
        }
        finally
        {
            deferral.Complete();
        }
    }

    private void ContentDialog_SecondaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        // Cancel - no action needed, dialog will close
    }

    private async void AutoDetectButton_Click(object sender, RoutedEventArgs e)
    {
        ShowOllamaLoading("Auto-detecting Ollama endpoint...");
        
        try
        {
            var detectedEndpoint = await OllamaClient.AutoDetectEndpointAsync();
            
            if (!string.IsNullOrEmpty(detectedEndpoint))
            {
                OllamaEndpointBox.Text = detectedEndpoint;
                OllamaStatusText.Text = $"Found Ollama at {detectedEndpoint}";
                OllamaStatusText.Foreground = App.Current.Resources["SystemControlForegroundAccentBrush"] as Microsoft.UI.Xaml.Media.Brush;
            }
            else
            {
                OllamaStatusText.Text = "No Ollama instance found on common ports";
                OllamaStatusText.Foreground = App.Current.Resources["SystemControlErrorTextForegroundBrush"] as Microsoft.UI.Xaml.Media.Brush;
            }
        }
        catch (Exception ex)
        {
            OllamaStatusText.Text = $"Auto-detect failed: {ex.Message}";
            OllamaStatusText.Foreground = App.Current.Resources["SystemControlErrorTextForegroundBrush"] as Microsoft.UI.Xaml.Media.Brush;
        }
        finally
        {
            HideOllamaLoading();
        }
    }

    private async void RefreshModelsButton_Click(object sender, RoutedEventArgs e)
    {
        await LoadOllamaModelsAsync();
    }

    private void OllamaModelComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoading) return;

        if (OllamaModelComboBox.SelectedItem is ComboBoxItem selectedItem &&
            selectedItem.Tag is string model)
        {
            _currentSettings.SelectedOllamaModel = model;
        }
    }

    private async Task LoadOllamaModelsAsync()
    {
        ShowOllamaLoading("Loading available models...");
        
        try
        {
            if (_ollamaClient == null)
            {
                _ollamaClient = new OllamaClient(_currentSettings.OllamaEndpoint);
            }

            var isAvailable = await _ollamaClient.IsAvailableAsync();
            if (!isAvailable)
            {
                OllamaStatusText.Text = "Cannot connect to Ollama. Check endpoint and ensure Ollama is running.";
                OllamaStatusText.Foreground = App.Current.Resources["SystemControlErrorTextForegroundBrush"] as Microsoft.UI.Xaml.Media.Brush;
                return;
            }

            var models = await _ollamaClient.GetAvailableModelsAsync();
            
            OllamaModelComboBox.Items.Clear();
            
            if (models.Count == 0)
            {
                OllamaStatusText.Text = "No models found. Install models using 'ollama pull <model_name>'";
                OllamaStatusText.Foreground = App.Current.Resources["SystemControlErrorTextForegroundBrush"] as Microsoft.UI.Xaml.Media.Brush;
                return;
            }

            foreach (var model in models)
            {
                var item = new ComboBoxItem
                {
                    Content = model,
                    Tag = model
                };
                OllamaModelComboBox.Items.Add(item);
            }

            // Select the previously selected model if available
            if (!string.IsNullOrEmpty(_currentSettings.SelectedOllamaModel))
            {
                var selectedItem = OllamaModelComboBox.Items
                    .OfType<ComboBoxItem>()
                    .FirstOrDefault(item => item.Tag?.ToString() == _currentSettings.SelectedOllamaModel);
                
                if (selectedItem != null)
                {
                    OllamaModelComboBox.SelectedItem = selectedItem;
                }
            }

            // If no selection, default to first model
            if (OllamaModelComboBox.SelectedIndex == -1 && OllamaModelComboBox.Items.Count > 0)
            {
                OllamaModelComboBox.SelectedIndex = 0;
            }

            OllamaStatusText.Text = $"Found {models.Count} model(s)";
            OllamaStatusText.Foreground = App.Current.Resources["SystemControlForegroundAccentBrush"] as Microsoft.UI.Xaml.Media.Brush;
        }
        catch (Exception ex)
        {
            OllamaStatusText.Text = $"Failed to load models: {ex.Message}";
            OllamaStatusText.Foreground = App.Current.Resources["SystemControlErrorTextForegroundBrush"] as Microsoft.UI.Xaml.Media.Brush;
        }
        finally
        {
            HideOllamaLoading();
        }
    }

    private void ShowOllamaLoading(string message)
    {
        // Show loading overlay
        OllamaLoadingText.Text = message;
        OllamaProgressRing.IsActive = true;
        OllamaLoadingOverlay.Visibility = Visibility.Visible;
        
        // Disable intelligence selection to prevent race conditions
        IntelligenceRadioButtons.IsEnabled = false;
        
        // Disable primary button
        this.IsPrimaryButtonEnabled = false;
    }

    private void HideOllamaLoading()
    {
        // Hide loading overlay
        OllamaProgressRing.IsActive = false;
        OllamaLoadingOverlay.Visibility = Visibility.Collapsed;
        
        // Re-enable intelligence selection
        IntelligenceRadioButtons.IsEnabled = true;
        
        // Re-enable primary button
        this.IsPrimaryButtonEnabled = true;
    }
}
