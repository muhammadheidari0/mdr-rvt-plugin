using System;
using System.Windows;
using System.Windows.Controls;
using Mdr.Revit.Infra.Config;

namespace Mdr.Revit.Addin.UI
{
    public sealed class SettingsWindow
    {
        private readonly PluginConfig _config;

        public SettingsWindow(PluginConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            ApiBaseUrl = _config.ApiBaseUrl ?? string.Empty;
            NativeFormat = _config.Publish?.NativeFormat ?? "dwg";
        }

        public string ApiBaseUrl { get; private set; } = string.Empty;

        public string NativeFormat { get; private set; } = "dwg";

        public bool? ShowDialog()
        {
            Window window = new Window
            {
                Title = "Plugin Settings",
                Width = 640,
                Height = 260,
                MinWidth = 520,
                MinHeight = 220,
                ResizeMode = ResizeMode.CanMinimize,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
            };

            Grid root = new Grid
            {
                Margin = new Thickness(16),
            };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            TextBlock info = new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 10),
                FontWeight = FontWeights.SemiBold,
                Text = "Edit runtime plugin settings.",
            };
            root.Children.Add(info);
            Grid.SetRow(info, 0);

            StackPanel apiRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 8),
            };
            apiRow.Children.Add(new TextBlock
            {
                Width = 140,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "API Base URL",
            });
            TextBox apiTextBox = new TextBox
            {
                Width = 430,
                Text = ApiBaseUrl,
            };
            apiRow.Children.Add(apiTextBox);
            root.Children.Add(apiRow);
            Grid.SetRow(apiRow, 1);

            StackPanel nativeRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 12),
            };
            nativeRow.Children.Add(new TextBlock
            {
                Width = 140,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Native Format",
            });
            TextBox nativeTextBox = new TextBox
            {
                Width = 140,
                Text = NativeFormat,
            };
            nativeRow.Children.Add(nativeTextBox);
            root.Children.Add(nativeRow);
            Grid.SetRow(nativeRow, 2);

            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
            };
            Button saveButton = new Button
            {
                Width = 90,
                Margin = new Thickness(0, 0, 8, 0),
                IsDefault = true,
                Content = "Save",
            };
            Button cancelButton = new Button
            {
                Width = 90,
                IsCancel = true,
                Content = "Cancel",
            };
            buttons.Children.Add(saveButton);
            buttons.Children.Add(cancelButton);
            root.Children.Add(buttons);
            Grid.SetRow(buttons, 3);

            saveButton.Click += (_, _) =>
            {
                string url = apiTextBox.Text?.Trim() ?? string.Empty;
                string format = nativeTextBox.Text?.Trim() ?? string.Empty;
                if (!TryValidateValues(url, format, out string errorMessage))
                {
                    MessageBox.Show(
                        errorMessage,
                        "Invalid Settings",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                ApiBaseUrl = url;
                NativeFormat = format.ToLowerInvariant();
                _config.ApiBaseUrl = ApiBaseUrl;
                if (_config.Publish == null)
                {
                    _config.Publish = new PublishPluginConfig();
                }

                _config.Publish.NativeFormat = NativeFormat;

                window.DialogResult = true;
                window.Close();
            };

            cancelButton.Click += (_, _) =>
            {
                window.DialogResult = false;
                window.Close();
            };

            window.Content = root;
            return window.ShowDialog();
        }

        internal static bool TryValidateValues(string apiBaseUrl, string nativeFormat, out string errorMessage)
        {
            if (string.IsNullOrWhiteSpace(apiBaseUrl))
            {
                errorMessage = "API Base URL is required.";
                return false;
            }

            if (!Uri.TryCreate(apiBaseUrl, UriKind.Absolute, out Uri? uri))
            {
                errorMessage = "API Base URL is not a valid absolute URL.";
                return false;
            }

            if (!string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                errorMessage = "API Base URL must use http or https.";
                return false;
            }

            if (!string.Equals(nativeFormat, "dwg", StringComparison.OrdinalIgnoreCase))
            {
                errorMessage = "Native format must be 'dwg' in this release.";
                return false;
            }

            errorMessage = string.Empty;
            return true;
        }
    }
}
