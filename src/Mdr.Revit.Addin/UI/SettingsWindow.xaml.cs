using System;
using System.Windows;
using System.Windows.Controls;
using Mdr.Revit.Infra.Config;
using Mdr.Revit.Infra.Security;

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
            ExcelDefaultDirectory = _config.Excel?.DefaultDirectory ?? "%LocalAppData%/MDR/RevitPlugin/excel";
            ExcelDefaultWorksheetName = _config.Excel?.DefaultWorksheetName ?? string.Empty;
        }

        public string ApiBaseUrl { get; private set; } = string.Empty;

        public string NativeFormat { get; private set; } = "dwg";

        public string ExcelDefaultDirectory { get; private set; } = "%LocalAppData%/MDR/RevitPlugin/excel";

        public string ExcelDefaultWorksheetName { get; private set; } = string.Empty;

        public string ExcelImportPassword { get; private set; } = string.Empty;

        public string ExcelImportPasswordConfirm { get; private set; } = string.Empty;

        public bool? ShowDialog()
        {
            Window window = new Window
            {
                Title = "Plugin Settings",
                Width = 640,
                Height = 420,
                MinWidth = 520,
                MinHeight = 360,
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

            StackPanel excelDirRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 8),
            };
            excelDirRow.Children.Add(new TextBlock
            {
                Width = 140,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Excel Directory",
            });
            TextBox excelDirTextBox = new TextBox
            {
                Width = 430,
                Text = ExcelDefaultDirectory,
            };
            excelDirRow.Children.Add(excelDirTextBox);
            root.Children.Add(excelDirRow);
            Grid.SetRow(excelDirRow, 3);

            StackPanel excelSheetRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 8),
            };
            excelSheetRow.Children.Add(new TextBlock
            {
                Width = 140,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Excel Worksheet",
            });
            TextBox excelSheetTextBox = new TextBox
            {
                Width = 220,
                Text = ExcelDefaultWorksheetName,
            };
            excelSheetRow.Children.Add(excelSheetTextBox);
            root.Children.Add(excelSheetRow);
            Grid.SetRow(excelSheetRow, 4);

            StackPanel excelPasswordRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 12),
            };
            excelPasswordRow.Children.Add(new TextBlock
            {
                Width = 140,
                VerticalAlignment = VerticalAlignment.Center,
                Text = "Excel Password",
            });
            PasswordBox excelPasswordBox = new PasswordBox
            {
                Width = 190,
                Margin = new Thickness(0, 0, 8, 0),
            };
            PasswordBox excelPasswordConfirmBox = new PasswordBox
            {
                Width = 190,
            };
            excelPasswordRow.Children.Add(excelPasswordBox);
            excelPasswordRow.Children.Add(excelPasswordConfirmBox);
            root.Children.Add(excelPasswordRow);
            Grid.SetRow(excelPasswordRow, 5);

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
            Grid.SetRow(buttons, 6);

            saveButton.Click += (_, _) =>
            {
                string url = apiTextBox.Text?.Trim() ?? string.Empty;
                string format = nativeTextBox.Text?.Trim() ?? string.Empty;
                string excelDirectory = excelDirTextBox.Text?.Trim() ?? string.Empty;
                string excelWorksheet = excelSheetTextBox.Text?.Trim() ?? string.Empty;
                string excelPassword = excelPasswordBox.Password ?? string.Empty;
                string excelPasswordConfirm = excelPasswordConfirmBox.Password ?? string.Empty;
                if (!TryValidateValues(
                        url,
                        format,
                        excelDirectory,
                        excelWorksheet,
                        excelPassword,
                        excelPasswordConfirm,
                        out string errorMessage))
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
                ExcelDefaultDirectory = excelDirectory;
                ExcelDefaultWorksheetName = excelWorksheet;
                ExcelImportPassword = excelPassword;
                ExcelImportPasswordConfirm = excelPasswordConfirm;
                _config.ApiBaseUrl = ApiBaseUrl;
                if (_config.Publish == null)
                {
                    _config.Publish = new PublishPluginConfig();
                }

                _config.Publish.NativeFormat = NativeFormat;
                if (_config.Excel == null)
                {
                    _config.Excel = new ExcelPluginConfig();
                }

                _config.Excel.DefaultDirectory = ExcelDefaultDirectory;
                _config.Excel.DefaultWorksheetName = ExcelDefaultWorksheetName;
                if (!string.IsNullOrWhiteSpace(ExcelImportPassword))
                {
                    new ExcelImportPasswordService().ConfigurePassword(
                        _config,
                        ExcelImportPassword,
                        ExcelImportPasswordConfirm);
                }

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
            return TryValidateValues(
                apiBaseUrl,
                nativeFormat,
                "%LocalAppData%/MDR/RevitPlugin/excel",
                string.Empty,
                string.Empty,
                string.Empty,
                out errorMessage);
        }

        internal static bool TryValidateValues(
            string apiBaseUrl,
            string nativeFormat,
            string excelDefaultDirectory,
            string excelDefaultWorksheetName,
            string excelImportPassword,
            string excelImportPasswordConfirm,
            out string errorMessage)
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

            if (string.IsNullOrWhiteSpace(excelDefaultDirectory))
            {
                errorMessage = "Excel default directory is required.";
                return false;
            }

            string worksheet = excelDefaultWorksheetName ?? string.Empty;
            if (worksheet.IndexOfAny(new[] { '\\', '/', '?', '*', '[', ']', ':' }) >= 0)
            {
                errorMessage = "Excel worksheet name contains invalid characters.";
                return false;
            }

            if (worksheet.Length > 31)
            {
                errorMessage = "Excel worksheet name cannot exceed 31 characters.";
                return false;
            }

            bool hasPassword = !string.IsNullOrWhiteSpace(excelImportPassword) ||
                               !string.IsNullOrWhiteSpace(excelImportPasswordConfirm);
            if (hasPassword)
            {
                if (excelImportPassword == null || excelImportPassword.Length < 6 || excelImportPassword.Length > 64)
                {
                    errorMessage = "Excel import password length must be between 6 and 64 characters.";
                    return false;
                }

                if (!string.Equals(excelImportPassword, excelImportPasswordConfirm, StringComparison.Ordinal))
                {
                    errorMessage = "Excel import password and confirmation do not match.";
                    return false;
                }
            }

            errorMessage = string.Empty;
            return true;
        }
    }
}
