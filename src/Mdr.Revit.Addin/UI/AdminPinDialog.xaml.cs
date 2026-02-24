using System;
using System.Windows;
using System.Windows.Controls;

namespace Mdr.Revit.Addin.UI
{
    public sealed class AdminPinDialog
    {
        private readonly AdminPinDialogMode _mode;

        public AdminPinDialog(AdminPinDialogMode mode)
        {
            _mode = mode;
        }

        public string Pin { get; private set; } = string.Empty;

        public string ConfirmPin { get; private set; } = string.Empty;

        public bool? ShowDialog()
        {
            Window window = new Window
            {
                Width = 420,
                Height = _mode == AdminPinDialogMode.Setup ? 240 : 190,
                MinWidth = 380,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Title = _mode == AdminPinDialogMode.Setup ? "Set Admin PIN" : "Admin PIN Required",
            };

            Grid root = new Grid
            {
                Margin = new Thickness(16),
            };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            TextBlock description = new TextBlock
            {
                Margin = new Thickness(0, 0, 0, 12),
                Text = _mode == AdminPinDialogMode.Setup
                    ? "Create an admin PIN to protect plugin settings."
                    : "Enter admin PIN to unlock plugin settings.",
            };
            root.Children.Add(description);
            Grid.SetRow(description, 0);

            StackPanel pinPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 8),
            };
            pinPanel.Children.Add(new TextBlock
            {
                Width = 130,
                VerticalAlignment = VerticalAlignment.Center,
                Text = _mode == AdminPinDialogMode.Setup ? "New PIN" : "PIN",
            });
            PasswordBox pinBox = new PasswordBox
            {
                Width = 220,
            };
            pinPanel.Children.Add(pinBox);
            root.Children.Add(pinPanel);
            Grid.SetRow(pinPanel, 1);

            PasswordBox confirmBox = new PasswordBox
            {
                Width = 220,
            };
            if (_mode == AdminPinDialogMode.Setup)
            {
                StackPanel confirmPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 0, 0, 12),
                };
                confirmPanel.Children.Add(new TextBlock
                {
                    Width = 130,
                    VerticalAlignment = VerticalAlignment.Center,
                    Text = "Confirm PIN",
                });
                confirmPanel.Children.Add(confirmBox);
                root.Children.Add(confirmPanel);
                Grid.SetRow(confirmPanel, 2);
            }

            StackPanel buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
            };
            Button okButton = new Button
            {
                Width = 90,
                Margin = new Thickness(0, 0, 8, 0),
                IsDefault = true,
                Content = "OK",
            };
            Button cancelButton = new Button
            {
                Width = 90,
                IsCancel = true,
                Content = "Cancel",
            };
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            root.Children.Add(buttonPanel);
            Grid.SetRow(buttonPanel, 3);

            okButton.Click += (_, _) =>
            {
                string enteredPin = pinBox.Password ?? string.Empty;
                if (string.IsNullOrWhiteSpace(enteredPin))
                {
                    MessageBox.Show(
                        "PIN is required.",
                        window.Title,
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                Pin = enteredPin;
                ConfirmPin = _mode == AdminPinDialogMode.Setup
                    ? (confirmBox.Password ?? string.Empty)
                    : string.Empty;

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
    }
}
