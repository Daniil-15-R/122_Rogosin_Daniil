using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.IO;

namespace _122_Rogosin_Daniil.Pages
{
    /// <summary>
    /// Логика взаимодействия для AddUserPage.xaml
    /// </summary>
    public partial class AddUserPage : Page
    {
        private User _currentUser = new User();
        private string _selectedPhotoPath = string.Empty;

        public AddUserPage(User selectedUser)
        {
            InitializeComponent();

            if (selectedUser != null)
                _currentUser = selectedUser;

            DataContext = _currentUser;
            cmbRole.SelectedIndex = 0;
        }

        /// <summary>
        /// Обработчик сохранения данных пользователя
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>
        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder errors = new StringBuilder();

            if (string.IsNullOrWhiteSpace(_currentUser.Login))
                errors.AppendLine("Укажите логин!");
            if (string.IsNullOrWhiteSpace(_currentUser.Password))
                errors.AppendLine("Укажите пароль!");
            if ((_currentUser.Role == null) || (cmbRole.Text == ""))
                errors.AppendLine("Выберите роль!");
            else
                _currentUser.Role = cmbRole.Text;
            if (string.IsNullOrWhiteSpace(_currentUser.FIO))
                errors.AppendLine("Укажите ФИО");

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }

            if (!string.IsNullOrWhiteSpace(_selectedPhotoPath))
            {
                try
                {
                    string photosDirectory = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "UserPhotos");
                    if (!Directory.Exists(photosDirectory))
                        Directory.CreateDirectory(photosDirectory);

                    string fileName = System.IO.Path.GetFileName(_selectedPhotoPath);
                    string destinationPath = System.IO.Path.Combine(photosDirectory, fileName);

                    if (!File.Exists(destinationPath) || !_selectedPhotoPath.Equals(destinationPath, StringComparison.OrdinalIgnoreCase))
                    {
                        File.Copy(_selectedPhotoPath, destinationPath, true);
                    }

                    _currentUser.Photo = $"UserPhotos/{fileName}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сохранении фотографии: {ex.Message}");
                    return;
                }
            }

            if (_currentUser.ID == 0)
                Entities.GetContext().User.Add(_currentUser);

            try
            {
                Entities.GetContext().SaveChanges();
                MessageBox.Show("Данные успешно сохранены!");
                NavigationService?.GoBack();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        /// <summary>
        /// Обработчик очистки полей формы
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>
        private void ButtonClean_Click(object sender, RoutedEventArgs e)
        {
            TBLogin.Text = "";
            TBPass.Text = "";
            cmbRole.SelectedItem = null;
            TBFio.Text = "";
            TBPhoto.Text = "";
            _selectedPhotoPath = string.Empty;
            ImagePreview.Source = null;
        }

        /// <summary>
        /// Обработчик выбора фотографии через диалоговое окно
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>
        private void ButtonSelectPhoto_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Image files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|All files (*.*)|*.*";
            openFileDialog.Title = "Выберите фотографию пользователя";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    _selectedPhotoPath = openFileDialog.FileName;
                    TBPhoto.Text = System.IO.Path.GetFileName(_selectedPhotoPath);

                    // Показываем превью изображения
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(_selectedPhotoPath);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    ImagePreview.Source = bitmap;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при загрузке изображения: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Обработчик удаления выбранной фотографии
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>
        private void ButtonRemovePhoto_Click(object sender, RoutedEventArgs e)
        {
            _selectedPhotoPath = string.Empty;
            TBPhoto.Text = "";
            ImagePreview.Source = null;
        }
    }
}