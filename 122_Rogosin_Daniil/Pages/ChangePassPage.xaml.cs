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
using System.Security.Cryptography;

namespace _122_Rogosin_Daniil.Pages
{
    /// <summary>
    /// Логика взаимодействия для ChangePassPage.xaml
    /// </summary>
    public partial class ChangePassPage : Page
    {
        public ChangePassPage()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Обрабатывает событие нажатия кнопки сохранения нового пароля
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>
        /// <remarks>
        /// Выполняет проверку текущего пароля, валидацию нового пароля и сохраняет изменения
        /// </remarks>
        /// <example>
        /// <code>
        /// // Пример использования:
        /// // 1. Введите логин и текущий пароль
        /// // 2. Введите новый пароль, соответствующий требованиям безопасности
        /// // 3. Подтвердите новый пароль
        /// // 4. Нажмите кнопку "Сохранить"
        /// </code>
        /// </example>
        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(CurrentPasswordBox.Password) ||
                string.IsNullOrEmpty(NewPasswordBox.Password) ||
                string.IsNullOrEmpty(ConfirmPasswordBox.Password) ||
                string.IsNullOrEmpty(TbLogin.Text))
            {
                MessageBox.Show("Все поля обязательны к заполнению!");
                return;
            }

            string login = TbLogin.Text.Trim();
            string hashedPass = GetHash(CurrentPasswordBox.Password);

            var user = Entities.GetContext().User
                .FirstOrDefault(u => u.Login == login && u.Password == hashedPass);

            if (user == null)
            {
                MessageBox.Show("Текущий пароль/Логин неверный!");
                return;
            }

            if (NewPasswordBox.Password != ConfirmPasswordBox.Password)
            {
                MessageBox.Show("Новый пароль и подтверждение не совпадают!");
                return;
            }

            bool hasUpper = false;
            bool hasLower = false;
            bool hasDigit = false;

            foreach (char c in NewPasswordBox.Password)
            {
                if (char.IsUpper(c)) hasUpper = true;
                if (char.IsLower(c)) hasLower = true;
                if (char.IsDigit(c)) hasDigit = true;
            }

            if (NewPasswordBox.Password.Length < 6)
            {
                MessageBox.Show("Пароль должен содержать минимум 6 символов!");
                return;
            }

            if (!hasDigit)
            {
                MessageBox.Show("Пароль должен содержать цифры!");
                return;
            }

            if (!hasUpper)
            {
                MessageBox.Show("Пароль должен содержать заглавные буквы!");
                return;
            }

            try
            {
                user.Password = GetHash(NewPasswordBox.Password);
                Entities.GetContext().SaveChanges();
                MessageBox.Show("Пароль успешно изменен!");
                NavigationService?.Navigate(new AuthPage());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении: {ex.Message}");
            }
        }
        /// <summary>
        /// Вычисляет хеш-сумму пароля с использованием алгоритма SHA1
        /// </summary>
        /// <param name="password">Пароль для хеширования</param>
        /// <returns>Хеш-сумма пароля в шестнадцатеричном формате</returns>
        private string GetHash(string password)
        {
            using (var hash = SHA1.Create())
            {
                return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(password))
                    .Select(x => x.ToString("X2")));
            }
        }
    }
}