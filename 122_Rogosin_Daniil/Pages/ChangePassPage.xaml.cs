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

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            // Проверка заполнения всех полей
            if (string.IsNullOrEmpty(CurrentPasswordBox.Password) ||
                string.IsNullOrEmpty(NewPasswordBox.Password) ||
                string.IsNullOrEmpty(ConfirmPasswordBox.Password) ||
                string.IsNullOrEmpty(TbLogin.Text))
            {
                MessageBox.Show("Все поля обязательны к заполнению!");
                return;
            }

            // Проверка правильности введенных данных аккаунта
            string hashedPass = GetHash(CurrentPasswordBox.Password);
            var user = Entities.GetContext().User
                .FirstOrDefault(u => u.Login == TbLogin.Text && u.Password == hashedPass);

            if (user == null)
            {
                MessageBox.Show("Текущий пароль/Логин неверный!");
                return;
            }

            // Проверка совпадения нового пароля и подтверждения
            if (NewPasswordBox.Password != ConfirmPasswordBox.Password)
            {
                MessageBox.Show("Новый пароль и подтверждение не совпадают!");
                return;
            }

            // Проверка корректности нового пароля (аналогично регистрации)
            bool en = true;
            bool number = false;

            foreach (char c in NewPasswordBox.Password)
            {
                if (c >= 'a' && c <= 'z') en = false;
                if (c >= '0' && c <= '9') number = true;
            }

            if (!number || en || NewPasswordBox.Password.Length < 6)
            {
                MessageBox.Show("Пароль должен содержать минимум 6 символов, включая цифры и заглавные буквы!");
                return;
            }

            // Сохранение нового пароля
            if (en && number)
            {
                user.Password = GetHash(NewPasswordBox.Password);
                Entities.GetContext().SaveChanges();
                MessageBox.Show("Пароль успешно изменен!");
                NavigationService?.Navigate(new AuthPage());
            }
        }

        // Метод для хеширования пароля
        private string GetHash(string password)
        {
            using (var sha256 = SHA256.Create())
            {
                var hashedBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
                return BitConverter.ToString(hashedBytes).Replace("-", "").ToLower();
            }
        }
    }
}