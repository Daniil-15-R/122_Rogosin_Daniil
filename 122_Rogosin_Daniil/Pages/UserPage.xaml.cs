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

namespace _122_Rogosin_Daniil.Pages
{
    /// <summary>
    /// Логика взаимодействия для UserPage.xaml
    /// </summary>
    public partial class UserPage : Page
    {
        public UserPage()
        {
            InitializeComponent();
            var currentUsers = Entities.GetContext().User.ToList();
            ListUser.ItemsSource = currentUsers;
        }

        private void clearFiltersButton_Click_1(object sender, RoutedEventArgs e)
        {
            fioFilterTextBox.Text = "";
            sortComboBox.SelectedIndex = 0;
            onlyAdminCheckBox.IsChecked = false;
        }
        /// <summary>
        /// Обрабатывает изменение текста в поле фильтрации по ФИО
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события изменения текста</param>
        private void fioFilterTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateUsers();
        }

        private void sortComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateUsers();
        }

        private void onlyAdminCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            UpdateUsers();
        }

        private void onlyAdminCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateUsers();
        }
        /// <summary>
        /// Обновляет список пользователей с применением фильтров и сортировки
        /// </summary>
        /// <remarks>
        /// Выполняет фильтрацию по ФИО и роли, а также сортировку по выбранному критерию
        /// </remarks>
        /// <example>
        /// <code>
        /// // Пример применения фильтров:
        /// // 1. Ввод текста в поле "ФИО" фильтрует пользователей по совпадению
        /// // 2. Выбор "По убыванию" сортирует список в обратном порядке
        /// // 3. Установка флажка "Только администраторы" показывает только пользователей с ролью "Admin"
        /// </code>
        /// </example>
        private void UpdateUsers()
        {
            if (!IsInitialized)
            {
                return;
            }
            try
            {
                List<User> currentUsers = Entities.GetContext().User.ToList();

                // Филтрация по фамилии
                if (!string.IsNullOrWhiteSpace(fioFilterTextBox.Text))
                {
                    currentUsers = currentUsers.Where(x => x.FIO.ToLower().Contains(fioFilterTextBox.Text.ToLower())).ToList();
                }

                // Фильтрация по роли
                if (onlyAdminCheckBox.IsChecked.Value)
                {
                    currentUsers = currentUsers.Where(x => x.Role == "Admin").ToList();
                }

                // Сортировка по убыванию/возрастанию
                ListUser.ItemsSource = (sortComboBox.SelectedIndex == 0) ?
                    currentUsers.OrderBy(x => x.FIO).ToList() :
                    currentUsers.OrderByDescending(x => x.FIO).ToList();
            }
            catch (Exception)
            {
                // Обработка ошибок
            }
        }
    }
}