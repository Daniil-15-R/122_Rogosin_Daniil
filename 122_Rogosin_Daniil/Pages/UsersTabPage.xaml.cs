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
using System.Data.Entity;

namespace _122_Rogosin_Daniil.Pages
{
    /// <summary>
    /// Логика взаимодействия для UsersTabPage.xaml
    /// </summary>
    public partial class UsersTabPage : Page
    {
        public UsersTabPage()
        {
            InitializeComponent();
            DataGridUser.ItemsSource = Entities.GetContext().User.ToList();
            this.IsVisibleChanged += Page_IsVisibleChanged;
        }
        /// <summary>
        /// Обрабатывает изменение видимости страницы для обновления данных
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события изменения видимости</param>
        /// <remarks>
        /// При повторном отображении страницы обновляет данные в DataGrid
        /// </remarks>
        private void Page_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                Entities.GetContext().ChangeTracker.Entries().ToList().ForEach(x => x.Reload());
                DataGridUser.ItemsSource = Entities.GetContext().User.ToList();
            }
        }

        private void ButtonAdd_Click(object sender, RoutedEventArgs e)
        {
            NavigationService?.Navigate(new AddUserPage(null));
        }
        /// <summary>
        /// Обрабатывает нажатие кнопки удаления выбранных пользователей
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>
        /// <remarks>
        /// Выполняет проверку выбора пользователей, подтверждение удаления и удаляет выбранных пользователей из базы данных
        /// </remarks>
        /// <exception cref="Exception">Может вызвать исключение при ошибке удаления из БД</exception>
        private void ButtonDel_Click(object sender, RoutedEventArgs e)
        {
            var usersForRemoving = DataGridUser.SelectedItems.Cast<User>().ToList();

            if (usersForRemoving.Count == 0)
            {
                MessageBox.Show("Выберите пользователей для удаления!", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show($"Вы точно хотите удалить записи в количестве {usersForRemoving.Count()} элементов?", "Внимание",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    // Альтернатива RemoveRange - удаление каждого элемента по отдельности
                    foreach (var user in usersForRemoving)
                    {
                        Entities.GetContext().User.Remove(user);
                    }

                    Entities.GetContext().SaveChanges();
                    MessageBox.Show("Данные успешно удалены!");

                    DataGridUser.ItemsSource = Entities.GetContext().User.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        private void ButtonEdit_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Pages.AddUserPage((sender as Button).DataContext as User));
        }
    }
}