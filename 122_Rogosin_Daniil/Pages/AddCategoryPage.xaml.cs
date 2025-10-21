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
    /// Логика взаимодействия для AddCategoryPage.xaml
    /// </summary>
    public partial class AddCategoryPage : Page
    {
        private Category _currentCategory = new Category();

        public AddCategoryPage(Category selectedCategory)
        {
            InitializeComponent();

            if (selectedCategory != null)
                _currentCategory = selectedCategory;

            DataContext = _currentCategory;
        }
        /// <summary>
        /// Обрабатывает событие нажатия кнопки сохранения категории
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>
        /// <remarks>
        /// Выполняет валидацию данных и сохраняет категорию в базу данных
        /// </remarks>
        private void ButtonSaveCategory_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder errors = new StringBuilder();

            if (string.IsNullOrWhiteSpace(_currentCategory.Name))
                errors.AppendLine("Укажите название категории!");

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }

            if (_currentCategory.ID == 0)
                Entities.GetContext().Category.Add(_currentCategory);

            try
            {
                Entities.GetContext().SaveChanges();
                MessageBox.Show("Данные успешно сохранены!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
        /// <summary>
        /// Обрабатывает событие нажатия кнопки очистки поля ввода
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>
        private void ButtonClean_Click(object sender, RoutedEventArgs e)
        {
            TBCategoryName.Text = "";
        }
    }
}