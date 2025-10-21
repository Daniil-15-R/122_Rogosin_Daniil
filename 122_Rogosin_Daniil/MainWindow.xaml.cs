using System;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Controls;

namespace _122_Rogosin_Daniil
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Обработчик события загрузки окна - запускает таймер и открывает страницу авторизации
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var timer = new DispatcherTimer();
            timer.Interval = new TimeSpan(0, 0, 1);
            timer.IsEnabled = true;
            timer.Tick += (o, t) => { DateTimeNow.Text = DateTime.Now.ToString(); };
            timer.Start();

            MainFrame.Navigate(new Pages.AuthPage());
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            if (MainFrame.CanGoBack)
            {
                MainFrame.GoBack();
            }
        }
        /// <summary>
        /// Обработчик переключения на светлую тему оформления
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Данные события</param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите закрыть окно?", "Подтверждение",
                MessageBoxButton.YesNo) == MessageBoxResult.No)
            {
                e.Cancel = true;
            }
            else
            {
                e.Cancel = false;
            }
        }

        private void LightTheme_Click(object sender, RoutedEventArgs e)
        {
            ChangeTheme("Dictionary.xaml");
        }

        private void DarkTheme_Click(object sender, RoutedEventArgs e)
        {
            ChangeTheme("DictionaryDark.xaml");
        }

        private void NatureTheme_Click(object sender, RoutedEventArgs e)
        {
            ChangeTheme("DictionaryNature.xaml");
        }

        private void ChangeTheme(string dictionaryName)
        {
            try
            {
                var uri = new Uri(dictionaryName, UriKind.Relative);
                ResourceDictionary resourceDict = Application.LoadComponent(uri) as ResourceDictionary;
                Application.Current.Resources.Clear();
                Application.Current.Resources.MergedDictionaries.Add(resourceDict);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при смене темы: {ex.Message}");
            }
        }
    }
}