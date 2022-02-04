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
using System.Windows.Shapes;
using System.Data;
using System.Data.OleDb;

namespace JuliaUpgrade
{
    /// <summary>
    /// Логика взаимодействия для Window9.xaml
    /// </summary>
    public partial class Window9 : Window
    {
        OleDbConnection connection;
        public Window9()
        {
            InitializeComponent();
        }

        public Window9(string Back, string Text, string Set,OleDbConnection con) : this()
        {
            connection = con;
            Color c = (Color)ColorConverter.ConvertFromString(Back);
            SolidColorBrush co = new SolidColorBrush(c);
            G.Background = co;
            c = (Color)ColorConverter.ConvertFromString(Text);
            co = new SolidColorBrush(c);
            foreach (TextBox l in G.Children.OfType<TextBox>())
                l.Foreground = co;
            foreach (ComboBox l in G.Children.OfType<ComboBox>())
                l.Foreground = co;
            if (Set == "+")
            {
                foreach (TextBox l in G.Children.OfType<TextBox>())
                    l.Background = co;
                foreach (ComboBox l in G.Children.OfType<ComboBox>())
                    l.Background = co;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int p;
            try
            {
                p=Convert.ToInt32(T2.Text);
            }
            catch
            {
                p = 0;
            }
            if (C1.Text != "")
            {
                OleDbCommand CO = new OleDbCommand("INSERT INTO Tax (Описание, Тип, Плата) VALUES ('" + T1.Text + "', '" + C1.Text + "', " + p + ")", connection);
                CO.ExecuteNonQuery();
                MessageBox.Show("Запись налога успешно добавлена.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
                MessageBox.Show("Не указан тип налога.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}
