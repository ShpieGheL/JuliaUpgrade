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
    /// Логика взаимодействия для Window10.xaml
    /// </summary>
    public partial class Window10 : Window
    {
        OleDbConnection connection;
        int n;
        public Window10()
        {
            InitializeComponent();
        }

        public Window10(string i1, string i2, string i3, string i4, string Back, string Text, string Set, OleDbConnection con) : this()
        {
            n = Convert.ToInt32(i1);
            connection = con;
            T1.Text = i2;
            C1.Text = i3;
            T2.Text = i4;
            Color c = (Color)ColorConverter.ConvertFromString(Back);
            SolidColorBrush co = new SolidColorBrush(c);
            G.Background = co;
            c = (Color)ColorConverter.ConvertFromString(Text);
            co = new SolidColorBrush(c);
            foreach (TextBox l in G.Children.OfType<TextBox>())
                l.Foreground = co;
            foreach (ComboBox l in G.Children.OfType<ComboBox>())
                l.Foreground = co;
            if (Set=="+")
            {
                foreach (TextBox l in G.Children.OfType<TextBox>())
                    l.Background = co;
                foreach (ComboBox l in G.Children.OfType<ComboBox>())
                    l.Background = co;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            double d;
            try
            {
                d=Convert.ToInt32(T2.Text);
            }
            catch
            {
                d = 0;
            }
            OleDbCommand co = new OleDbCommand("UPDATE Tax SET Описание='" + T1.Text + "', Тип='" + C1.Text + "', Плата=" + d + " WHERE Номер="+n, connection);
            co.ExecuteNonQuery();
            MessageBox.Show("Информация о налоге успешно изменена.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            this.Close();
        }
    }
}
