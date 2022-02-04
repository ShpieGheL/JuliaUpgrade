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
using System.Data.OleDb;
using System.Data;

namespace JuliaUpgrade
{
    /// <summary>
    /// Логика взаимодействия для Window8.xaml
    /// </summary>
    public partial class Window8 : Window
    {
        OleDbConnection connection;
        int n;
        string d1;
        public Window8()
        {
            InitializeComponent();
        }

        public Window8(string i, string d, string op, string mp, string mm,string Back, string Text, string Set,OleDbConnection con) : this()
        {
            d1 = d;
            n = Convert.ToInt32(i);
            connection = con;
            D1.Text = d;
            T1.Text = op;
            T2.Text = mp;
            T3.Text = mm;
            Color c = (Color)ColorConverter.ConvertFromString(Back);
            SolidColorBrush co = new SolidColorBrush(c);
            G.Background = co;
            c = (Color)ColorConverter.ConvertFromString(Text);
            co = new SolidColorBrush(c);
            foreach (TextBox l in G.Children.OfType<TextBox>())
                l.Foreground = co;
            foreach (ComboBox l in G.Children.OfType<ComboBox>())
                l.Foreground = co;
            foreach (DatePicker l in G.Children.OfType<DatePicker>())
                l.Foreground = co;
            if (Set == "+")
            {
                foreach (TextBox l in G.Children.OfType<TextBox>())
                    l.Background = co;
                foreach (ComboBox l in G.Children.OfType<ComboBox>())
                    l.Background = co;
                foreach (DatePicker l in G.Children.OfType<DatePicker>())
                    l.Background = co;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OleDbCommand co = new OleDbCommand("UPDATE Control SET Дата='" + D1.SelectedDate + "', Описание='" + T1.Text + "', Заработок=" + T2.Text + ", Траты=" + T3.Text + " WHERE Номер=" + n, connection);
            co.ExecuteNonQuery();
            MessageBox.Show("Запись успешно изменена.", "Успех!", MessageBoxButton.OK, MessageBoxImage.Information);
            this.Close();
        }
    }
}
