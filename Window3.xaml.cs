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
using System.IO;

namespace JuliaUpgrade
{
    /// <summary>
    /// Логика взаимодействия для Window3.xaml
    /// </summary>
    public partial class Window3 : Window
    {
        public string pathdb;
        public string path;
        public char c = ' ';
        public string Back = "";
        public string Text = "";
        public string Set = "";
        public Window3()
        {
            InitializeComponent();
            string strBack = "";
            string strText = "";
            string strSet = "";
            foreach (char s in Environment.SystemDirectory)
            {
                c = s;
                break;
            }
            pathdb = $@"{c}:\Users\{Environment.UserName}\Desktop\JL\JuliaDB.accdb";
            path = $"{c}:/Users/{Environment.UserName}/Desktop/JL";
            int p = 0;
            foreach (char c in File.ReadAllText($"{path}/Config.txt"))
            {
                if (c == ' ')
                    p++;
                if (c != ' ')
                    switch (p)
                    {
                        case 0:
                            strBack += c;
                            break;
                        case 1:
                            strText += c;
                            break;
                        case 2:
                            strSet += c;
                            break;
                    }
            }
            Back = strBack;
            Text = strText;
            Set = strSet;
            BackCol();
            TextCol();
        }

        private void BackCol()
        {
            Color c = (Color)ColorConverter.ConvertFromString(Back);
            SolidColorBrush co = new SolidColorBrush(c);
            if (Set == "+")
                foreach (TextBox l in G1.Children.OfType<TextBox>())
                    l.Background = co;
            G1.Background = co;
        }

        private void TextCol()
        {
            Color c = (Color)ColorConverter.ConvertFromString(Text);
            SolidColorBrush co = new SolidColorBrush(c);
            foreach (Label l in G1.Children.OfType<Label>())
                l.Foreground = co;
            foreach (TextBox l in G1.Children.OfType<TextBox>())
                l.Foreground = co;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OleDbConnection connection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={pathdb}");
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Work_types", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Work_types");
            if (Tab1T1.Text != "" || Tab1T2.Text != "")
            {
                OleDbCommand co = new OleDbCommand("INSERT INTO Work_types ([Вид работы],Цена) VALUES ('" + Tab1T1.Text + "','" + Convert.ToInt32(Tab1T2.Text) + "')", connection);
                co.ExecuteNonQuery();
                MessageBox.Show("Вид работы успешно добавлен", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
                MessageBox.Show("Не введены некоторые данные","Ошибка",MessageBoxButton.OK,MessageBoxImage.Error);
        }
    }
}