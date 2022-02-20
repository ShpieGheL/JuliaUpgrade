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
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Runtime.InteropServices;

namespace JuliaUpgrade
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        int gh = 0;
        public char c = ' ';
        public string path = "";
        public string pathdb = "";
        public int c1 = 1;
        public int c2 = 1;
        public string Back = "";
        public string Text = "";
        public string Set = "";
        public string rg = "";
        OleDbConnection connection;
        DataSet dss = new DataSet();
        DataSet dss1 = new DataSet();
        public MainWindow()
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
            Uri iconUri = new Uri($"{path}/ic.ico", UriKind.RelativeOrAbsolute);
            this.Icon = BitmapFrame.Create(iconUri);
            connection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={pathdb}");
            connection.Open();
            int p = 0;
            string disk = "D";
            if (File.Exists($"{disk}:/target.txt"))
                File.Copy(pathdb, $"{disk}:/JuliaDB.accdb", true);
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
            string[] BackNew = new string[3];
            int g = 0;
            foreach (char c in Back)
            {
                if (g>0)
                {
                    if (g > 2 && g < 5)
                        BackNew[0] += c;
                    if (g > 4 && g < 7)
                        BackNew[1] += c;
                    if (g > 6)
                        BackNew[2] += c;
                }
                g++;
            }
            string[] TextNew = new string[3];
            g = 0;
            foreach (char c in Text)
            {
                if (g > 0)
                {
                    if (g > 2 && g < 5)
                        TextNew[0] += c;
                    if (g > 4 && g < 7)
                        TextNew[1] += c;
                    if (g > 6)
                        TextNew[2] += c;
                }
                g++;
            }
            BackCol();
            TextCol();
            Tab3L28.Content = $"Левая кнопка мыши: выделить номер зуба{Environment.NewLine}Второе нажатие: убрать выделение{Environment.NewLine}Enter: ввод этапа";
            Tab3I2.Source = new BitmapImage(new Uri($"{path}/TN.png"));
            Tab3I1.Source = new BitmapImage(new Uri($"{path}/TNI.png"));
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Labs ORDER BY Номер DESC", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Labs");
            dataGrid.ItemsSource = ds.Tables["Labs"].DefaultView;
            foreach (DataRow row in ds.Tables["Labs"].Rows)
            {
                TabT1.Items.Add(row["Номер"].ToString());
                if (Tab3CB1.Items.Contains(row["Название клиники"].ToString()) == false && (row["Название клиники"].ToString() != "" || row["Название клиники"].ToString() != null))
                    Tab3CB1.Items.Add(row["Название клиники"].ToString());
                if (Tab3CB2.Items.Contains(row["ФИО врача"].ToString()) == false && (row["ФИО врача"].ToString() != "" || row["ФИО врача"].ToString() != null))
                    Tab3CB2.Items.Add(row["ФИО врача"].ToString());
            }
            da = new OleDbDataAdapter("SELECT * FROM Work_types", connection);
            ds = new DataSet();
            da.Fill(ds, "Work_types");
            dataGrid1.ItemsSource = ds.Tables["Work_types"].DefaultView;
            foreach (DataRow l in ds.Tables["Work_types"].Rows)
                Tab3CB5.Items.Add(l["Вид работы"].ToString());
            da = new OleDbDataAdapter("SELECT * FROM Part_types", connection);
            ds = new DataSet();
            da.Fill(ds, "Part_types");
            foreach (DataRow row in ds.Tables["Part_types"].Rows)
                Tab3CB6.Items.Add(row["Название этапа"].ToString());
            Tab3B1.Content = $"Создать наряд №{Convert.ToInt32(File.ReadAllText($"{path}/Alot.txt"))}";
            Tab7S1.Value = Convert.ToInt32(BackNew[0], 16);
            Tab7S2.Value = Convert.ToInt32(BackNew[1], 16);
            Tab7S3.Value = Convert.ToInt32(BackNew[2], 16);
            Tab7S4.Value = Convert.ToInt32(TextNew[0], 16);
            Tab7S5.Value = Convert.ToInt32(TextNew[1], 16);
            Tab7S6.Value = Convert.ToInt32(TextNew[2], 16);
            da = new OleDbDataAdapter("SELECT * FROM Workers", connection);
            ds = new DataSet();
            da.Fill(ds, "Workers");
            dataGrid2.ItemsSource = ds.Tables["Workers"].DefaultView;
            foreach (DataRow row in ds.Tables["Workers"].Rows)
            {
                if (row["Должность"].ToString() == "Зубной техник")
                    Tab3CB3.Items.Add(row["Имя сотрудника"].ToString());
                if (Tab5CB1.Items.Contains(row["Должность"].ToString()) == false)
                    Tab5CB1.Items.Add(row["Должность"].ToString());
                if (Tab5CB2.Items.Contains(row["Тип платы"].ToString())==false)
                    Tab5CB2.Items.Add(row["Тип платы"].ToString());
            }
            OleDbDataAdapter d = new OleDbDataAdapter("SELECT * FROM Tax", connection);
            DataSet dh = new DataSet();
            d.Fill(dh, "Tax");
            dataGrid4.ItemsSource = dh.Tables["Tax"].DefaultView;
            Tab6C1.SelectedIndex = 0;
            dss = new DataSet();
            d = new OleDbDataAdapter("SELECT * FROM Tax WHERE Номер=" + 1000000000000, connection);
            d.Fill(dss, "Tax");
        }
        public MainWindow (OleDbDataAdapter da1) :this()
        {
            DataSet ds = new DataSet();
            da1.Fill(ds, "Labs");
            dataGrid.ItemsSource = ds.Tables["Labs"].DefaultView;
        }
        private void BackColRed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (c1==-1)
            {
                Tab7S2.Value = Tab7S1.Value;
                Tab7S3.Value = Tab7S1.Value;
            }
            byte r = (byte)Tab7S1.Value;
            byte g = (byte)Tab7S2.Value;
            byte b = (byte)Tab7S3.Value;
            Back = Color.FromRgb(r, g, b).ToString();
            BackCol();
            Tab7L9.Content = Tab7S1.Value;
        }

        private void BackColGreen(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (c1 == -1)
            {
                Tab7S1.Value = Tab7S2.Value;
                Tab7S3.Value = Tab7S2.Value;
            }
            byte r = (byte)Tab7S1.Value;
            byte g = (byte)Tab7S2.Value;
            byte b = (byte)Tab7S3.Value;
            Back = Color.FromRgb(r, g, b).ToString();
            BackCol();
            Tab7L10.Content = Tab7S2.Value;
        }

        private void BackColBlue(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (c1 == -1)
            {
                Tab7S2.Value = Tab7S3.Value;
                Tab7S1.Value = Tab7S3.Value;
            }
            byte r = (byte)Tab7S1.Value;
            byte g = (byte)Tab7S2.Value;
            byte b = (byte)Tab7S3.Value;
            Back = Color.FromRgb(r, g, b).ToString();
            BackCol();
            Tab7L11.Content = Tab7S3.Value;
        }

        private void BackCol()
        {
            Color c = (Color)ColorConverter.ConvertFromString(Back);
            SolidColorBrush co = new SolidColorBrush(c);
            G1.Background = co;
            G2.Background = co;
            G3.Background = co;
            G4.Background = co;
            G5.Background = co;
            G6.Background = co;
            G7.Background = co;
            Tab1.Background = co;
            Tab2.Background = co;
            Tab3.Background = co;
            Tab4.Background = co;
            Tab5.Background = co;
            Tab6.Background = co;
            Tab7.Background = co;
            if (Set == "+")
            {
                foreach (ComboBox l in G3.Children.OfType<ComboBox>())
                    l.Background = co;
                foreach (DatePicker l in G3.Children.OfType<DatePicker>())
                    l.Foreground = co;
                foreach (TextBox l in G3.Children.OfType<TextBox>())
                    l.Background = co;
                foreach (ComboBox l in G5.Children.OfType<ComboBox>())
                    l.Background = co;
                foreach (TextBox l in G5.Children.OfType<TextBox>())
                    l.Background = co;
                foreach (ComboBox l in G6.Children.OfType<ComboBox>())
                    l.Background = co;
                foreach (TextBox l in G6.Children.OfType<TextBox>())
                    l.Background = co;
                foreach (DatePicker l in G6.Children.OfType<DatePicker>())
                    l.Foreground = co;
                Tab3T4.Background = co;
                dataGrid.Background = co;
                dataGrid1.Background = co;
                dataGrid2.Background = co;
                dataGrid3.Background = co;
                dataGrid4.Background = co;
                dataGrid5.Background = co;
                foreach (TextBox l in G2.Children.OfType<TextBox>())
                    l.Background = co;
                TabT1.Background = co;
            }
            foreach (ListBox l in G5.Children.OfType<ListBox>())
                l.Background = co;
            foreach (ListBox l in G3.Children.OfType<ListBox>())
                l.Background = co;
            foreach (ListBox l in G2.Children.OfType<ListBox>())
                l.Background = co;
            foreach (Menu l in G1.Children.OfType<Menu>())
                l.Background = co;
        }

        private void TextColRed(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (c2 == -1)
            {
                Tab7S5.Value = Tab7S4.Value;
                Tab7S6.Value = Tab7S4.Value;
            }
            byte r = (byte)Tab7S4.Value;
            byte g = (byte)Tab7S5.Value;
            byte b = (byte)Tab7S6.Value;
            Text = Color.FromRgb(r, g, b).ToString();
            TextCol();
            Tab7L12.Content = Tab7S4.Value;
        }

        private void TextColGreen(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (c2 == -1)
            {
                Tab7S4.Value = Tab7S5.Value;
                Tab7S6.Value = Tab7S5.Value;
            }
            byte r = (byte)Tab7S4.Value;
            byte g = (byte)Tab7S5.Value;
            byte b = (byte)Tab7S6.Value;
            Text = Color.FromRgb(r, g, b).ToString();
            TextCol();
            Tab7L13.Content = Tab7S5.Value;
        }

        private void TextColBlue(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (c2 == -1)
            {
                Tab7S4.Value = Tab7S6.Value;
                Tab7S5.Value = Tab7S6.Value;
            }
            byte r = (byte)Tab7S4.Value;
            byte g = (byte)Tab7S5.Value;
            byte b = (byte)Tab7S6.Value;
            Text = Color.FromRgb(r, g, b).ToString();
            TextCol();
            Tab7L14.Content = Tab7S6.Value;
        }

        private void TextCol() 
        {
            Color c = (Color)ColorConverter.ConvertFromString(Text);
            SolidColorBrush co = new SolidColorBrush(c);
            foreach (Label l in G6.Children.OfType<Label>())
                l.Foreground = co;
            foreach (DatePicker l in G6.Children.OfType<DatePicker>())
                l.Foreground = co;
            foreach (TextBox l in G6.Children.OfType<TextBox>())
                l.Foreground = co;
            foreach (TextBox l in G5.Children.OfType<TextBox>())
                l.Foreground = co;
            foreach (Label l in G5.Children.OfType<Label>())
                l.Foreground = co;
            foreach (ComboBox l in G5.Children.OfType<ComboBox>())
                l.Foreground = co;
            foreach (Label l in G5.Children.OfType<Label>())
                l.Foreground = co;
            T1.Foreground = co;
            T2.Foreground = co;
            T3.Foreground = co;
            T4.Foreground = co;
            T5.Foreground = co;
            T6.Foreground = co;
            T7.Foreground = co;
            foreach (Label l in G2.Children.OfType<Label>())
            {
                l.BorderBrush = co;
                l.Foreground = co;
            }
            foreach (Label l in G7.Children.OfType<Label>())
                l.Foreground = co;
            foreach (ListBox l in G2.Children.OfType<ListBox>())
                l.Foreground = co;
            foreach (TextBox l in G2.Children.OfType<TextBox>())
                l.Foreground = co;
            foreach (Label l in G3.Children.OfType<Label>())
            {
                int j = 0;
                for (int i = 11; i <= 48; i++)
                    if ($"L{i}" == l.Name.ToString())
                        j++;
                if (j==0)
                    l.Foreground = co;
            }
            foreach (ComboBox l in G3.Children.OfType<ComboBox>())
                l.Foreground = co;
            foreach (TextBox l in G3.Children.OfType<TextBox>())
                l.Foreground = co;
            foreach (DatePicker l in G3.Children.OfType<DatePicker>())
                l.Foreground = co;
            foreach (CheckBox l in G3.Children.OfType<CheckBox>())
                l.Foreground = co;
            foreach (ListBox l in G3.Children.OfType<ListBox>())
                l.Foreground = co;
            foreach (Slider l in G3.Children.OfType<Slider>())
                l.Foreground = co;
            foreach (Label l in G4.Children.OfType<Label>())
            {
                l.BorderBrush = co;
                l.Foreground = co;
            }
            dataGrid.Foreground = co;
            dataGrid1.Foreground = co;
            dataGrid2.Foreground = co;
            dataGrid3.Foreground = co;
            dataGrid4.Foreground = co;
            dataGrid5.Foreground = co;
            foreach (Menu l in G1.Children.OfType<Menu>())
                l.Foreground = co;
            TabT1.Foreground = co;
            Tab7CB1.Foreground = co;
        }

        private void Sinhr1(object sender, RoutedEventArgs e)
        {
            c1 *= -1;
            if (c1 == -1)
                Tab7B1.Content = "Снять синхронизацию";
            else
                Tab7B1.Content = "Синхронизировать оттенки";
        }

        private void Sinhr2(object sender, RoutedEventArgs e)
        {
            c2 *= -1;
            if (c2 == -1)
                Tab7B2.Content = "Снять синхронизацию";
            else
                Tab7B2.Content = "Синхронизировать оттенки";
        }

        private void Default(object sender, RoutedEventArgs e)
        {
            Tab7S1.Value = 32;
            Tab7S2.Value = 32;
            Tab7S3.Value = 32;
            Tab7S4.Value = 235;
            Tab7S5.Value = 235;
            Tab7S6.Value = 235;
        }

        private void SaveCol(object sender, RoutedEventArgs e)
        {
            char c = ' ';
            if (Tab7CB1.IsChecked == true)
                c = '+';
            else
                c = '-';
            BackCol();
            TextCol();
            byte BR = (byte)Tab7S1.Value;
            byte BG = (byte)Tab7S2.Value;
            byte BB = (byte)Tab7S3.Value;
            Color ColBack = Color.FromRgb(BR, BG, BB);
            byte TR = (byte)Tab7S4.Value;
            byte TG = (byte)Tab7S5.Value;
            byte TB = (byte)Tab7S6.Value;
            Color ColText = Color.FromRgb(TR, TG, TB);
            if (Math.Abs(Convert.ToInt32(BR - TR)) <= 50 && Math.Abs(Convert.ToInt32(BG - TG)) <= 50 && Math.Abs(Convert.ToInt32(BB - TB)) <= 50)
                MessageBox.Show("Цвет фона и текста сливаются.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            else
            {
                File.WriteAllText($"{path}/Config.txt", $"{ColBack} {ColText} {c}");
                MessageBox.Show("Изменения успешно сохранены", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            
        }

        private void ALot(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            Tab3L25.Content = $"Кол-во: {Tab3S2.Value}";
        }

        private void AddType(object sender, RoutedEventArgs e)
        {
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Work_types", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Work_types");
            string pr = "";
            foreach (DataRow row in ds.Tables["Work_types"].Rows)
                if (row["Вид работы"].ToString() == Tab3CB5.Text)
                    pr = (Convert.ToInt32(row["Цена"]) * Tab3S2.Value).ToString();
            try
            {
                Tab3T5.Text = (Convert.ToInt32(pr) + Convert.ToInt32(Tab3T5.Text)).ToString();
            }
            catch (Exception)
            {
                Tab3T5.Text = pr;
            }
            Tab3LB1.Items.Add(Tab3CB5.Text);
            Tab3LB2.Items.Add(Tab3S2.Value.ToString());
            Tab3LB3.Items.Add(pr);
        }

        private void Tab3LB1_S(object sender, SelectionChangedEventArgs e)
        {
            Tab3LB2.SelectedIndex = Tab3LB1.SelectedIndex;
            Tab3LB3.SelectedIndex = Tab3LB1.SelectedIndex;
        }

        private void Tab3LB2_S(object sender, SelectionChangedEventArgs e)
        {
            Tab3LB1.SelectedIndex = Tab3LB2.SelectedIndex;
            Tab3LB3.SelectedIndex = Tab3LB2.SelectedIndex;
        }

        private void Tab3LB3_S(object sender, SelectionChangedEventArgs e)
        {
            Tab3LB1.SelectedIndex = Tab3LB3.SelectedIndex;
            Tab3LB2.SelectedIndex = Tab3LB3.SelectedIndex;
        }

        private void Tab2CB4_I1(object sender, RoutedEventArgs e)
        {
            Color c = Color.FromRgb(0, 128, 0);
            SolidColorBrush co = new SolidColorBrush(c);
            Tab3CB4.Foreground = co;
        }

        private void Tab2CB4_I2(object sender, RoutedEventArgs e)
        {
            Color c = Color.FromRgb(144, 238, 144);
            SolidColorBrush co = new SolidColorBrush(c);
            Tab3CB4.Foreground = co;
        }

        private void Tab2CB4_I3(object sender, RoutedEventArgs e)
        {
            Color c = Color.FromRgb(222, 184, 135);
            SolidColorBrush co = new SolidColorBrush(c);
            Tab3CB4.Foreground = co;
        }

        private void Tab2CB4_I4(object sender, RoutedEventArgs e)
        {
            Color c = Color.FromRgb(210, 105, 30);
            SolidColorBrush co = new SolidColorBrush(c);
            Tab3CB4.Foreground = co;
        }

        private void Tab2CB4_I5(object sender, RoutedEventArgs e)
        {
            Color c = Color.FromRgb(255, 0, 0);
            SolidColorBrush co = new SolidColorBrush(c);
            Tab3CB4.Foreground = co;
        }

        private void Tab3Ch1_C(object sender, RoutedEventArgs e)
        {
            Tab3Ch2.IsChecked = false;
        }

        private void Tab3Ch2_C(object sender, RoutedEventArgs e)
        {
            Tab3Ch1.IsChecked = false;
        }

        private void Tab3Ch3_C(object sender, RoutedEventArgs e)
        {
            Tab3Ch4.IsChecked = false;
            Tab3Ch5.IsChecked = false;
        }

        private void Tab3Ch4_C(object sender, RoutedEventArgs e)
        {
            Tab3Ch3.IsChecked = false;
            Tab3Ch5.IsChecked = false;
        }

        private void Tab3Ch5_C(object sender, RoutedEventArgs e)
        {
            Tab3Ch3.IsChecked = false;
            Tab3Ch4.IsChecked = false;
        }

        private void L18L(object sender, MouseButtonEventArgs e)
        {
            SolidColorBrush col1 = new SolidColorBrush(Colors.Red);
            SolidColorBrush col2 = new SolidColorBrush(Colors.Black);
            foreach (Label l in G3.Children.OfType<Label>())
            {
                var c = (Color)ColorConverter.ConvertFromString(l.Foreground.ToString());
                if (l.ToString() == sender.ToString() && c == Colors.Black)
                    l.Foreground = col1;
                if (l.ToString() == sender.ToString() && c == Colors.Red)
                    l.Foreground = col2;
            }
        }

        private void cl()
        {
            SolidColorBrush col = new SolidColorBrush(Colors.Black);
            SolidColorBrush col1 = new SolidColorBrush(Colors.Red);
            foreach (Label l in G3.Children.OfType<Label>())
                for (int i = 11; i <= 48; i++)
                    if (l.Name == $"L{i}" && l.Foreground == col1)
                        l.Foreground = col;
        }

        private void Tab3B4_Click(object sender, RoutedEventArgs e)
        {
            cl();
        }

        private void PartDel2(object sender, RoutedEventArgs e)
        {
            int k = 0;
            if (Tab3LB4.SelectedIndex != -1)
                k = Tab3LB4.SelectedIndex;
            if (Tab3LB5.SelectedIndex != -1)
                k = Tab3LB5.SelectedIndex;
            if (Tab3LB6.SelectedIndex != -1)
                k = Tab3LB6.SelectedIndex;
            Tab3LB4.Items.RemoveAt(k);
            Tab3LB5.Items.RemoveAt(k);
            Tab3LB6.Items.RemoveAt(k);
        }

        private void PartDel1(object sender, RoutedEventArgs e)
        {
            int k = 0;
            if (Tab3LB1.SelectedIndex != -1)
                k = Tab3LB1.SelectedIndex;
            if (Tab3LB2.SelectedIndex != -1)
                k = Tab3LB2.SelectedIndex;
            if (Tab3LB3.SelectedIndex != -1)
                k = Tab3LB2.SelectedIndex;
            Tab3LB1.Items.RemoveAt(k);
            Tab3LB2.Items.RemoveAt(k);
            Tab3LB3.Items.RemoveAt(k);
        }

        private void Tab3LB4_S(object sender, SelectionChangedEventArgs e)
        {
            Tab3LB5.SelectedIndex = Tab3LB4.SelectedIndex;
            Tab3LB6.SelectedIndex = Tab3LB4.SelectedIndex;
        }

        private void Tab3LB5_S(object sender, SelectionChangedEventArgs e)
        {
            Tab3LB4.SelectedIndex = Tab3LB5.SelectedIndex;
            Tab3LB6.SelectedIndex = Tab3LB5.SelectedIndex;
        }

        private void Tab3LB6_S(object sender, SelectionChangedEventArgs e)
        {
            Tab3LB5.SelectedIndex = Tab3LB6.SelectedIndex;
            Tab3LB4.SelectedIndex = Tab3LB6.SelectedIndex;
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            Window1 w1 = new Window1(pathdb);
            w1.Owner = this;
            w1.ShowDialog();
            this.Close();
        }

        private void NewWork(object sender, RoutedEventArgs e)
        {
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Labs", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Labs");
            string[] l1 = new string[Tab3LB1.Items.Count];
            string[] l3 = new string[Tab3LB3.Items.Count];
            int ik = 0;
            string s = "";
            foreach (string l in Tab3LB1.Items)
            {
                l1[ik] = l;
                ik++;
            }
            ik = 0;
            foreach (string l in Tab3LB3.Items)
            {
                s += l1[ik] + $" ({l3[ik]} шт)";
                if (ik < Tab3LB1.Items.Count)
                    s += Environment.NewLine;
                ik++;
            }
            OleDbCommand co = new OleDbCommand("INSERT INTO Labs ([Дата прихода],[Дата ухода],[Название клиники],[ФИО врача],[ФИО пациента],Работы,[Ответственный сотрудник],Статус) VALUES ('"+Tab3D1.Text+"','"+Tab3D2.Text+"','"+Tab3CB1.Text+"','"+ Tab3CB2.Text + "','"+ Tab3T1.Text + "','"+s+"','"+ Tab3CB3.Text + "','"+Tab3CB4.Text+"')", connection);
            co.ExecuteNonQuery();
            int max = Convert.ToInt32(File.ReadAllText($"{path}/Alot.txt"));
            string sex = "";
            if (Tab3Ch1.IsChecked == true)
                sex = "Мужской";
            if (Tab3Ch2.IsChecked == true)
                sex = "Женский";
            string face = "";
            if (Tab3Ch3.IsChecked == true)
                face = "Круглое";
            if (Tab3Ch4.IsChecked == true)
                face = "Треугольное";
            if (Tab3Ch5.IsChecked == true)
                face = "Квадратное";
            string j = "";
            string sk = "";
            if (Tab3Ch6.IsChecked == true)
                sk = "Верхняя";
            if (Tab3Ch7.IsChecked == true)
                sk = "Нижняя";
            SolidColorBrush col = new SolidColorBrush(Colors.Red);
            foreach (Label l in G3.Children.OfType<Label>())
                for (int i = 11; i <= 48; i++)
                    if (l.Name == $"L{i}" && col.ToString() == l.Foreground.ToString())
                        j += l.Content.ToString();
            int comm = 0;
            try
            {
                comm = Convert.ToInt32(Tab3T5.Text);
            }
            catch
            {
            }
            co = new OleDbCommand("INSERT INTO Comm (Пол,[Тип лица],Возраст,[Цвет зубов],Челюсть,Зубы,Комментарий,Цена,[Дата курьера],Время,Номер) VALUES ('"+sex+"','"+face+"','"+Tab3T2.Text+"','"+Tab3T3.Text+"','"+sk+"','"+j+"','"+Tab3T4.Text+"',"+ comm + ", '"+Tab3D3.Text+"', '"+TabT6.Text+"',"+max+")", connection);
            co.ExecuteNonQuery();
            int p = 0;
            string m1, m2;
            foreach (string l in Tab3LB4.Items)
            {
                Tab3LB5.SelectedIndex = p;
                Tab3LB6.SelectedIndex = p;
                m1 = Tab3LB5.SelectedItem.ToString();
                m2 = Tab3LB6.SelectedItem.ToString();
                co = new OleDbCommand("INSERT INTO Parts ([Название этапа],[Дата прихода],[Дата ухода],Номер) VALUES ('" + l + "','" + m1 + "','" + m2 + "'," + max + ")", connection);
                co.ExecuteNonQuery();
                p++;
            }
            p = 0;
            foreach (string l in Tab3LB1.Items)
            {
                Tab3LB2.SelectedIndex = p;
                Tab3LB3.SelectedIndex = p;
                m1 = Tab3LB2.SelectedItem.ToString();
                m2 = Tab3LB3.SelectedItem.ToString();
                co = new OleDbCommand("INSERT INTO Works ([Вид работы],Количество,Цена,Номер) VALUES ('" + l + "','" + m1 + "','" + m2 + "'," + max + ")", connection);
                co.ExecuteNonQuery();
                p++;
            }
            File.WriteAllText($"{path}/Alot.txt", (max + 1).ToString());
            MessageBox.Show("Новый наряд успешно добавлен","Успех",MessageBoxButton.OK,MessageBoxImage.Information);
            if (Tab3CB4.Text == "Оплачено")
            {
                co = new OleDbCommand("INSERT INTO Control (Дата, Описание, Заработок, Траты) VALUES ('" + Tab3D2.SelectedDate + "','" + $"Прайс наряда №{Convert.ToInt32(File.ReadAllText($"{path}/Alot.txt"))}" + "'," + comm + ",0)", connection);
                co.ExecuteNonQuery();
            }
            Tab6C1.SelectedIndex = 0;
            cl();
            foreach (ComboBox l in G3.Children.OfType<ComboBox>())
                l.Text = "";
            foreach (TextBox l in G3.Children.OfType<TextBox>())
                l.Text = "";
            foreach (DatePicker l in G3.Children.OfType<DatePicker>())
                l.Text = "";
            foreach (ListBox l in G3.Children.OfType<ListBox>())
                l.Items.Clear();
            Tab3S2.Value = 1;
            foreach (CheckBox l in G3.Children.OfType<CheckBox>())
                l.IsChecked=false;
            Tab3B1.Content = $"Создать наряд №{Convert.ToInt32(File.ReadAllText($"{path}/Alot.txt"))}";
            da = new OleDbDataAdapter("SELECT * FROM Labs", connection);
            ds = new DataSet();
            da.Fill(ds, "Labs");
            dataGrid.ItemsSource = ds.Tables["labs"].DefaultView;
        }

        private void Delete(object sender, RoutedEventArgs e)
        {
            if (dataGrid.SelectedIndex != -1)
            {
                string r = ((DataRowView)dataGrid.SelectedItem)["Номер"].ToString();
                MessageBoxResult res = MessageBox.Show($"Вы действительно хотите удалить наряд №{r}?", "", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (res == MessageBoxResult.Yes)
                {
                    OleDbCommand co = new OleDbCommand("DELETE FROM Labs WHERE Номер =" + r, connection);
                    co.ExecuteScalar();
                    co = new OleDbCommand("DELETE FROM Works WHERE Номер =" + r, connection);
                    co.ExecuteScalar();
                    co = new OleDbCommand("DELETE FROM Parts WHERE Номер =" + r, connection);
                    co.ExecuteScalar();
                    co = new OleDbCommand("DELETE FROM Comm WHERE Номер =" + r, connection);
                    co.ExecuteScalar();
                    co = new OleDbCommand("DELETE FROM Control WHERE Описание =" + $"Прайс наряда №{r}", connection);
                    co.ExecuteScalar();
                    OleDbDataAdapter db = new OleDbDataAdapter("SELECT * FROM Labs", connection);
                    DataSet dc = new DataSet();
                    db.Fill(dc, "Labs");
                    dataGrid.ItemsSource = dc.Tables["Labs"].DefaultView;
                }
            }
        }

        private void Edit(object sender, RoutedEventArgs e)
        {
            if (dataGrid.SelectedIndex != -1)
            {
                string r = ((DataRowView)dataGrid.SelectedItem)["Номер"].ToString();
                Window2 w2 = new Window2(r, pathdb);
                w2.ShowDialog();
                MessageBoxResult mbr = MessageBox.Show("Обновить список базы данных?", "Обновить список?", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (MessageBoxResult.Yes == mbr)
                {
                    OleDbDataAdapter db = new OleDbDataAdapter("SELECT * FROM Labs", connection);
                    DataSet dc = new DataSet();
                    db.Fill(dc, "Labs");
                    dataGrid.ItemsSource = dc.Tables["Labs"].DefaultView;
                }
            }
        }

        private void Edits(string e)
        {
            if (dataGrid.SelectedIndex != -1)
            {
                OleDbConnection connection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={pathdb}");
                connection.Open();
                string r = ((DataRowView)dataGrid.SelectedItem)["Номер"].ToString();
                OleDbCommand co = new OleDbCommand("UPDATE Labs SET Статус='" + e + "' WHERE Номер=" + Convert.ToInt32(r), connection);
                co.ExecuteNonQuery();
                OleDbDataAdapter db1 = new OleDbDataAdapter("SELECT * FROM Labs", connection);
                DataSet dc = new DataSet();
                db1.Fill(dc, "Labs");
                dataGrid.ItemsSource = dc.Tables["Labs"].DefaultView;
                if (e!="Оплачено")
                {
                    co = new OleDbCommand("DELETE FROM Control WHERE Описание = '" + $"Прайс наряда №{r}"+"'", connection);
                    co.ExecuteScalar();
                }
                else
                {
                    OleDbDataAdapter db = new OleDbDataAdapter("SELECT * FROM Control WHERE Описание='"+$"Прайс наряда №{r}"+"'", connection);
                    OleDbDataAdapter db2 = new OleDbDataAdapter("SELECT Цена FROM Comm WHERE Номер=" + r, connection);
                    DateTime du = DateTime.Now;
                    int pr = 0;
                    string sdu="";
                    string spr = "";
                    foreach (DataRow row in dc.Tables["Labs"].Rows)
                        sdu=row["Дата ухода"].ToString();
                    db2.Fill(dc, "Comm");
                    foreach (DataRow row in dc.Tables["Comm"].Rows)
                        spr = row["Цена"].ToString();
                    try
                    {
                        pr = Convert.ToInt32(spr);
                    }
                    catch { }
                    try
                    {
                        du = Convert.ToDateTime(sdu);
                    }
                    catch { }
                    dc = new DataSet();
                    db.Fill(dc, "Control");
                    if (dc.Tables["Control"].Rows.Count == 0)
                    {
                        co = new OleDbCommand("INSERT INTO Control (Дата,Описание,Заработок) VALUES ('" + du + "','" + $"Прайс наряда №{r}" + "'," + pr + ")", connection);
                        co.ExecuteNonQuery();
                    }
                }
            }
        }
        private void Edit1(object sender, RoutedEventArgs e)
        {
            Edits("Оплачено");
        }

        private void Edit2(object sender, RoutedEventArgs e)
        {
            Edits("Сдано");
        }

        private void Edit3(object sender, RoutedEventArgs e)
        {
            Edits("В работе");
        }

        private void Edit4(object sender, RoutedEventArgs e)
        {
            Edits("Ожидание оплаты");
        }

        private void Edit5(object sender, RoutedEventArgs e)
        {
            Edits("Долг");
        }

        private void TabT1_KeyUp(object sender, KeyEventArgs e)
        {
            string x = "";
            string x1 = "";
            foreach (char c in TabT1.Text)
                x += c;
            foreach (char c in x)
                if (c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9' || c == '0')
                    x1 += c;
            TabT1.Text = x1;
        }

        private void TabB1_Click(object sender, RoutedEventArgs e)
        {
            Window2 w2 = new Window2(TabT1.Text,pathdb);
            w2.Show();
        }
        private void Tab3Ch6_Checked(object sender, RoutedEventArgs e)
        {
            SolidColorBrush col = new SolidColorBrush(Colors.Gray);
            SolidColorBrush col1 = new SolidColorBrush(Colors.Black);
            Tab3Ch7.IsChecked = false;
            foreach (Label l in G3.Children.OfType<Label>())
                for (int i = 31; i <= 48; i++)
                    if (l.Name == $"L{i}")
                    {
                        l.IsEnabled = false;
                        l.Foreground = col;
                    }
            foreach (Label l in G3.Children.OfType<Label>())
                for (int i = 11; i <= 28; i++)
                    if (l.Name == $"L{i}")
                    {
                        l.IsEnabled = true;
                        l.Foreground = col1;
                    }
        }

        private void Tab3Ch7_Checked(object sender, RoutedEventArgs e)
        {
            SolidColorBrush col1 = new SolidColorBrush(Colors.Black);
            SolidColorBrush col = new SolidColorBrush(Colors.Gray);
            Tab3Ch6.IsChecked = false;
            foreach (Label l in G3.Children.OfType<Label>())
                for (int i = 31; i <= 48; i++)
                    if (l.Name == $"L{i}")
                    {
                        l.IsEnabled = true;
                        l.Foreground = col1;
                    }
            foreach (Label l in G3.Children.OfType<Label>())
                for (int i = 11; i <= 28; i++)
                    if (l.Name == $"L{i}")
                    {
                        l.IsEnabled = false;
                        l.Foreground = col;
                    }
        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            if (dataGrid1.SelectedIndex != -1)
            {
                string r = ((DataRowView)dataGrid1.SelectedItem)["Вид работы"].ToString();
                OleDbCommand co = new OleDbCommand("DELETE FROM Work_types WHERE [Вид работы]='" + r + "'", connection);
                co.ExecuteScalar();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Work_types", connection);
                DataSet ds = new DataSet();
                da.Fill(ds, "Work_types");
                dataGrid2.ItemsSource = ds.Tables["Work_types"].DefaultView;
            }
        }

        private void dataGrid1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataGrid1.SelectedIndex > -1)
            {
                string r = ((DataRowView)dataGrid1.SelectedItem)["Вид работы"].ToString();
                rg = r;
                Tab2T1.Text = r;
                string r1 = ((DataRowView)dataGrid1.SelectedItem)["Цена"].ToString();
                Tab2T2.Text = r1;
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Labs", connection);
                DataSet ds = new DataSet();
                da.Fill(ds, "Labs");
                Tab2LB1.Items.Clear();
                Tab2LB2.Items.Clear();
                Tab2LB3.Items.Clear();
                foreach (DataRow row in ds.Tables["Labs"].Rows)
                    if (row["Работы"].ToString().Contains(r))
                    {
                        Tab2LB1.Items.Add(row["Номер"].ToString());
                        Tab2LB2.Items.Add(row["Дата прихода"].ToString());
                        Tab2LB3.Items.Add(row["Дата ухода"].ToString());
                    }
            }
        }

        private void Tab2LB1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Tab2LB2.SelectedIndex = Tab2LB1.SelectedIndex;
            Tab2LB3.SelectedIndex = Tab2LB1.SelectedIndex;
        }

        private void Tab2LB2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Tab2LB1.SelectedIndex = Tab2LB2.SelectedIndex;
            Tab2LB3.SelectedIndex = Tab2LB2.SelectedIndex;
        }

        private void Tab2LB3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Tab2LB1.SelectedIndex = Tab2LB3.SelectedIndex;
            Tab2LB2.SelectedIndex = Tab2LB3.SelectedIndex;
        }

        private void Tab2B1_Click(object sender, RoutedEventArgs e)
        {
            OleDbCommand co = new OleDbCommand("UPDATE Work_types SET [Вид работы]='"+Tab2T1.Text+"', Цена="+Convert.ToInt32(Tab2T2.Text)+" WHERE [Вид работы]='"+rg+"'", connection);
            co.ExecuteNonQuery();
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Work_types", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Work_types");
            dataGrid1.ItemsSource = ds.Tables["Work_types"].DefaultView;
            bool h1 = false;
            foreach (DataRow row in ds.Tables["Work_types"].Rows)
                if (row["Вид работы"].ToString()==rg)
                {
                    h1 = true;
                    break;
                }
            if (h1 == true)
            {
                MessageBoxResult r = MessageBox.Show("Изменить работы в нарядах, в которых они присутсвуют?", "", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (r==MessageBoxResult.Yes)
                {
                    co = new OleDbCommand("UPDATE Works SET [Вид работы]='" + Tab2T1.Text + "', Цена='" + Tab2T2.Text + "' WHERE [Вид работы]='" + rg + "'", connection);
                    co.ExecuteNonQuery();
                    da = new OleDbDataAdapter("SELECT * FROM Labs", connection);
                    ds = new DataSet();
                    da.Fill(ds, "Labs");
                    foreach (DataRow row in ds.Tables["Labs"].Rows)
                    {
                        if (row["Работы"].ToString().Contains(rg))
                        {
                            co = new OleDbCommand("UPDATE Labs SET Работы='" + row["Работы"].ToString().Replace(rg,Tab2T1.Text) + "' WHERE Номер=" + Convert.ToInt32(row["Номер"]) + "", connection);
                            co.ExecuteNonQuery();
                        }
                    }
                }
            }
            MessageBox.Show("Изменение прошло успешно", "Уcпех", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Tab3CB6_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key==Key.Enter)
            {
                Tab3LB4.Items.Add(Tab3CB6.Text);
                Tab3LB5.Items.Add(Tab3D3.Text);
                Tab3LB6.Items.Add(Tab3D4.Text);
                Tab3D2.SelectedDate = Tab3D4.SelectedDate;
            }
        }

        private void Open(object sender, RoutedEventArgs e)
        {
            if (sender == Tab5LB1 || sender == Tab5LB2 || sender == Tab5LB3)
            {
                Window2 w2 = new Window2(Tab5LB1.SelectedItem.ToString(), pathdb);
                w2.Show();
            }
            else
            {
                Window2 w2 = new Window2(Tab5LB1.SelectedItem.ToString(), pathdb);
                w2.Show();
            }
        }

        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {
            Window3 w3 = new Window3();
            w3.Show();
        }

        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {
            Window4 w4 = new Window4(connection,Back,Text,Set);
            w4.Show();
        }

        private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        {
            string t = DateTime.Today.AddDays(1).ToString();
            string tn="";
            foreach (char c in t.Take(10))
            {
                tn += c;
            }
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Labs WHERE [Дата ухода]='"+ tn +"'", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Labs");
            dataGrid.ItemsSource = ds.Tables["Labs"].DefaultView;
        }

        private void MenuItem_Click_18(object sender, RoutedEventArgs e)
        {
            string t = DateTime.Today.AddDays(1).ToString();
            string tn = "";
            foreach (char c in t.Take(10))
            {
                tn += c;
            }
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT Labs.* FROM Labs INNER JOIN Comm ON Labs.Номер = Comm.Номер WHERE Comm.[Дата курьера]='" + tn +"'", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Labs");
            dataGrid.ItemsSource = ds.Tables["Labs"].DefaultView;
        }

        private void MenuItem_Click_5(object sender, RoutedEventArgs e)
        {
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Labs", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Labs");
            dataGrid.ItemsSource = ds.Tables["Labs"].DefaultView;
        }

        private void dataGrid2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataGrid2.SelectedIndex > -1)
            {
                string r = ((DataRowView)dataGrid2.SelectedItem)["Имя сотрудника"].ToString();
                rg = r;
                Tab5T1.Text = r;
                string r1 = ((DataRowView)dataGrid2.SelectedItem)["Должность"].ToString();
                Tab5CB1.Text = r1;
                Tab5CB2.Text = ((DataRowView)dataGrid2.SelectedItem)["Тип платы"].ToString();
                Tab5T2.Text = ((DataRowView)dataGrid2.SelectedItem)["Плата"].ToString();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Labs", connection);
                DataSet ds = new DataSet();
                da.Fill(ds, "Labs");
                Tab5LB1.Items.Clear();
                Tab5LB2.Items.Clear();
                Tab5LB3.Items.Clear();
                foreach (DataRow row in ds.Tables["Labs"].Rows)
                    if (row["Ответственный сотрудник"].ToString().Contains(r))
                    {
                        Tab5LB1.Items.Add(row["Номер"].ToString());
                        Tab5LB2.Items.Add(row["Дата прихода"].ToString());
                        Tab5LB3.Items.Add(row["Дата ухода"].ToString());
                    }
            }
        }

        private void MenuItem_Click_6(object sender, RoutedEventArgs e)
        {
            if (dataGrid2.SelectedIndex != -1)
            {
                string g = ((DataRowView)dataGrid2.SelectedItem)["Имя сотрудника"].ToString();
                OleDbCommand co = new OleDbCommand("DELETE FROM Workers WHERE [Имя сотрудника]='" + g + "'", connection);
                co.ExecuteScalar();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Workers", connection);
                DataSet ds = new DataSet();
                da.Fill(ds, "Workers");
                dataGrid2.ItemsSource = ds.Tables["Workers"].DefaultView;
                foreach (TextBox t in G5.Children.OfType<TextBox>())
                    t.Text = "";
            }
        }

        private void Tab5LB1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Tab5LB2.SelectedIndex = Tab5LB1.SelectedIndex;
            Tab5LB3.SelectedIndex = Tab5LB1.SelectedIndex;
        }

        private void Tab5LB2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Tab5LB1.SelectedIndex = Tab5LB2.SelectedIndex;
            Tab5LB3.SelectedIndex = Tab5LB1.SelectedIndex;
        }

        private void Tab5LB3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Tab5LB1.SelectedIndex = Tab5LB3.SelectedIndex;
            Tab5LB2.SelectedIndex = Tab5LB3.SelectedIndex;
        }

        private void Tab5B1_Click(object sender, RoutedEventArgs e)
        {
            OleDbCommand co = new OleDbCommand("INSERT INTO Workers ([Имя сотрудника],Должность,[Тип платы],Плата) VALUES ('"+Tab5T1.Text+"','"+Tab5CB1.Text+"','"+Tab5CB2.Text+"','"+Tab5T2.Text+"')", connection);
            co.ExecuteNonQuery();
            MessageBox.Show("Новый работник успешно добавлен", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Workers", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Workers");
            dataGrid2.ItemsSource = ds.Tables["Workers"].DefaultView;
        }

        private void Tab5B2_Click(object sender, RoutedEventArgs e)
        {
            OleDbCommand co = new OleDbCommand("UPDATE Workers SET [Имя сотрудника]='"+Tab5T1.Text+"', Должность='"+Tab5CB1.Text+"', [Тип платы]='"+Tab5CB2.Text+"', Плата='"+Tab5T2.Text+"' WHERE [Имя сотрудника]='"+rg+"'", connection);
            co.ExecuteNonQuery();
            MessageBox.Show("Данные работника успешно изменены", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Workers", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Workers");
            dataGrid2.ItemsSource = ds.Tables["Workers"].DefaultView;
        }

        private void Tab5B3_Click(object sender, RoutedEventArgs e)
        {
            var m = MessageBox.Show("Вы уверены, что хотите уволить этого сотрудника?", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (m == MessageBoxResult.Yes)
            {
                var m1 = MessageBox.Show("Удалить все наряды, связанные с этим работником?", "Дополнительно", MessageBoxButton.YesNo, MessageBoxImage.Question);
                OleDbCommand co = new OleDbCommand("DELETE FROM Workers WHERE [Имя сотрудника]='" + Tab5T1.Text + "'");
                co.ExecuteScalar();
                if (m1 == MessageBoxResult.Yes)
                {
                    co = new OleDbCommand("DELETE FROM Labs WHERE [Ответственный сотрудник]='" + Tab5T1.Text + "'");
                    co.ExecuteScalar();
                }
            }
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Workers", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Workers");
            dataGrid2.ItemsSource = ds.Tables["Workers"].DefaultView;
        }

        private void MenuItem_Click_7(object sender, RoutedEventArgs e)
        {
            Window5 w5 = new Window5(connection,Back,Text,Set);
            w5.Show();
        }

        private void MenuItem_Click_8(object sender, RoutedEventArgs e)
        {

        }

        private void MenuItem_Click_9(object sender, RoutedEventArgs e)
        {
            Window6 w6 = new Window6(connection,Back,Text,Set);
            w6.Show();
        }

        private void MenuItem_Click_10(object sender, RoutedEventArgs e)
        {
            Window7 w7 = new Window7(connection, Back, Text, Set);
            w7.Show();
        }

        private void MenuItem_Click_11(object sender, RoutedEventArgs e)
        {
            File.Copy($"{pathdb}", $"C:/Users/{Environment.UserName}/Desktop/JuliaDB.accdb", true);
            MessageBox.Show("База данных успешно скопирована и перемещена на рабочий стол.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void MenuItem_Click_12(object sender, RoutedEventArgs e)
        {
            if (dataGrid3.SelectedIndex != -1)
            {
                Window8 w8 = new Window8(((DataRowView)dataGrid3.SelectedItem)["Номер"].ToString(), ((DataRowView)dataGrid3.SelectedItem)["Дата"].ToString(), ((DataRowView)dataGrid3.SelectedItem)["Описание"].ToString(), ((DataRowView)dataGrid3.SelectedItem)["Заработок"].ToString(), ((DataRowView)dataGrid3.SelectedItem)["Траты"].ToString(),Back,Text,Set, connection);
                w8.ShowDialog();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Control", connection);
                DataSet ds = new DataSet();
                da.Fill(ds, "Control");
                dataGrid3.ItemsSource = ds.Tables["Control"].DefaultView;
            }
        }

        private void MenuItem_Click_13(object sender, RoutedEventArgs e)
        {
            if (dataGrid3.SelectedIndex != -1)
            {
                string g = ((DataRowView)dataGrid3.SelectedItem)["Номер"].ToString();
                MessageBoxResult res = MessageBox.Show($"Вы действительно хотите удалить наряд №{g}?", "Вы уверены", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (res == MessageBoxResult.Yes)
                {
                    OleDbCommand co = new OleDbCommand("DELETE FROM Control WHERE Номер=" + Convert.ToInt32(g), connection);
                    co.ExecuteScalar();
                    OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Control", connection);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "Control");
                    dataGrid3.ItemsSource = ds.Tables["Control"].DefaultView;
                }
            }
        }

        private void Tab6B1_Click(object sender, RoutedEventArgs e)
        {
            int mp,mm;
            try
            {
                mp = Convert.ToInt32(Tab6T5.Text);
            }
            catch
            {
                mp = 0;
            }
            try
            {
                mm = Convert.ToInt32(Tab6T6.Text);
            }
            catch
            {
                mm = 0;
            }
            OleDbCommand co = new OleDbCommand($"INSERT INTO Control (Дата, Описание, Заработок, Траты) VALUES ('" + Tab6D1.SelectedDate + "', '" + Tab6T4.Text + "', " + mp + ", " + mm + ")", connection);
            co.ExecuteNonQuery();
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Control", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Control");
            dataGrid3.ItemsSource = ds.Tables["Control"].DefaultView;
        }

        private void Tab6B2_Click(object sender, RoutedEventArgs e)
        {
            Window9 w9 = new Window9(Back, Text, Set, connection);
            w9.ShowDialog();
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Tax", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Tax");
            dataGrid4.ItemsSource = ds.Tables["Tax"].DefaultView;
        }

        private void MenuItem_Click_14(object sender, RoutedEventArgs e)
        {
            if (dataGrid4.SelectedIndex != -1)
            {
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Tax WHERE Номер=" + Convert.ToInt32(((DataRowView)dataGrid4.SelectedItem)["Номер"].ToString()), connection);
                da.Fill(dss, "Tax");
                dataGrid5.ItemsSource = dss.Tables["Tax"].DefaultView;
                vi();
            }
        }

        private void MenuItem_Click_15(object sender, RoutedEventArgs e)
        {
            if (dataGrid4.SelectedIndex != -1) 
            {
                OleDbCommand co = new OleDbCommand("DELETE FROM Tax WHERE Номер=" + Convert.ToInt32(((DataRowView)dataGrid4.SelectedItem)["Номер"].ToString()), connection);
                co.ExecuteScalar();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Tax", connection);
                DataSet ds = new DataSet();
                da.Fill(ds, "Tax");
                dataGrid4.ItemsSource = ds.Tables["Tax"].DefaultView;
            }
        }

        private void MenuItem_Click_16(object sender, RoutedEventArgs e)
        {
            if (dataGrid4.SelectedIndex!=-1)
            {
                Window10 w10 = new Window10(((DataRowView)dataGrid4.SelectedItem)["Номер"].ToString(), ((DataRowView)dataGrid4.SelectedItem)["Описание"].ToString(), ((DataRowView)dataGrid4.SelectedItem)["Тип"].ToString(), ((DataRowView)dataGrid4.SelectedItem)["Плата"].ToString(), Back, Text, Set, connection);
                w10.ShowDialog();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Tax", connection);
                DataSet ds = new DataSet();
                da.Fill(ds, "Tax");
                dataGrid4.ItemsSource = ds.Tables["Tax"].DefaultView;
            }
        }

        private void vi()
        {
            gh++;
            if (gh > 1)
            {
                double pp = 0;
                int pm = 0;
                foreach (DataRow row in dss1.Tables["Control"].Rows)
                {
                    pp += Convert.ToInt32(row["Заработок"].ToString());
                    pm += Convert.ToInt32(row["Траты"].ToString());
                }
                if (dss!=new DataSet())
                foreach (DataRow row in dss.Tables["Tax"].Rows)
                {
                    if (row["Тип"].ToString() == "Фиксированный")
                        pp -= Convert.ToDouble(row["Плата"].ToString());
                    else
                        pp -= pp*Convert.ToDouble(row["Плата"].ToString()) / 100;
                }
                Tab6T1.Text = pp.ToString();
                Tab6T2.Text = pm.ToString();
                Tab6T3.Text = (pp - pm).ToString();
            }
        }

        private void Tab6C11_Selected(object sender, RoutedEventArgs e)
        {
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Control WHERE YEAR(Дата)="+DateTime.Now.Year, connection);
            dss1 = new DataSet();
            da.Fill(dss1, "Control");
            dataGrid3.ItemsSource = dss1.Tables["Control"].DefaultView;
            vi();
        }

        private void Tab6C12_Selected(object sender, RoutedEventArgs e)
        {
            int[] mx = new int[3];
            if (DateTime.Now.Month < 4)
            {
                mx[0] = 1;
                mx[1] = 2;
                mx[2] = 3;
            }
            if (DateTime.Now.Month > 3 && DateTime.Now.Month < 7)
            {
                mx[0] = 4;
                mx[1] = 5;
                mx[2] = 6;
            }
            if (DateTime.Now.Month > 6 && DateTime.Now.Month < 10)
            {
                mx[0] = 7;
                mx[1] = 8;
                mx[2] = 9;
            }
            if (DateTime.Now.Month > 9)
            {
                mx[0] = 10;
                mx[1] = 11;
                mx[2] = 12;
            }
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Control WHERE MONTH(Дата)="+mx[0]+" OR MONTH(Дата)="+mx[1]+" OR MONTH(Дата)="+mx[2], connection);
            dss1 = new DataSet();
            da.Fill(dss1, "Control");
            dataGrid3.ItemsSource = dss1.Tables["Control"].DefaultView;
            vi();
        }

        private void Tab6C13_Selected(object sender, RoutedEventArgs e)
        {
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Control WHERE MONTH(Дата)=" + DateTime.Now.Month, connection);
            dss1 = new DataSet();
            da.Fill(dss1, "Control");
            dataGrid3.ItemsSource = dss1.Tables["Control"].DefaultView;
            vi();
        }

        private void MenuItem_Click_17(object sender, RoutedEventArgs e)
        {
            dss = new DataSet();
            OleDbDataAdapter d = new OleDbDataAdapter("SELECT * FROM Tax WHERE Номер="+1000000000000, connection);
            d.Fill(dss, "Tax");
            dataGrid5.ItemsSource = dss.Tables["Tax"].DefaultView;
            vi();
        }
    }
}
