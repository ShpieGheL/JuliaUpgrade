﻿using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using System.Data.OleDb;
using System.Data;
using System.Windows.Controls;

namespace JuliaUpgrade
{
    /// <summary>
    /// Логика взаимодействия для Window4.xaml
    /// </summary>
    public partial class Window6 : Window
    {
        DateTime d = DateTime.MinValue;
        string t;
        int b = 0;
        List<int> d1 = new List<int>();
        List<int> d2 = new List<int>();
        OleDbConnection connection;
        List<KeyValuePair<string, int>> valueList = new List<KeyValuePair<string, int>>();
        public Window6()
        {
            InitializeComponent();
            d1.Add(DateTime.Now.Month);
            d1.Add(DateTime.Now.AddMonths(-1).Month);
            d1.Add(DateTime.Now.AddMonths(-2).Month);
            d1.Add(DateTime.Now.Year);
            if (d1.Contains(12) && d1.Contains(1))
                d1.Add(DateTime.Now.AddYears(-1).Year);
            d2.Add(DateTime.Now.Month);
            d2.Add(DateTime.Now.AddMonths(-1).Month);
            d2.Add(DateTime.Now.AddMonths(-2).Month);
            d2.Add(DateTime.Now.AddMonths(-3).Month);
            d2.Add(DateTime.Now.AddMonths(-4).Month);
            d2.Add(DateTime.Now.AddMonths(-5).Month);
            d2.Add(DateTime.Now.Year);
            if (d2.Contains(12) && d2.Contains(1))
                d2.Add(DateTime.Now.AddYears(-1).Year);
        }
        public Window6(OleDbConnection con, string Back, string Text, string Set) : this()
        {
            connection = con;
            Color c = (Color)ColorConverter.ConvertFromString(Back);
            SolidColorBrush co = new SolidColorBrush(c);
            G.Background = co;
            chart.Background = co;
            if (Set == "+")
            {
                C1.Background = co;
                C8.Background = co;
            }
            c = (Color)ColorConverter.ConvertFromString(Text);
            co = new SolidColorBrush(c);
            chart.Foreground = co;
            L1.Foreground = co;
            L2.Foreground = co;
            C1.Foreground = co;
            C8.Foreground = co;
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM Labs", connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "Labs");
            foreach (DataRow row in ds.Tables["Labs"].Rows)
                if (C1.Items.Contains(row["Название клиники"].ToString()) == false)
                    C1.Items.Add(row["Название клиники"].ToString());
        }

        public void chrt()
        {
            chart.DataContext = null;
            string[] x = new string[5] { "Оплачено", "Сдано", "В работе", "Ожидание оплаты", "Долг" };
            int[] y = new int[5];
            int i = 0;
            foreach (string x1 in x)
            {
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT Comm.Цена,Comm.Номер FROM Labs INNER JOIN Comm ON Comm.Номер=Labs.Номер WHERE Статус='"+x1+"' AND Labs.[Название клиники]='"+t+"'", connection);
                DataSet ds = new DataSet();
                da.Fill(ds, "Comm");
                foreach (DataRow row in ds.Tables["Comm"].Rows)
                {
                    da = new OleDbDataAdapter("SELECT Labs.[Дата ухода] FROM Labs WHERE Labs.Номер="+Convert.ToInt32(row["Номер"].ToString()), connection);
                    DataSet ds1 = new DataSet();
                    da.Fill(ds1, "Labs");
                    foreach (DataRow row1 in ds1.Tables["Labs"].Rows)
                    {
                        if (b==0)
                            if (row1["Дата ухода"].ToString() != "" && row1["Дата ухода"].ToString() != null)
                                y[i] += Convert.ToInt32(row["Цена"].ToString());
                        if (Convert.ToDateTime(row1["Дата ухода"].ToString()).AddDays(7) >= DateTime.Now && b==1)
                            if (row1["Дата ухода"].ToString() != "" && row1["Дата ухода"].ToString() != null)
                                y[i] += Convert.ToInt32(row["Цена"].ToString());
                        if (Convert.ToDateTime(row1["Дата ухода"].ToString()).Month == DateTime.Now.Month && Convert.ToDateTime(row1["Дата ухода"].ToString()).Year == DateTime.Now.Year && b == 2)
                            if (row1["Дата ухода"].ToString() != "" && row1["Дата ухода"].ToString() != null)
                                y[i] += Convert.ToInt32(row["Цена"].ToString());
                        if (d1.Contains(Convert.ToDateTime(row1["Дата ухода"].ToString()).Month) && d1.Contains(Convert.ToDateTime(row1["Дата ухода"].ToString()).Year) && b == 3)
                            if (row1["Дата ухода"].ToString() != "" && row1["Дата ухода"].ToString() != null)
                                y[i] += Convert.ToInt32(row["Цена"].ToString());
                        if (d2.Contains(Convert.ToDateTime(row1["Дата ухода"].ToString()).Month) && d2.Contains(Convert.ToDateTime(row1["Дата ухода"].ToString()).Year) && b == 4)
                            if (row1["Дата ухода"].ToString() != "" && row1["Дата ухода"].ToString() != null)
                                y[i] += Convert.ToInt32(row["Цена"].ToString());
                        if (Convert.ToDateTime(row1["Дата ухода"].ToString()).Year == DateTime.Now.Year && b == 5)
                            if (row1["Дата ухода"].ToString() != "" && row1["Дата ухода"].ToString() != null)
                                y[i] += Convert.ToInt32(row["Цена"].ToString());
                    }
                }
                i++;
            }
            valueList.Clear();
            for (i = 0; i < x.Length; i++)
                if (y[i] != 0)
                    valueList.Add(new KeyValuePair<string, int>(x[i], y[i]));
            chart.DataContext = valueList;
        }

        private void C9_Selected(object sender, RoutedEventArgs e)
        {
            b = 0;
            chrt();
        }

        private void C10_Selected(object sender, RoutedEventArgs e)
        {
            b = 1;
            chrt();
        }

        private void C11_Selected(object sender, RoutedEventArgs e)
        {
            b = 2;
            chrt();
        }

        private void C12_Selected(object sender, RoutedEventArgs e)
        {
            b = 3;
            chrt();
        }

        private void C13_Selected(object sender, RoutedEventArgs e)
        {
            b = 4;
            chrt();
        }

        private void C14_Selected(object sender, RoutedEventArgs e)
        {
            b = 5;
            chrt();
        }

        private void C1_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            t = (sender as ComboBox).SelectedItem as string;
            chrt();
        }
    }
}
