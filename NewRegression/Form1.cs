﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NewRegression
{
    public partial class Form1 : Form
    {
        int n = 0;
        const int n_max = 100;
        double[] x = new double[n_max]; //массив зн-ний x
        double[] y = new double[n_max]; //массив экспериментальных зн-ний y 
        double[] y1 = new double[n_max]; //массив зн-ний y очередного приближения
        double[] k1 = new double[n_max];
        double[] k2 = new double[n_max];
        double[] p = new double[6];     //массив параметров
        int[] extr = new int[4];  //массив точек, подозрительных на экстремум функции
        double eps, eps1, h, h1, d;     // точность, шаг, абсолютное значение шага
        int k;                         //кол-во итераций
        public Form1()
        {
            InitializeComponent();
            dataGridView1.RowCount = 1;    //задаем кол-во строк и столбцов каждой таблицы на форме
            dataGridView1.ColumnCount = 5;
            String[] st = { "N п/п", "X", "Y", "F(x)", "(Y-F(x))^2" }; //заголовки столбцов для исходной таблицы
            for (int i = 0; i < 5; i++)
                dataGridView1.Rows[0].Cells[i].Value = st[i];
            openFileDialog1.Filter = "Text files(*.txt)|*.txt"; //в диалоге открытия файла устанавливаем фильтр только для отображения текстовых файлов
        }

        private void btnLoad_Click(object sender, EventArgs e) //загрузка данных выборки из файла
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel) //если файл не выбран
                return;
            dataGridView1.RowCount = 1;
            string filename = openFileDialog1.FileName;  // получаем выбранный файл
            String[] ss = new String[6];
            ss = System.IO.File.ReadAllLines(filename); //построчно считываем данные из текстового файла в переменную  - массив строк
            n = ss.Count(); //кол-во строк
            textBox19.Text = n.ToString();
            foreach (var line in ss)  //построчно выводим данные в таблицу, добавляя новые строки
            {
                var array = line.Split();
                dataGridView1.Rows.Add(array);
            }

            chart1.Series[0].Points.Clear();  //очистка графика от предыдущих расчетов
            chart1.Series[1].Points.Clear();
            for (int i = 0; i < n; i++)
            {
                x[i] = Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[1].Value.ToString());
                y[i] = Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[2].Value.ToString());
                chart1.Series[0].Points.AddXY(x[i], y[i]); //вывод графика по экспериментальным данным
            }
            extr = Find_max();
            //зададим начальное приближение
            int ind_min = Find_local_min(extr[1], extr[2]);
            p[0] = (y[extr[1]] + y[extr[1] - 1]) / 2; p[1] = y[ind_min] / 2; p[2] = (x[extr[1]] + x[extr[1] - 1]) / 2;
            p[3] = (y[extr[2]] + y[extr[2] - 1]) / 2; p[4] = y[ind_min] / 2; p[5] = (x[extr[2]] + x[extr[2] - 1]) / 2;
            textBox7.Text = Convert.ToString(p[0]);
            textBox8.Text = Convert.ToString(p[1]);
            textBox9.Text = Convert.ToString(p[2]);
            textBox10.Text = Convert.ToString(p[3]);
            textBox11.Text = Convert.ToString(p[4]);
            textBox12.Text = Convert.ToString(p[5]);
            btnCalc.Enabled = true;
        }


        public int[] Find_max() //поиск маскимумов табличной функции+границы отрезка
        {
            int[] ind_max = new int[4];
            int n1 = 0;
            ind_max[0] = n1;
            for (int i = 1; i < n - 1; i++)
            {
                if (y[i] > y[i - 1] && y[i] > y[i + 1])
                {
                    n1++;
                    ind_max[n1] = i;
                }
            }
            ind_max[3] = n - 1;
            return ind_max;
        }
        public int Find_local_min(int i1, int i2)  //поиск локального минимума табличной цункции на заданном отрезке
        {
            int min = 0;
            for (int i = i1; i < i2; i++)
            {
                if (y[i] < y[i - 1] && y[i] < y[i + 1])
                {
                    min = i;
                }
            }
            return min;
        }

        public double[] f(double[] x1, double[] a)      // расчет массива значений аппроксимирующей функции
        {
            double[] Y = new double[n];
            for (int i = 0; i < n; i++)
            {
                double p1 = -a[1] * (x1[i] - a[2]) * (x1[i] - a[2]);
                double p2 = -a[4] * (x1[i] - a[5]) * (x1[i] - a[5]);
                Y[i] = a[0] * Math.Exp(p1) + a[3] * Math.Exp(p2);
            }
            return Y;
        }

        public double[] Fk1(double[] x1, double[] a)      // расчет массива значений аппроксимирующей функции
        {
            double[] T1 = new double[n];
            for (int i = 0; i < n; i++)
            {
                double p1 = -a[1] * (x1[i] - a[2]) * (x1[i] - a[2]);
                //double p2 = -a[4] * (x1[i] - a[5]) * (x1[i] - a[5]);
                T1[i] = a[0] * Math.Exp(p1);
            }
            return T1;
        }

        public double[] Fk2(double[] x1, double[] a)      // расчет массива значений аппроксимирующей функции
        {
            double[] T2 = new double[n];
            for (int i = 0; i < n; i++)
            {
                //double p1 = -a[1] * (x1[i] - a[2]) * (x1[i] - a[2]);
                double p2 = -a[4] * (x1[i] - a[5]) * (x1[i] - a[5]);
                T2[i] = a[3] * Math.Exp(p2);
            }
            return T2;
        }

        private void btnCalc_Click(object sender, EventArgs e)
            {
                double[] kv = new double[n]; //массив квадратичных отклонений
                double[] kv1 = new double[n];
                double[] kv2 = new double[n];
                //начальное приближение
                p[0] = Convert.ToDouble(textBox7.Text.Trim());
                p[1] = Convert.ToDouble(textBox8.Text.Trim());
                p[2] = Convert.ToDouble(textBox9.Text.Trim());
                p[3] = Convert.ToDouble(textBox10.Text.Trim());
                p[4] = Convert.ToDouble(textBox11.Text.Trim());
                p[5] = Convert.ToDouble(textBox12.Text.Trim());
                //расчет методом покоординатного спуска
                eps = 0.001; //погрешность
                h = 0.005; //шаг поиска - начальное значение
                k = 50; //кол-во итераций
                eps1 = eps / k;
                do
                {
                    d = Math.Abs(h);
                    for (int i = 0; i < 6; i++)
                    {
                        h1 = h;
                        scan(i);
                    }
                    h = h / k;
                }
                while (d > eps);

                y1 = f(x, p);
                for (int i = 0; i < n; i++)
                    kv[i] = (y[i] - y1[i]) * (y[i] - y1[i]);
                k1 = Fk1(x, p);
                for (int i = 0; i < n; i++)
                    kv1[i] = (y[i] - k1[i]) * (y[i] - k1[i]);
                k2 = Fk2(x, p);
                for (int i = 0; i < n; i++)
                    kv2[i] = (y[i] - k2[i]) * (y[i] - k2[i]);


                //вывод расчетных коэф-тов функции
                textBox1.Text = p[0].ToString("F6");
                textBox2.Text = p[1].ToString("F6");
                textBox3.Text = p[2].ToString("F6");
                textBox4.Text = p[3].ToString("F6");
                textBox5.Text = p[4].ToString("F6");
                textBox6.Text = p[5].ToString("F6");
                //вывод расчетных значений функции и квадр.отклонений
                chart1.Series[1].Points.Clear();
                chart1.Series[2].Points.Clear();
                chart1.Series[3].Points.Clear();
                for (int i = 0; i < n; i++)
                {
                    dataGridView1.Rows[i + 1].Cells[3].Value = y1[i].ToString("F6");
                    dataGridView1.Rows[i + 1].Cells[4].Value = kv[i].ToString("F6");
                    //вывод графиков экспериментальных и аппроксимирующих значений функции
                    chart1.Series[1].Points.AddXY(x[i], y1[i]);
                    chart1.Series[2].Points.AddXY(x[i], k1[i]);
                    chart1.Series[3].Points.AddXY(x[i], k2[i]);
                }
                textBox13.Text = f_o().ToString("F6"); //вывод суммы квадратов отклонений
                double kv_err = Math.Sqrt(f_o()) / n;  //расчет среднекв. ошибки
                textBox14.Text = kv_err.ToString("F6");
                textBox18.Visible = true; label19.Visible = true;
                textBox17.Visible = true; label18.Visible = true;
                textBox16.Visible = true; label17.Visible = true;
                textBox15.Visible = true; label16.Visible = true;
                //вывод максимумов
                double xm = p[2];
                double ym = p[0] * Math.Exp(-p[1] * (xm - p[2]) * (xm - p[2])) + p[3] * Math.Exp(-p[4] * (xm - p[5]) * (xm - p[5]));
                textBox15.Text = xm.ToString("F3");
                textBox16.Text = ym.ToString("F6");
                xm = p[5];
                ym = p[0] * Math.Exp(-p[1] * (xm - p[2]) * (xm - p[2])) + p[3] * Math.Exp(-p[4] * (xm - p[5]) * (xm - p[5]));
                textBox17.Text = xm.ToString("F3");
                textBox18.Text = ym.ToString("F6");
            }

            public double f_o()// вычисление суммы отклонений - целевая функция
        {
            double sum = 0;
            for (int i = 0; i < n; i++)
                sum += (y[i] - y1[i]) * (y[i] - y1[i]);
            return sum;
        }
        public void scan(int nom)  //оптимизация одномерной функции
        {
            Boolean a;
            Double z, z1, d1;
            y1 = f(x, p);
            z = f_o();
            do
            {
                d1 = Math.Abs(h1);
                p[nom] = p[nom] + h1;
                y1 = f(x, p);
                z1 = f_o();
                a = (z1 >= z);
                if (a == true) h1 = -h1 / k;
                z = z1;
            }
            while (a == false && d1 > eps1);
        }
        

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
        }
    }
}
