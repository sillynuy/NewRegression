using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization;
using System.Windows.Forms.DataVisualization.Charting;

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
        double[] k1sh = new double[n_max];
        double[] k2 = new double[n_max];
        double[] k2sh = new double[n_max];
        double[] p = new double[6];     //массив параметров
        double[] psh = new double[6];
        double[] Ysh = new double[n_max];
        double[] Ysh2 = new double[n_max];
        int[] extr = new int[4];  //массив точек, подозрительных на экстремум функции
        double eps, eps1, h, h1, d, IPS, T3, IPR, T5, maf, F30;     // точность, шаг, абсолютное значение шага
        int k;                         //кол-во итераций
        public Form1()
        {
            InitializeComponent();
            dataGridView1.RowCount = 1;    //задаем кол-во строк и столбцов каждой таблицы на форме
            dataGridView1.ColumnCount = 7;
            String[] st = { "N п/п", "X", "Y", "F(x)", "(Y-F(x))^2", "F'", "F''" }; //заголовки столбцов для исходной таблицы
            for (int i = 0; i < 7; i++)
                dataGridView1.Rows[0].Cells[i].Value = st[i];
            openFileDialog1.Filter = "Text files(*.txt)|*.txt"; //в диалоге открытия файла устанавливаем фильтр только для отображения текстовых файлов
            chart1.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
            //this.MouseMove += chart1_MouseMove;

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
            chart1.Series[2].Points.Clear();
            chart1.Series[3].Points.Clear();
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

        public double[] F(double[] x1, double[] a)      // расчет массива значений аппроксимирующей функции
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

        public double[] Fk1sh(double[] x1, double[] a)      // расчет массива значений аппроксимирующей функции
        {
            double[] T1 = new double[n];
            for (int i = 0; i < n; i++)
            {
                double p1 = -a[1] * ((x1[i] - a[2]) * (x1[i] - a[2]));
                //double p2 = -a[4] * (x1[i] - a[5]) * (x1[i] - a[5]);
                T1[i] = (-2) * (a[0] * a[1] * (x1[i] - a[2])) * Math.Exp(p1);
            }
            return T1;
        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            
            Point mousePoint = new Point(e.X, e.Y);

            chart1.ChartAreas[0].CursorX.Interval = 0;
            chart1.ChartAreas[0].CursorY.Interval = 0;

            chart1.ChartAreas[0].CursorX.SetCursorPixelPosition(mousePoint, true);
            chart1.ChartAreas[0].CursorY.SetCursorPixelPosition(mousePoint, true);

            HitTestResult result = chart1.HitTest(e.X, e.Y);
            try
            {

                double xValue = chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.X);
                double yValue = chart1.ChartAreas[0].AxisY.PixelPositionToValue(e.Y);

                if (result.PointIndex > -1 && result.ChartArea != null)
                {
                    label35.Text = "Y-value: " + result.Series.Points[result.PointIndex].YValues[0].ToString();
                    label35.Location = new Point(1400, e.Y + 25);
                }



                label34.Text = string.Concat(string.Concat(Math.Round(xValue, 2).ToString(), " , "), Math.Round(yValue, 2).ToString());
                label34.Location = new Point(1400, e.Y - 5);
            }
            catch
            {

            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                chart1.Series[0].Enabled = true;
                chart1.Series[0].IsVisibleInLegend = true;
            }
            else
            {
                chart1.Series[0].Enabled = false;
                chart1.Series[0].IsVisibleInLegend = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                chart1.Series[1].Enabled = true;
                chart1.Series[1].IsVisibleInLegend = true;
            }
            else
            {
                chart1.Series[1].Enabled = false;
                chart1.Series[1].IsVisibleInLegend = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                chart1.Series[2].Enabled = true;
                chart1.Series[2].IsVisibleInLegend = true;
            }
            else
            {
                chart1.Series[2].Enabled = false;
                chart1.Series[2].IsVisibleInLegend = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                chart1.Series[3].Enabled = true;
                chart1.Series[3].IsVisibleInLegend = true;
            }
            else
            {
                chart1.Series[3].Enabled = false;
                chart1.Series[3].IsVisibleInLegend = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                chart1.Series[4].Enabled = true;
                chart1.Series[4].IsVisibleInLegend = true;
            }
            else
            {
                chart1.Series[4].Enabled = false;
                chart1.Series[4].IsVisibleInLegend = false;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                chart1.Series[5].Enabled = true;
                chart1.Series[5].IsVisibleInLegend = true;
            }
            else
            {
                chart1.Series[5].Enabled = false;
                chart1.Series[5].IsVisibleInLegend = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        /*private void button2_Click(object sender, EventArgs e)
        {
            Chart chart1 = new Chart();
            chart1.Serializer.Save(stream);


            //new Form().Show();
        }*/

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

        public double[] Fk2sh(double[] x1, double[] a)      // расчет массива значений аппроксимирующей функции
        {
            double[] T2 = new double[n];
            for (int i = 0; i < n; i++)
            {
                double p2 = a[4] * ((x1[i] - a[5]) * (x1[i] - a[5]));
                //double p2 = -a[4] * (x1[i] - a[5]) * (x1[i] - a[5]);
                T2[i] = (-2) * (a[3] * a[4] * (x1[i] - a[5]) * Math.Exp(p2) );
            }
            return T2;
        }

        public double[] Fsh(double[] x1, double[] a)      // расчет массива значений аппроксимирующей функции
        {
            double[] Ysh = new double[n];
            for (int i = 0; i < n; i++)
            {
                Ysh[i] = (-2) * (a[0] * a[1] * (x1[i] - a[2])) *
                    Math.Exp(-a[1] * ((x1[i] - a[2]) * (x1[i] - a[2])))
                    - (2 * a[3] * a[4] * (x1[i] - a[5])
                    * Math.Exp(-a[4] * ((x1[i] - a[5]) * (x1[i] - a[5]))));
            }



            return Ysh;
        }

        public double[] Fsh2(double[] x1, double[] a)      // расчет массива значений аппроксимирующей функции
        {
            for (int i = 0; i < n; i++)
            {
                double p1 = -a[1] * Math.Pow((x1[i] - a[2]),2);
                double p2 = -a[4] * Math.Pow((x1[i] - a[5]),2);
                Ysh2[i] = ( (-2) * (a[0] * a[1] * Math.Exp(p1))) +
                    (4 * a[0] * Math.Pow(a[1],2) * Math.Pow((x1[i] - a[2]),2) * Math.Exp(p1)) -
                    (2 * (a[3] * a[4] * Math.Exp(p2))) +
                    (4 * a[3] * Math.Pow(a[4],2) * Math.Pow((x1[i] - a[5]), 2) * Math.Exp(p2));
            }



            return Ysh2;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*p[0] = Convert.ToDouble(textBox1.Text.Trim());
            p[1] = Convert.ToDouble(textBox2.Text.Trim());
            p[2] = Convert.ToDouble(textBox3.Text.Trim());
            p[3] = Convert.ToDouble(textBox4.Text.Trim());
            p[4] = Convert.ToDouble(textBox5.Text.Trim());
            p[5] = Convert.ToDouble(textBox6.Text.Trim());
            chart1.Series[0].Enabled = false;
            chart1.Series[2].Enabled = false;
            chart1.Series[3].Enabled = false;*/
        }

        private void btnCalc_Click(object sender, EventArgs e)
        {
            try
            {
                double[] kv = new double[n]; //массив квадратичных отклонений
                double[] kvsh = new double[n];
                double[] kv1 = new double[n];
                double[] kv1sh = new double[n];
                double[] kv2 = new double[n];
                double[] kv2sh = new double[n];
                //начальное приближение
                p[0] = Convert.ToDouble(textBox7.Text.Trim());
                p[1] = Convert.ToDouble(textBox8.Text.Trim());
                p[2] = Convert.ToDouble(textBox9.Text.Trim());
                p[3] = Convert.ToDouble(textBox10.Text.Trim());
                p[4] = Convert.ToDouble(textBox11.Text.Trim());
                p[5] = Convert.ToDouble(textBox12.Text.Trim());

                /*psh[0] = Convert.ToDouble(textBox7.Text.Trim());
                psh[1] = Convert.ToDouble(textBox8.Text.Trim());
                psh[2] = Convert.ToDouble(textBox9.Text.Trim());
                psh[3] = Convert.ToDouble(textBox10.Text.Trim());
                psh[4] = Convert.ToDouble(textBox11.Text.Trim());
                psh[5] = Convert.ToDouble(textBox12.Text.Trim());*/

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

                y1 = F(x, p);
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

                Ysh = Fsh(x, p);
                /*for (int i = 0; i < n; i++)
                    kvsh[i] = (y[i] - Ysh[i]) * (y[i] - Ysh[i]);
                k1sh = Fk1sh(x, psh);
                for (int i = 0; i < n; i++)
                    kv1sh[i] = (y[i] - k1sh[i]) * (y[i] - k1sh[i]);
                k2sh = Fk2sh(x, psh);
                for (int i = 0; i < n; i++)
                    kv2sh[i] = (y[i] - k2sh[i]) * (y[i] - k2sh[i]);*/

                /*textBox28.Text = psh[0].ToString("F6");
                textBox29.Text = psh[1].ToString("F6");
                textBox30.Text = psh[2].ToString("F6");
                textBox31.Text = psh[3].ToString("F6");
                textBox32.Text = psh[4].ToString("F6");
                textBox33.Text = psh[5].ToString("F6");*/

                Ysh2 = Fsh2(x, p);

                //вывод расчетных значений функции и квадр.отклонений
                chart1.Series[1].Points.Clear();
                chart1.Series[2].Points.Clear();
                chart1.Series[3].Points.Clear();
                chart1.Series[4].Points.Clear();
                chart1.Series[5].Points.Clear();
                for (int i = 0; i < n; i++)
                {
                    dataGridView1.Rows[i + 1].Cells[3].Value = y1[i].ToString("F6");
                    dataGridView1.Rows[i + 1].Cells[4].Value = kv[i].ToString("F6");
                    dataGridView1.Rows[i + 1].Cells[5].Value = Ysh[i].ToString("F6");
                    dataGridView1.Rows[i + 1].Cells[6].Value = Ysh2[i].ToString("F6");
                    //вывод графиков экспериментальных и аппроксимирующих значений функции
                    chart1.Series[1].Points.AddXY(x[i], y1[i]);
                    chart1.Series[2].Points.AddXY(x[i], k1[i]);
                    chart1.Series[3].Points.AddXY(x[i], k2[i]);
                    chart1.Series[4].Points.AddXY(x[i], Ysh[i]);
                    chart1.Series[5].Points.AddXY(x[i], Ysh2[i]);
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

                double testt = Ysh.Max();

                //textBox21.Text = testt.ToString();
                int max = Array.IndexOf(Ysh, Ysh.Max());
                ///textBox20.Text = x.GetValue(max).ToString();
                xm = p[2];
                ym = (-2) * (p[0] * p[1] * (xm - p[2])) *
                    Math.Exp(-p[1] * ((xm - p[2]) * (xm - p[2])))
                    - (2 * p[3] * p[4] * (xm - p[5]) *
                    Math.Exp(-p[4] * ((xm - p[5]) * (xm - p[5]))));
                textBox20.Text = xm.ToString();
                textBox21.Text = ym.ToString();

                xm = p[5];
                ym = (-2) * (p[0] * p[1] * (xm - p[2])) *
                    Math.Exp(-p[1] * ((xm - p[2]) * (xm - p[2])))
                    - (2 * p[3] * p[4] * (xm - p[5]) *
                    Math.Exp(-p[4] * ((xm - p[5]) * (xm - p[5]))));
                textBox22.Text = xm.ToString();
                textBox23.Text = ym.ToString();


                /*var dataPoint = chart1.Series["Series5"].Points.FindByValue(1.0, "Y");
                var xxxx = dataPoint.XValue;

                textBox20.Text = xxxx.ToString();*/
                /*var dataPoint2 = chart1.Series["Series6"].Points.FindMinByValue();
                double yyyy = dataPoint2.XValue;*/
                /*textBox23.Text = Ysh.Min().ToString();*/
                int min = Array.IndexOf(Ysh, Ysh.Min());
                /*textBox22.Text = x.GetValue(min).ToString();*/
                //textBox22.Text = yyyy.ToString();


                T3 = Convert.ToDouble(y1.GetValue(max));
                T5 = Convert.ToDouble(y1.GetValue(min));

                double max1 = Convert.ToDouble(Ysh.GetValue(max));
                double min1 = Convert.ToDouble(Ysh.GetValue(min));

                double div = Ysh.Min() / Ysh.Max();

                IPS = max1 / T3;
                IPR = Math.Abs(min1 / T5);

                textBox24.Text = IPS.ToString();
                textBox25.Text = IPR.ToString();
                textBox26.Text = Math.Abs(div).ToString();

                double m1 = Convert.ToDouble(textBox16.Text.Trim());
                double m2 = Convert.ToDouble(textBox18.Text.Trim());

                if (m1 > m2)
                    maf = m1;
                else
                    maf = m2;
                //double T3 = max1 / test;

                F30 = maf * 0.3;
                textBox27.Text = F30.ToString();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Сперва нужно загрузить файлы" + Environment.NewLine + 
                    ex.Message);
            }
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
            y1 = F(x, p);
            z = f_o();
            do
            {
                d1 = Math.Abs(h1);
                p[nom] = p[nom] + h1;
                y1 = F(x, p);
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


        private void chart1_MouseWheel(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Delta < 0)
                {
                    for (int i = 0; i < 6; i++)
                    {
                        chart1.Series[i].IsVisibleInLegend = true;
                    }
                    chart1.ChartAreas[0].AxisX.ScaleView.ZoomReset();
                    chart1.ChartAreas[0].AxisY.ScaleView.ZoomReset();
                }

                if (e.Delta > 0)
                {
                    for (int i = 0; i < 6; i++)
                    {
                        chart1.Series[i].IsVisibleInLegend = false;
                    }                        
                    
                    double xMin = chart1.ChartAreas[0].AxisX.ScaleView.ViewMinimum;
                    double xMax = chart1.ChartAreas[0].AxisX.ScaleView.ViewMaximum;
                    double yMin = chart1.ChartAreas[0].AxisY.ScaleView.ViewMinimum;
                    double yMax = chart1.ChartAreas[0].AxisY.ScaleView.ViewMaximum;

                    double posXStart = chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) - (xMax - xMin) / 4;
                    double posXFinish = chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) + (xMax - xMin) / 4;
                    double posYStart = chart1.ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) - (yMax - yMin) / 4;
                    double posYFinish = chart1.ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) + (yMax - yMin) / 4;

                    chart1.ChartAreas[0].AxisX.ScaleView.Zoom(posXStart, posXFinish);
                    chart1.ChartAreas[0].AxisY.ScaleView.Zoom(posYStart, posYFinish);
                }
            }
            catch { }
        }
    }
}
