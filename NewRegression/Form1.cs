using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Common;
using System.Data.SQLite;


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
        double[] Ysh = new double[n_max]; //производная 1
        double[] Ysh2 = new double[n_max]; //производная 2
        int[] extr = new int[4];  //массив точек, подозрительных на экстремум функции
        double eps, eps1, h, h1, d, IPS, T3, IPR, T5, maf, F30;     // точность, шаг, абсолютное значение шага
        int k;                         //кол-во итераций

        string ConnectionString = "Data Source = MyDatabase.sqlite; Version=3;";

        double SelectionBorder1 = 0;
        double SelectionBorder2 = 0;
        int SelectionClicks = 1;
        string filename = "";

        string[] RawDataFull;
        string[] RawDataSelection;

        //ПОЛНЫЕ ДАННЫЕ
        double[] Col1;
        double[] Col2;
        double[] Col3;
        double[] Col4;
        double[] Col5;
        //ВЫБОРКА
        double[] Sel1;
        double[] Sel2;
        double[] Sel3;
        double[] Sel4;
        double[] Sel5;

        double[] y2;
        double[] y3;
        double[] y4;

        //СТЕПЕНЬ СЖАТИЯ
        int CompressionRateOverview;
        int CompressionRateSelection = 1;

        public string name;

        SQLiteConnection m_dbConnection;

        double[] dInitialApp;
        double[] dAFE = new double[4];
        double[] dFE = new double[4];
        double[] dCoeffs = new double[6];
        double[] dIndexes = new double[4];

        double dESMax1;
        double dESMax2;

        string sInitialApp;
        string sAFExtremums;
        string sFExtremums;
        string sCoeffs;
        string sIndexes;
        string sYsh;
        string sYsh2;

        string sESMax1;
        string sESMax2;

        string sX;
        string sY1;
        string sY2;
        string sY3;
        string sY4;

        public Form1()
        {
            InitializeComponent();
            dataGridView1.RowCount = 11;    //задаем кол-во строк и столбцов каждой таблицы на форме
            dataGridView1.ColumnCount = 7;
            String[] st = { "N п/п", "X", "Y", "F(x)", "(Y-F(x))^2", "F'", "F''" }; //заголовки столбцов для исходной таблицы
            for (int i = 0; i < 7; i++)
                //dataGridView1.Rows[0].Cells[i].Value = st[i];
                openFileDialog1.Filter = "Text files(*.txt)|*.txt"; //в диалоге открытия файла устанавливаем фильтр только для отображения текстовых файлов
            chart1.ChartAreas[0].AxisX.ScaleView.Zoomable = true;

            Select.Enabled = false;
            SelectBack.Enabled = false;
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
            /*chart1.Series[3].Points.Clear();
            for (int i = 0; i < n; i++)
            {
                x[i] = Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[1].Value.ToString());
                y[i] = Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[2].Value.ToString());
                chart1.Series[0].Points.AddXY(x[i], y[i]); //вывод графика по экспериментальным данным
            }*/

            //InitialApproximation();

            btnCalc.Enabled = true;
        }

        public void InitialApproximation()
        {
            extr = Find_max();
            //зададим начальное приближение
            int ind_min = Find_local_min(extr[1], extr[2]);
            p[0] = (y[extr[1]] + y[extr[1] - 1]) / 2;
            p[1] = y[ind_min] / 2;
            p[2] = (x[extr[1]] + x[extr[1] - 1]) / 2;
            p[3] = (y[extr[2]] + y[extr[2] - 1]) / 2;
            p[4] = y[ind_min] / 2;
            p[5] = (x[extr[2]] + x[extr[2] - 1]) / 2;
            textBox7.Text = Convert.ToString(p[0]);
            textBox8.Text = Convert.ToString(p[1]);
            textBox9.Text = Convert.ToString(p[2]);
            textBox10.Text = Convert.ToString(p[3]);
            textBox11.Text = Convert.ToString(p[4]);
            textBox12.Text = Convert.ToString(p[5]);

            dInitialApp = p;
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

        public double[] fsh(double[] x1, double[] a)      // расчет массива значений аппроксимирующей функции
        {
            Ysh = new double[n];
            for (int i = 0; i < n; i++)
            {
                double p1 = -a[1] * ((x1[i] - a[2]) * (x1[i] - a[2]));
                //double p1 = (-a[1] * (x[i] - a[2]) ) * (-a[1] * (x[i] - a[2]));
                double p2 = -a[4] * ((x1[i] - a[5]) * (x1[i] - a[5]));
                Ysh[i] = (-2) * (a[0] * a[1] * (x1[i] - a[2])) * Math.Exp(p1) - (2 * a[3] * a[4] * (x1[i] - a[5]) * Math.Exp(p2));
                //Ysh[i] = (-2) * (a[0] *a[1] * (x1[i] -a[2]) * Math.Exp(a[1] * ((x1[i] -a[2]) * (x1[i] - a[2]))))+(-2) * (a[3] * a[4] * (x1[i] -a[5]) * Math.Exp(a[4] * ((x1[i] -a[5]) * (x1[i] -a[5]))));
            }

            return Ysh;
        }

        public double[] fsh2(double[] x1, double[] a)      // расчет массива значений аппроксимирующей функции
        {
            Ysh2 = new double[n];
            for (int i = 0; i < n; i++)
            {
                double p1 = -a[1] * Math.Pow((x1[i] - a[2]), 2);
                double p2 = -a[4] * Math.Pow((x1[i] - a[5]), 2);
                Ysh2[i] = ((-2) * (a[0] * a[1] * Math.Exp(p1))) +
                    (4 * a[0] * Math.Pow(a[1], 2) * Math.Pow((x1[i] - a[2]), 2) * Math.Exp(p1)) -
                    (2 * (a[3] * a[4] * Math.Exp(p2))) +
                    (4 * a[3] * Math.Pow(a[4], 2) * Math.Pow((x1[i] - a[5]), 2) * Math.Exp(p2));
            }

            return Ysh2;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            p[0] = Convert.ToDouble(textBox1.Text.Trim());
            p[1] = Convert.ToDouble(textBox2.Text.Trim());
            p[2] = Convert.ToDouble(textBox3.Text.Trim());
            p[3] = Convert.ToDouble(textBox4.Text.Trim());
            p[4] = Convert.ToDouble(textBox5.Text.Trim());
            p[5] = Convert.ToDouble(textBox6.Text.Trim());
            chart1.Series[0].Enabled = false;
            chart1.Series[2].Enabled = false;
            chart1.Series[3].Enabled = false;
            /*double xm = p[2];
            double ym = (-2) * (p[0] * p[1] * (xm - p[2])) * Math.Exp(-p[1] * ((xm - p[2]) * (xm - p[2]))) - (2 * p[3] * p[4] * (xm - p[5]) * Math.Exp(-p[4] * ((xm - p[5]) * (xm - p[5]))));
            textBox20.Text = xm.ToString("F3");
            textBox21.Text = ym.ToString("F6");
            xm = p[5];
            ym = (-2) * (p[0] * p[1] * (xm - p[2])) * Math.Exp(-p[1] * ((xm - p[2]) * (xm - p[2]))) - (2 * p[3] * p[4] * (xm - p[5]) * Math.Exp(-p[4] * ((xm - p[5]) * (xm - p[5]))));
            textBox22.Text = xm.ToString("F3");
            textBox23.Text = ym.ToString("F6");*/

        }

        private void btnCalc_Click(object sender, EventArgs e)
        {
            chart1.Visible = true;
            chart1.Size = new Size(915, 343);
            chart2.Size = new Size(915, 343);

            textBox28.Text = "";
            textBox29.Text = "";

            InitialApproximation();

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

            dCoeffs = p;

            Ysh = fsh(x, p);
            Ysh2 = fsh2(x, p);
            //вывод расчетных значений функции и квадр.отклонений
            /*chart1.Series[1].Points.Clear();
            chart1.Series[2].Points.Clear();
            chart1.Series[3].Points.Clear();
            chart1.Series[4].Points.Clear();
            chart1.Series[5].Points.Clear();*/
            /*for (int i = 0; i < n; i++)
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
            }*/

            //вывод производных на график
            chart2.Series[0].Name = "df1";
            chart2.Series.Add("df2");
            chart2.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart2.Series[1].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart2.Series[0].Points.DataBindXY(x, Ysh);
            chart2.Series[1].Points.DataBindXY(x, Ysh2);
            chart2.ChartAreas[0].InnerPlotPosition.X = 16;
            chart2.ChartAreas[0].InnerPlotPosition.Y = 8;
            chart2.ChartAreas[0].InnerPlotPosition.Width = 90;
            chart2.ChartAreas[0].InnerPlotPosition.Height = 90;
            chart1.ChartAreas[0].InnerPlotPosition.X = 16;
            chart1.ChartAreas[0].InnerPlotPosition.Y = 8;
            chart1.ChartAreas[0].InnerPlotPosition.Width = 90;
            chart1.ChartAreas[0].InnerPlotPosition.Height = 90;

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
            dAFE[0] = xm;
            dAFE[1] = ym;
            xm = p[5];
            ym = p[0] * Math.Exp(-p[1] * (xm - p[2]) * (xm - p[2])) + p[3] * Math.Exp(-p[4] * (xm - p[5]) * (xm - p[5]));
            textBox17.Text = xm.ToString("F3");
            textBox18.Text = ym.ToString("F6");
            dAFE[2] = xm;
            dAFE[3] = ym;

            textBox21.Text = Ysh.Max().ToString();
            dFE[1] = Convert.ToDouble(Ysh.Max());
            int max = Array.IndexOf(Ysh, Ysh.Max());
            textBox20.Text = x.GetValue(max).ToString();
            dFE[0] = Convert.ToDouble(x.GetValue(max));

            textBox23.Text = Ysh.Min().ToString();
            dFE[3] = Convert.ToDouble(Ysh.Min());
            int min = Array.IndexOf(Ysh, Ysh.Min());
            textBox22.Text = x.GetValue(min).ToString();
            dFE[2] = Convert.ToDouble(x.GetValue(min));

            T3 = Convert.ToDouble(y1.GetValue(max));
            T5 = Convert.ToDouble(y1.GetValue(min));

            double max1 = Convert.ToDouble(Ysh.GetValue(max));
            double min1 = Convert.ToDouble(Ysh.GetValue(min));

            double div = Ysh.Min() / Ysh.Max();

            IPS = max1 / T3;
            IPR = Math.Abs(min1 / T5);

            textBox24.Text = IPS.ToString();
            dIndexes[0] = Convert.ToDouble(IPS);
            textBox25.Text = IPR.ToString();
            dIndexes[1] = Convert.ToDouble(IPR);
            textBox26.Text = Math.Abs(div).ToString();
            dIndexes[2] = Convert.ToDouble(Math.Abs(div));

            double m1 = Convert.ToDouble(textBox16.Text.Trim());
            double m2 = Convert.ToDouble(textBox18.Text.Trim());

            if (m1 > m2)
                maf = m1;
            else
                maf = m2;
            //double T3 = max1 / test;

            F30 = maf * 0.3;
            textBox27.Text = F30.ToString();
            dIndexes[3] = Convert.ToDouble(F30);
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

            chart1.Visible = true;

            //DBCreate();

            double[] temp1 = { 20.3, 17.9, 65.36, 20.1, 9.8, 4.7 };
            string temp2 = "14.7/56/2.8/9.3/7.1/3.63/9.65";

            //double[] temp11 = StringToArray(temp2);
            //string temp22 = ArrayToString(temp1);
        }

        private void DBCreate()
        {
            //создание новой базы
            SQLiteConnection.CreateFile("MyDatabase.sqlite");

            //создание коннекшна
            m_dbConnection = new SQLiteConnection("Data Source=MyDatabase.sqlite;Version=3;");

            //открытие коннекшна
            m_dbConnection.Open();

            DBQueryExecute(@"
                CREATE TABLE Entry(
                id_entry integer PRIMARY KEY, 
                datemark varchar(20), 
                name varchar(60)
                )
                ");
            DBQueryExecute(@"
                CREATE TABLE OrdinaryContraction(
                id_entry integer PRIMARY KEY REFERENCES Entry(id_entry), 
                InitialApproximation text, 
                AFExtremums text, 
                FExtremums text, 
                Coeffs text, 
                Indexes text,
                Derivative1 text,
                Derivative2 text
                )
                ");
            DBQueryExecute(@"
                CREATE TABLE ExtrasystolicContraction(
                id_entry integer PRIMARY KEY REFERENCES Entry(id_entry),
                Max1 real, 
                Max2 real
                )
                ");
            DBQueryExecute(@"
                CREATE TABLE Source(
                id_entry integer PRIMARY KEY REFERENCES Entry(id_entry), 
                x string, 
                y1 string, 
                y2 string, 
                y3 string, 
                y4 string
                )
                ");

            //закрытие коннекшна
            m_dbConnection.Close();
        }

        private void DBQueryExecute(string str)
        {
            SQLiteCommand command = new SQLiteCommand(str, m_dbConnection);
            command.ExecuteNonQuery();
        }

        private void chart1_MouseWheel(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Delta < 0)
                {
                    chart1.ChartAreas[0].AxisX.ScaleView.ZoomReset();
                    chart1.ChartAreas[0].AxisY.ScaleView.ZoomReset();
                }

                if (e.Delta > 0)
                {
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

        private void chart1_MouseDown(object sender, MouseEventArgs e)
        {
            //double datapoint = Convert.ToDouble(chart1.ChartAreas[0].AxisX.GetPosition(50));
            double ClickPos = Convert.ToDouble(chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.X));

            if (ClickPos < Col1.Min()) ClickPos = Col1.Min();
            else if (ClickPos > Col1.Max()) ClickPos = Col1.Max();

            SelectionClicks++;

            if (SelectionClicks % 2 == 1) SelectionBorder1 = ClickPos;
            else SelectionBorder2 = ClickPos;
            //label3.Text = Convert.ToString(chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.Y));

            if (SelectionBorder1 > SelectionBorder2)
            {
                SelectionBorder1 = SelectionBorder1 - SelectionBorder2;
                SelectionBorder2 += SelectionBorder1;
                SelectionBorder1 = SelectionBorder2 - SelectionBorder1;
            }

            if (SelectionBorder1 * SelectionBorder2 == 0)
                Select.Enabled = true;

            //ShowSelectionBorder(SelectionBorder1, SelectionBorder2);

            lblBorderInfo.Text = "LEFT " + Convert.ToString(Math.Round(SelectionBorder1, 3)) + ", RIGHT " + Convert.ToString(Math.Round(SelectionBorder2, 3));
        }

        private void btnOpenDat_Click(object sender, EventArgs e)
        {
            CompressionRateOverview = Convert.ToInt32(tbCompRate.Text);

            // Displays an OpenFileDialog so the user can select a Cursor.  
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Файлы DAT|*.dat";
            openFileDialog2.Title = "Выберите DAT-файл";

            if (openFileDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename = openFileDialog2.FileName;
                textBox19.Text = filename;
            }

            if (!string.IsNullOrEmpty(filename) && File.Exists(filename))
            {
                //извлечение ВСЕХ данных
                RawDataFull = ExtractFromFile(filename, CompressionRateOverview);

                //GLOBALS SHOULD BE ASSIGNED HERE
                Col1 = CutArrayFromRows(RawDataFull, 0);
                Col2 = CutArrayFromRows(RawDataFull, 1);
                Col3 = CutArrayFromRows(RawDataFull, 2);
                Col4 = CutArrayFromRows(RawDataFull, 3);
                Col5 = CutArrayFromRows(RawDataFull, 4);
                y = CutArrayFromRows(RawDataFull, 1);
                x = CutArrayFromRows(RawDataFull, 0);
                n = y.Length;

                //построение графика
                //BuildGraph(RawDataFull, false, true, true, true, true);
                BuildGraph2(false, true, true, true, true);
            }
        }

        private void SelectBack_Click(object sender, EventArgs e)
        {
            //BuildGraph(RawDataFull, false, true, true, true, true);
            BuildGraph2(false, true, true, true, true);
        }

        private void Select_Click(object sender, EventArgs e)
        {
            SelectBack.Enabled = true;
            int[] Borders = FindSelectionBorders(filename, SelectionBorder1, SelectionBorder2);
            //MessageBox.Show(Convert.ToString(Borders[0]) + ", " + Convert.ToString(Borders[1]));
            SelectionBorder1 = Borders[0];
            SelectionBorder2 = Borders[1];
            RawDataSelection = ExtractFromFile(filename, CompressionRateSelection, SelectionBorder1, SelectionBorder2);

            /*Sel1 = new double[RawDataSelection.Length];
            Sel2 = new double[RawDataSelection.Length];
            Sel3 = new double[RawDataSelection.Length];
            Sel4 = new double[RawDataSelection.Length];
            Sel5 = new double[RawDataSelection.Length];*/

            Sel1 = CutArrayFromRows(RawDataSelection, 0);
            Sel2 = CutArrayFromRows(RawDataSelection, 1);
            Sel3 = CutArrayFromRows(RawDataSelection, 2);
            Sel4 = CutArrayFromRows(RawDataSelection, 3);
            Sel5 = CutArrayFromRows(RawDataSelection, 4);

            /*for (int i = 0; i < RawDataSelection.Length; i++)
            {
                Sel1[i] = CutValuesFromStringDat(RawDataSelection[i])[0];
                Sel2[i] = CutValuesFromStringDat(RawDataSelection[i])[1];
                Sel3[i] = CutValuesFromStringDat(RawDataSelection[i])[2];
                Sel4[i] = CutValuesFromStringDat(RawDataSelection[i])[3];
                Sel5[i] = CutValuesFromStringDat(RawDataSelection[i])[4];
            }
            */

            //присовение глобальным массивам новых значений
            x = Sel1;
            y = Sel2;
            n = x.Length;

            y2 = Sel3;
            y3 = Sel4;
            y4 = Sel5;

            //сглаживание массива иксов
            double[] x2 = new double[x.Length];
            x2 = SmoothArray(x);
            x = x2;

            //BuildGraph(RawDataSelection, false, true, true, true, true);
            BuildGraph2(true, true, true, true, true);
        }

        private void bSaveDB_Click(object sender, EventArgs e)
        {
            //найти макс id в базе
            int MaxID = 0;

            Boolean EntriesExist = false;
            /*string sql = "select exists (select max(id_entry) FROM Entry)";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                            if (Convert.ToBoolean(rdr["exists (select max(id_entry) FROM Entry)"]))
                                EntriesExist = true;
                    }
                }
            }*/

            string sql = "select count(id_entry) as Rows FROM Entry";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                            if (Convert.ToInt32(rdr["Rows"]) > 0)
                                EntriesExist = true;
                    }
                }
            }

            if (EntriesExist)
            {
                sql = "select max(id_entry) AS max FROM Entry";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            if (rdr.Read())
                                MaxID = Convert.ToInt32(rdr["max"]);
                        }
                    }
                }
            }
            else MaxID = 0;

            MaxID++;
            string CurrentDate = DateTime.Now.ToString("s", System.Globalization.CultureInfo.InvariantCulture);
            string Name = tbEntryName.Text;

            sql = @"insert into Entry
                Values(@id, @date, @name)";
            
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    cmd.Parameters.AddWithValue("@id", MaxID);
                    cmd.Parameters.AddWithValue("@date", CurrentDate);
                    cmd.Parameters.AddWithValue("@name", Name);
                    cmd.ExecuteNonQuery();
                }
            }

            if ((textBox1.Text != "") && (textBox28.Text == ""))
            {
                //перевод массивов в строки
                sInitialApp = ArrayToString(dInitialApp);
                sAFExtremums = ArrayToString(dAFE);
                sFExtremums = ArrayToString(dFE);
                sCoeffs = ArrayToString(dCoeffs);
                sIndexes = ArrayToString(dIndexes);
                sYsh = ArrayToString(Ysh);
                sYsh2 = ArrayToString(Ysh2);

                //ЗАПИСЬ В БАЗУ
                sql = @"INSERT INTO OrdinaryContraction
                Values (@id, @InitApp, @AFE, @FE, @Coeffs, @Indexes, @Df1, @Df2)";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@id", MaxID);
                        cmd.Parameters.AddWithValue("@InitApp", sInitialApp);
                        cmd.Parameters.AddWithValue("@AFE", sAFExtremums);
                        cmd.Parameters.AddWithValue("@FE", sFExtremums);
                        cmd.Parameters.AddWithValue("@Coeffs", sCoeffs);
                        cmd.Parameters.AddWithValue("@Indexes", sIndexes);
                        cmd.Parameters.AddWithValue("@Df1", sYsh);
                        cmd.Parameters.AddWithValue("@Df2", sYsh2);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            if ((textBox1.Text == "") && (textBox28.Text != ""))
            {
                //ЗАПИСЬ В БАЗУ
                sql = @"INSERT INTO ExtrasystolicContraction
                Values (@id, @Max1, @Max2)";

                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@id", MaxID);
                        cmd.Parameters.AddWithValue("@Max1", dESMax1);
                        cmd.Parameters.AddWithValue("@Max2", dESMax2);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            //залить соус в базу
            //перевод массивов в строки
            sX = ArrayToString(x);
            sY1 = ArrayToString(y);
            sY2 = ArrayToString(y2);
            sY3 = ArrayToString(y3);
            sY4 = ArrayToString(y4);

            //ЗАПИСЬ СОУСА В БАЗУ 
            sql = @"INSERT INTO Source
                    Values (@id, @x, @y1, @y2, @y3, @y4)";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    cmd.Parameters.AddWithValue("@id", MaxID);
                    cmd.Parameters.AddWithValue("@x", sX);
                    cmd.Parameters.AddWithValue("@y1", sY1);
                    cmd.Parameters.AddWithValue("@y2", sY2);
                    cmd.Parameters.AddWithValue("@y3", sY3);
                    cmd.Parameters.AddWithValue("@y4", sY4);
                    cmd.ExecuteNonQuery();
                }
            }

            MessageBox.Show("Измерение сохранено в базе данных!");
        }

        private void bOpenDB_Click(object sender, EventArgs e)
        {
            /*DBLoad();

            string sql = "Select * FROM OrdinaryContraction";
            SQLiteCommand com = new SQLiteCommand(sql, m_dbConnection);
            SQLiteDataReader r = com.ExecuteReader();

            string output = "";
            while (r.Read())
                output += Convert.ToString(r["id_entry"] + " " + r["InitialApproximation"] + " " + r["AFEXtremums"] + " " + r["FExtremums"] + " " + r["Coeffs"] + " " + r["Indexes"]) + "\n";

            DBClose();

            MessageBox.Show(output);*/

            tabControl1.SelectedIndex = 1;

            //найти число строк в базе
            int MaxID = 0;
            string sql = "select count(id_entry) AS Ids FROM Entry";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                            MaxID = Convert.ToInt32(rdr["Ids"]);
                    }
                }
            }

            //вывести всё в ДГВ из Entry
            dbview.RowCount = MaxID;
            dbview.ColumnCount = 2;

            int i = 0;
            sql = "select * from Entry";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            dbview[0, i].Value = Convert.ToString(rdr["id_entry"]);
                            dbview[1, i].Value = Convert.ToString(rdr["datemark"]).Replace('T', ' ').Replace('-', '.') + " " + Convert.ToString(rdr["name"]);
                            i++;
                        }
                    }
                }
            }
        }

        private void bDBLoad_Click(object sender, EventArgs e)
        {
            //выбранный ид искать сначала в ординари, потом - в экстрасистолик
            //найти выбранный ид в ординари
            bool InOrdinary = false;
            int SelectedId = Convert.ToInt32(dbview[0, dbview.CurrentCell.RowIndex].Value);

            string sql = "select exists (select * from OrdinaryContraction where id_entry = @SelectedId) as InOrd";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                            InOrdinary = Convert.ToBoolean(rdr["InOrd"]);
                    }
                }
            }

            if (InOrdinary)
            {
                sql = "select * from OrdinaryContraction where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            if (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                sInitialApp = Convert.ToString(rdr["InitialApproximation"]);
                                sAFExtremums = Convert.ToString(rdr["AFextremums"]);
                                sFExtremums = Convert.ToString(rdr["FExtremums"]);
                                sCoeffs = Convert.ToString(rdr["Coeffs"]);
                                sIndexes = Convert.ToString(rdr["Indexes"]);
                                sYsh = Convert.ToString(rdr["Derivative1"]);
                                sYsh2 = Convert.ToString(rdr["Derivative2"]);
                            }
                        }
                    }
                }

                //перевести все строковые переменные в массивы
                dInitialApp = StringToArray(sInitialApp);
                dAFE = StringToArray(sAFExtremums);
                dFE = StringToArray(sFExtremums);
                dCoeffs = StringToArray(sCoeffs);
                dIndexes = StringToArray(sIndexes);
                Ysh = StringToArray(sYsh);
                Ysh2 = StringToArray(sYsh2);
                //значения из массивов - в текстбоксы
                textBox1.Text = Convert.ToString(Math.Round(dCoeffs[0], 3));
                textBox2.Text = Convert.ToString(Math.Round(dCoeffs[1], 3));
                textBox3.Text = Convert.ToString(Math.Round(dCoeffs[2], 3));
                textBox4.Text = Convert.ToString(Math.Round(dCoeffs[3], 3));
                textBox5.Text = Convert.ToString(Math.Round(dCoeffs[4], 4));
                textBox6.Text = Convert.ToString(Math.Round(dCoeffs[5], 5));

                textBox15.Text = Convert.ToString(Math.Round(dAFE[0], 3));
                textBox16.Text = Convert.ToString(Math.Round(dAFE[1], 3));
                textBox17.Text = Convert.ToString(Math.Round(dAFE[2], 3));
                textBox18.Text = Convert.ToString(Math.Round(dAFE[3], 3));

                textBox20.Text = Convert.ToString(Math.Round(dFE[0], 3));
                textBox21.Text = Convert.ToString(Math.Round(dFE[1], 3));
                textBox22.Text = Convert.ToString(Math.Round(dFE[2], 3));
                textBox23.Text = Convert.ToString(Math.Round(dFE[3], 3));

                textBox24.Text = Convert.ToString(Math.Round(dIndexes[0], 3));
                textBox25.Text = Convert.ToString(Math.Round(dIndexes[1], 3));
                textBox26.Text = Convert.ToString(Math.Round(dIndexes[2], 3));
                textBox27.Text = Convert.ToString(Math.Round(dIndexes[3], 3));

                textBox7.Text = Convert.ToString(Math.Round(dInitialApp[0], 3));
                textBox8.Text = Convert.ToString(Math.Round(dInitialApp[1], 3));
                textBox9.Text = Convert.ToString(Math.Round(dInitialApp[2], 3));
                textBox10.Text = Convert.ToString(Math.Round(dInitialApp[3], 3));
                textBox11.Text = Convert.ToString(Math.Round(dInitialApp[4], 3));
                textBox12.Text = Convert.ToString(Math.Round(dInitialApp[5], 3));

                //активировать главную вкладку
                tabControl1.SelectedIndex = 0;
            }
            else
            {
                sql = "select * from ExtrasystolicContraction where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            if (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                sESMax1 = Convert.ToString(rdr["Max1"]);
                                sESMax2 = Convert.ToString(rdr["Max2"]);
                            }
                        }
                    }
                }
            }

            //загрузить из БД соус
            sql = "select * from Source where id_entry = @SelectedId";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            //перевести всё из базы в строковые переменные
                            sX = Convert.ToString(rdr["x"]);
                            sY1 = Convert.ToString(rdr["y1"]);
                            sY2 = Convert.ToString(rdr["y2"]);
                            sY3 = Convert.ToString(rdr["y3"]);
                            sY4 = Convert.ToString(rdr["y4"]);
                        }
                    }
                }
            }

            //перевести соус в массивы
            x = StringToArray(sX);
            y = StringToArray(sY1);
            y2 = StringToArray(sY2);
            y3 = StringToArray(sY3);
            y4 = StringToArray(sY4);

            //построить графики
            chart1.Series[1].Points.DataBindXY(x, y);
            chart2.Series[0].Points.DataBindXY(x, Ysh);
            chart2.Series[1].Points.DataBindXY(x, Ysh2);
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1)
            {
                bOpenDB_Click(sender, e);
            }
        }

        private void btnCalcExs_Click(object sender, EventArgs e)
        {
            chart1.Visible = true;
            chart2.Visible = false;
            chart1.Size = new Size(915, 600);

            double[] max = FindMaximums(y);
            dESMax1 = max[0];
            dESMax2 = max[1];

            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";

            textBox15.Text = "";
            textBox16.Text = "";
            textBox17.Text = "";
            textBox18.Text = "";

            textBox20.Text = "";
            textBox21.Text = "";
            textBox22.Text = "";
            textBox23.Text = "";

            textBox24.Text = "";
            textBox25.Text = "";
            textBox26.Text = "";
            textBox27.Text = "";

            textBox28.Text = Convert.ToString(max[0]);
            textBox29.Text = Convert.ToString(max[1]);
        }

        //метод поиска максимумов в экстрасистолическом сокращении
        private double[] FindMaximums(double[] Arr)
        {
            double[] max = new double[2];
            int i = 1;
            int index = 0;
            while ((index <= 1) || (i < Arr.Length - 1))
            {
                if ((Arr[i] > Arr[i + 1]) && (Arr[i] > Arr[i - 1]))
                {
                    max[index] = Arr[i];
                    index++;
                }
                i++;
            }

            return max;
        }

        private double[] SmoothArray(double[] Arr)
        {
            double[] Arr2 = Arr;

            int i = 0;
            int j = 0;
            double CurrentValue = 0;
            double NextValue = 0;
            int CurrentIndex = 0;
            int NextIndex = 0;
            
            while (CurrentIndex < Arr.Length - 1)
            {
                CurrentValue = Arr[CurrentIndex];
                while ((Arr[NextIndex] == CurrentValue) && (NextIndex < Arr.Length - 1))
                    NextIndex++;

                NextValue = Arr[NextIndex];

                //сгладить фрагмент массива
                for (i = CurrentIndex; i < NextIndex; i++)
                {
                    Arr2[i] = CurrentValue + (i - CurrentIndex) * (NextValue - CurrentValue) / (NextIndex - CurrentIndex);
                }

                CurrentIndex = NextIndex;
            }

            return Arr2;
        }

        private void bDBDelete_Click(object sender, EventArgs e)
        {
            int SelectedId = 0;
            SelectedId = Convert.ToInt32(dbview[0, dbview.CurrentCell.RowIndex].Value);
            bool Exists = false;

            //1 - УДАЛЕНИЕ ИЗ ENTRY
            string sql = "SELECT exists (SELECT * FROM Entry WHERE id_entry = @id_e) as Checked";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    cmd.Parameters.AddWithValue("@id_e", SelectedId);
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                            Exists = Convert.ToBoolean(rdr["Checked"]);
                    }
                }
            }

            if (Exists)
            {
                sql = "DELETE FROM Entry WHERE id_entry = @id";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@id", SelectedId);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            //2 - УДАЛЕНИЕ ИЗ OrdinaryContraction
            sql = "SELECT EXISTS (SELECT * FROM OrdinaryContraction WHERE id_entry = @id) as Checked";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    cmd.Parameters.AddWithValue("@id", SelectedId);
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                            Exists = Convert.ToBoolean(rdr["Checked"]);
                    }
                }
            }

            if (Exists)
            {
                sql = "DELETE FROM OrdinaryContraction WHERE id_entry = @id";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@id", SelectedId);
                        cmd.ExecuteNonQuery();
                    }
                }

            }

            //3 - УДАЛЕНИЕ ИЗ ExtrasystolicContraction
            sql = "SELECT EXISTS (SELECT * FROM ExtrasystolicContraction WHERE id_entry = @id) as Checked";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    cmd.Parameters.AddWithValue("@id", SelectedId);
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                            Exists = Convert.ToBoolean(rdr["Checked"]);
                    }
                }
            }

            if (Exists)
            {
                sql = "DELETE FROM ExtrasystolicContraction WHERE id_entry = @id";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@id", SelectedId);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            //4 - УДАЛЕНИЕ ИЗ SOURCE
            sql = "SELECT EXISTS (SELECT * FROM Source WHERE id_entry = @id) as Checked";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    cmd.Parameters.AddWithValue("@id", SelectedId);
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                            Exists = Convert.ToBoolean(rdr["Checked"]);
                    }
                }
            }

            if (Exists)
            {
                sql = "DELETE FROM Source WHERE id_entry = @id";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@id", SelectedId);
                        cmd.ExecuteNonQuery();
                    }
                }
            }

            MessageBox.Show("Данные удалены!");
            bOpenDB_Click(sender, e);
        }

        #region Nuy Methods
        public string[] ExtractFromFile(string filename, double compressionrate)
        {
            string[] Extraction;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //поиск первой пустой строки
            int rightindex = 1;
            Excel.Range RightCell;
            RightCell = (Excel.Range)xlWorkSheet.Cells[1, 1];

            //найти границу пустоты с точностью до 5000
            int SearchStep = 5000;
            while (RightCell.Value2 != null)
            {
                rightindex += SearchStep;
                RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
            }

            //найти точную границу методом половинного деления
            int leftindex = rightindex - SearchStep;
            Excel.Range LeftCell;
            LeftCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];
            RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
            while (SearchStep > 0)
            {
                LeftCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];
                RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
                if ((LeftCell.Value2 != null) && (RightCell.Value2 == null))
                {
                    SearchStep = SearchStep / 2;
                    leftindex += SearchStep;
                }
                if ((LeftCell.Value2 == null) && (RightCell.Value2 == null))
                    leftindex -= SearchStep;
            }

            xlRange = (Excel.Range)xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[1, 1], (Excel.Range)xlWorkSheet.Cells[leftindex + 1, 1]);

            int j = 0;

            //извлечение элементов из рейнджа в массив
            Extraction = new string[xlRange.Count / Convert.ToInt32(compressionrate) + 1];
            for (int i = 0; i < xlRange.Count; i += Convert.ToInt32(compressionrate))
            {
                Extraction[j] = xlRange[i + 1, 1].Value2;
                j++;
            }

            xlWorkBook.Close();
            xlApp.Quit();

            return Extraction;
        }

        public string[] ExtractFromFile(string filename, double cRate, double left, double right)
        {
            string[] Extraction;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlRange = (Excel.Range)xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[left, 1], (Excel.Range)xlWorkSheet.Cells[right, 1]);

            //извлечение элементов из рейнджа в массив
            Extraction = new string[xlRange.Count / Convert.ToInt32(cRate)];
            for (int i = 0; i < xlRange.Count; i += Convert.ToInt32(cRate))
            {
                Extraction[i / Convert.ToInt32(cRate)] = xlRange[i, 1].Value2;
            }

            return Extraction;
        }

        private string ArrayToString(double[] Arr)
        {
            string res = "";
            for (int i = 0; i < Arr.Length; i++)
            {
                if (!(res == ""))
                    res += "/" + Convert.ToString(Arr[i]);
                else res = Convert.ToString(Arr[i]);
            }

            return res;
        }

        private double[] StringToArray(string str)
        {
            //считаем разделители
            int i = 0;
            int SeparatorCount = 0;
            for (i = 0; i < str.Length; i++)
                if (str[i] == '/')
                    SeparatorCount++;

            int NextSpr = 0;
            int PrevSpr = -1;
            string tempStr = "";

            double[] Arr = new double[SeparatorCount + 1];

            i = 0;
            while (i < SeparatorCount)
            {
                NextSpr = str.IndexOf('/', PrevSpr + 1);

                //Arr[i] = Convert.ToDouble("14.7");
                tempStr = str.Substring(PrevSpr + 1, NextSpr - PrevSpr - 1).Replace('.', ',');
                Arr[i] = Convert.ToDouble(tempStr);

                PrevSpr = NextSpr;
                i++;
            }

            tempStr = str.Substring(PrevSpr + 1, str.Length - PrevSpr - 1).Replace('.', ',');
            Arr[Arr.Length - 1] = Convert.ToDouble(tempStr);

            return Arr;
        }

        private double[] CutArrayFromRows(string[] data, int index)
        {
            double[] Col = new double[data.Length];

            for (int i = 0; i < data.Length; i++)
                Col[i] = Convert.ToDouble(CutValuesFromStringDat(data[i])[index]);

            return Col;
        }

        private void BuildGraph(string[] Data, bool c1, bool c2, bool c3, bool c4, bool c5)
        {
            /*Col1 = new double[Data.Length];
            Col2 = new double[Data.Length];
            Col3 = new double[Data.Length];
            Col4 = new double[Data.Length];
            Col5 = new double[Data.Length];

            for (int i = 0; i < Col1.Length; i++)
            {
                Col1[i] = CutValuesFromStringDat(Data[i])[0];
                Col2[i] = CutValuesFromStringDat(Data[i])[1];
                Col3[i] = CutValuesFromStringDat(Data[i])[2];
                Col4[i] = CutValuesFromStringDat(Data[i])[3];
                Col5[i] = CutValuesFromStringDat(Data[i])[4];
            }*/

            Col1 = CutArrayFromRows(RawDataFull, 0);
            Col2 = CutArrayFromRows(RawDataFull, 1);
            Col3 = CutArrayFromRows(RawDataFull, 2);
            Col4 = CutArrayFromRows(RawDataFull, 3);
            Col5 = CutArrayFromRows(RawDataFull, 4);

            /*if (c1) chart1.Series[0].Points.DataBindY(Col1);
            if (c2) chart1.Series[1].Points.DataBindY(Col2);
            if (c3) chart1.Series[2].Points.DataBindY(Col3);
            if (c4) chart1.Series[3].Points.DataBindY(Col4);
            if (c5) chart1.Series[4].Points.DataBindY(Col5);*/

            chart1.Series[1].Points.DataBindXY(Col1, Col2);
            chart1.Series[2].Points.DataBindXY(Col1, Col3);
            chart1.Series[3].Points.DataBindXY(Col1, Col4);
            chart1.Series[4].Points.DataBindXY(Col1, Col5);

            chart1.Series[1].BorderWidth = 2;
            chart1.Series[2].BorderWidth = 2;
            chart1.Series[3].BorderWidth = 2;
            chart1.Series[4].BorderWidth = 2;
            chart2.Series[0].BorderWidth = 2;
            chart2.Series[1].BorderWidth = 2;

        }

        private void BuildGraph2(bool Selection, bool c2, bool c3, bool c4, bool c5)
        {
            if (Selection)
            {
                if (c2) chart1.Series[1].Points.DataBindXY(Sel1, Sel2);
                if (c3) chart1.Series[2].Points.DataBindXY(Sel1, Sel3);
                if (c4) chart1.Series[3].Points.DataBindXY(Sel1, Sel4);
                if (c5) chart1.Series[4].Points.DataBindXY(Sel1, Sel5);
            }
            else
            {
                if (c2) chart1.Series[1].Points.DataBindXY(Col1, Col2);
                if (c3) chart1.Series[2].Points.DataBindXY(Col1, Col3);
                if (c4) chart1.Series[3].Points.DataBindXY(Col1, Col4);
                if (c5) chart1.Series[4].Points.DataBindXY(Col1, Col5);
            }

            chart1.Series[1].BorderWidth = 2;
            chart1.Series[2].BorderWidth = 2;
            chart1.Series[3].BorderWidth = 2;
            chart1.Series[4].BorderWidth = 2;
        }

        //вывести на график границы выборки
        private void ShowSelectionBorder(double left, double right)
        {
            System.Windows.Forms.DataVisualization.Charting.ChartArea CA;
            System.Windows.Forms.DataVisualization.Charting.Series S1;
            System.Windows.Forms.DataVisualization.Charting.VerticalLineAnnotation VA1;
            System.Windows.Forms.DataVisualization.Charting.VerticalLineAnnotation VA2;
            VA1 = new System.Windows.Forms.DataVisualization.Charting.VerticalLineAnnotation();
            VA2 = new System.Windows.Forms.DataVisualization.Charting.VerticalLineAnnotation();

            if (chart1.Annotations.Count == 0)
            {
                //chart1.Annotations.Remove(VA1);
                //chart1.Annotations.Remove(VA2);

                CA = chart1.ChartAreas[0];  // pick the right ChartArea..
                S1 = chart1.Series["Borders"];      // ..and Series!

                // the vertical line
                VA1.AxisX = CA.AxisX;
                VA1.AllowMoving = true;
                VA1.IsInfinitive = true;
                VA1.ClipToChartArea = CA.Name;
                VA1.Name = "myLine1";
                VA1.LineColor = Color.Red;
                VA1.LineWidth = 2;         // use your numbers!
                VA1.X = left;

                VA2.AxisX = CA.AxisX;
                VA2.AllowMoving = true;
                VA2.IsInfinitive = true;
                VA2.ClipToChartArea = CA.Name;
                VA2.Name = "myLine2";
                VA2.LineColor = Color.Red;
                VA2.LineWidth = 2;         // use your numbers!
                VA2.X = right;

                chart1.Annotations.Add(VA1);
                chart1.Annotations.Add(VA2);
            }
            else
            {
                chart1.Annotations.Remove(VA1);
                chart1.Annotations.Remove(VA2);
            }
        }

        private int[] FindSelectionBorders(string filename, double left, double right)
        {
            int[] Borders = new int[2];
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //найти LEFT и RIGHT по имеющимся таймингам
            int SearchStep = 5000;
            int leftindex = 1; int rightindex = leftindex;
            Excel.Range CurrentCell;
            CurrentCell = (Excel.Range)xlWorkSheet.Cells[1 + SearchStep, 1];
            Excel.Range PrevCell;
            PrevCell = (Excel.Range)xlWorkSheet.Cells[1, 1];
            double CurValue = CutValuesFromStringDat(CurrentCell.Value2)[0];
            double PrevValue = CutValuesFromStringDat(PrevCell.Value2)[0];

            //найти примерные границы левого края
            while ((((CurValue - left) * (PrevValue - left)) > 0) || !(string.IsNullOrEmpty(CurrentCell.Value2)))
            {
                leftindex += SearchStep;
                CurrentCell = (Excel.Range)xlWorkSheet.Cells[leftindex + SearchStep, 1];
                PrevCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];

                if (string.IsNullOrEmpty(CurrentCell.Value2))
                    break;
                else
                {
                    CurValue = CutValuesFromStringDat(CurrentCell.Value2)[0];
                    PrevValue = CutValuesFromStringDat(PrevCell.Value2)[0];
                }
            }

            rightindex = leftindex + SearchStep;

            //найти точную границу методом половинного деления
            Excel.Range LeftCell;
            Excel.Range RightCell;
            LeftCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];
            RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
            while (SearchStep > 0)
            {
                LeftCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];
                RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
                if ((LeftCell.Value2 != null) && (RightCell.Value2 == null))
                {
                    SearchStep = SearchStep / 2;
                    rightindex -= SearchStep;
                }
                if ((LeftCell.Value2 != null) && (RightCell.Value2 != null))
                    rightindex += SearchStep;
            }

            rightindex--;
            int RightNotEmpty = rightindex;

            //найденная выше граница является ГРАНИЦЕЙ ЗНАЧЕНИЙ, ГДЕ ВООБЩЕ В ПРИНЦИПЕ ЕСТЬ КАКИЕ-ТО ЗНАЧЕНИЯ
            //теперь необходимо найти левую границу ИСКОМОГО ДИАПАЗОНА
            //LEFT = 5000, RIGHT = 9570

            SearchStep = RightNotEmpty - leftindex;
            Excel.Range MidCell;
            MidCell = (Excel.Range)xlWorkSheet.Cells[leftindex];

            //поиск левой точной границы
            while (Math.Abs(CutValuesFromStringDat(LeftCell.Value2)[0] - left) > 0.1)
            {
                double cval = CutValuesFromStringDat(LeftCell.Value2)[0];
                SearchStep = Convert.ToInt32(80 * (left - cval));
                leftindex += SearchStep;
                LeftCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];
            }

            RightCell = (Excel.Range)xlWorkSheet.Cells[RightNotEmpty, 1];
            //поиск правой точной границы
            while (Math.Abs(CutValuesFromStringDat(RightCell.Value2)[0] - right) > 0.1)
            {
                double cval = CutValuesFromStringDat(RightCell.Value2)[0];
                SearchStep = Convert.ToInt32(80 * (cval - right));
                rightindex -= SearchStep;
                RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
            }

            Borders[0] = leftindex;
            Borders[1] = rightindex;
            return Borders;
        }

        private double[] CutValuesFromStringDat(string s)
        {
            double[] Output = new double[5];

            int CutCount = 0;
            int PrevComma = -1;
            int NextComma = -1;
            string tempStr = "";

            while (CutCount < 11)
            {
                NextComma = s.IndexOf(',', PrevComma + 1);

                if ((CutCount == 0) || (CutCount == 1) || (CutCount == 4) || (CutCount == 7) || (CutCount == 10))
                {
                    tempStr = s.Substring(PrevComma + 1, NextComma - PrevComma - 1).Replace('.', ',');

                    switch (CutCount)
                    {
                        case 0: Output[0] = Convert.ToDouble(tempStr); break;
                        case 1: Output[1] = Convert.ToDouble(tempStr); break;
                        case 4: Output[2] = Convert.ToDouble(tempStr); break;
                        case 7: Output[3] = Convert.ToDouble(tempStr); break;
                        case 10: Output[4] = Convert.ToDouble(tempStr); break;
                    }

                    PrevComma = NextComma;
                    CutCount++;
                }
                else
                {
                    PrevComma = NextComma;
                    CutCount++;
                }
            }
            return Output;
        }


        #endregion
    }
}