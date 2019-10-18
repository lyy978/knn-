using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace knnForms
{
    public partial class Form1 : Form
    {
        int temp = 0;
        int len;

        List<int> trainX1 = new List<int>();
        List<int> trainX2 = new List<int>();
        List<int> trainX3 = new List<int>();
        List<double> trainX4 = new List<double>();
        List<double> trainX5 = new List<double>();
        List<double> trainY2 = new List<double>();
        List<double> trainY1 = new List<double>();
        List<double> trainY3 = new List<double>();


        public Form1()
        {
            InitializeComponent();
            button1.Enabled = false;
            ExcelEdit ed = new ExcelEdit();
            ed.Open("E:\\C#代码\\knnprogram\\result1.xlsx");      //open a excel file
            Excel.Worksheet worksheet = (Excel.Worksheet)ed.GetSheet("sheet1"); //choose the sheet
            ed.getTestnum(ed);
            len = worksheet.UsedRange.Rows.Count;
            trainX1 = ed.trainX1;
            trainX2 = ed.trainX2;
            trainX3 = ed.trainX3;
            trainX4 = ed.trainX4;
            trainX5 = ed.trainX5;
            trainY1 = ed.trainY1;
            trainY2 = ed.trainY2;
            trainY3 = ed.trainY3;
            button1.Enabled = true;


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)   //输入框1 输入
        {
          
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
     
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
        
            string t1;
            double testx1;
            double testx2;
            double testx3;
            double testx4;
            double testx5;
            double testy1=0;
            double testy2=0;
            double testy3=0;
            double m1 = 0;
            double m2 = 0;
            double m3 = 0;
            double b;
            double c;
            double d;
            double f;

            t1 = textBox1.Text;
            testx1 = Convert.ToDouble(t1);         //   Get the value of the textbox1

            t1 = textBox2.Text;
            testx2 = Convert.ToDouble(t1);         //   Get the value of the textbox1

            t1 = textBox3.Text;
            testx3 = Convert.ToDouble(t1);         //   Get the value of the textbox1

            t1 = textBox4.Text;
            testx4 = Convert.ToDouble(t1);         //   Get the value of the textbox1

            t1 = textBox5.Text;
            testx5 = Convert.ToDouble(t1);         //   Get the value of the textbox1

            t1 = textBox9.Text;
            b = Convert.ToDouble(t1);         //   Get the value of the textbox1

            t1 = textBox10.Text;
            c = Convert.ToDouble(t1);         //   Get the value of the textbox1

            t1 = textBox11.Text;
            d = Convert.ToDouble(t1);         //   Get the value of the textbox1

            t1 = textBox12.Text;
            f = Convert.ToDouble(t1);         //   Get the value of the textbox1



        //    MessageBox.Show(""+testx1+testx2+testx3+testx4+testx5);
            Knn knn = new Knn();
            knn.distance( ref trainX1, ref trainX2, ref trainX3, ref trainX4, ref trainX5, ref trainY2, ref trainY1, ref trainY3, testx1,  testx2, testx3,  testx4, testx5, ref testy1 ,ref  testy2, ref testy3);
            m1=knn.calculateM(testy1,testx1,testx2,b,c,d,f);
            m2 = knn.calculateM(testy2, testx1, testx2, b, c, d, f);
            m3 = knn.calculateM(testy3, testx1, testx2, b, c, d, f);
            textBox8.Text = Convert.ToString(m1);
            textBox7.Text = Convert.ToString(m2);
            textBox6.Text = Convert.ToString(m3);

        }
        

        private void button2_Click(object sender, EventArgs e)         //get the value of training set
        {
            // Console.WriteLine("Running...");
            ExcelEdit ed = new ExcelEdit();
            ed.Open("E:\\C#代码\\knnprogram\\result1.xlsx");      //open a excel file
            Excel.Worksheet worksheet = (Excel.Worksheet)ed.GetSheet("sheet1"); //choose the sheet
            ed.getTestnum(ed);                            //get training set data
            MessageBox.Show("ok!");
            temp = 1;

        }

        private void textBox8_TextChanged(object sender, EventArgs e)  //t1
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)  //2
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)   //t3
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)  //b
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e) //c
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)  //d
        {

        }

        private void textBox12_TextChanged(object sender, EventArgs e)  //f
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }
    }

    public class Knn
    {

        public void distance( ref List<int> trainX1, ref List<int> trainX2, ref List<int> trainX3, ref List<double> trainX4, ref List<double> trainX5, ref List<double> trainY2, ref List<double> trainY1, ref List<double> trainY3, double testX1, double testX2, double testX3, double testX4, double testX5, ref double testY1, ref double testY2,ref double testY3)
        {
            List<double> dis = new List<double>();
            double temp = 0;
            double temp1 = 0;
            double temp2 = 0;
            double temp3 = 0;
            double temp4 = 0;
            double temp5 = 0;
            int i = 0;                              //length of training set
            int j = 0;                              //length of test data
            int k = 0;                              //the index of the shortest distance
            double min;
            //  int count = 0;
                  
            dis.Clear();
            //   Console.WriteLine("discount:{0}", dis.Count); 
            //   count++;
            for (i = 0; i < trainX1.Count; i++)
            {
                temp1 = (testX1 - trainX1[i]) * (testX1 - trainX1[i]);
                temp2 = (testX2 - trainX2[i]) * (testX2 - trainX2[i]);
                temp3 = (testX3 - trainX3[i]) * (testX3 - trainX3[i]);
                temp4 = (testX4 - trainX4[i]) * (testX4 - trainX4[i]);
                temp5 = (testX5 - trainX5[i]) * (testX5 - trainX5[i]);
                temp = temp1 + temp2 + temp3 + temp4 + temp5;
                dis.Add(temp);
            }
            min = (double)dis.Min<double>();
            for (i = 1; i < trainX1.Count; i++)
            {
                if (dis[i] == min)
                {
                    k = i;
                    break;
                }
            }
            testY1=trainY1[k];
            testY2=trainY2[k];
            testY3=trainY3[k];
            dis.Clear();
        }

        public double calculateM(double testY, double testx1,double testx2,double b,double c ,double d,double f )  //T:testy  j:testx2  i:testx1
        {
            double m;

            m = ((0.5 * testY + testx2) * Math.PI * testx1 / (1 / b - c / d)) * ((0.5 * testY + testx2) * Math.PI * testx1 / (1 / b - c / d)) * c / (0.21 * Math.PI * testx1 * f);

            return m;
        }

    }

}
