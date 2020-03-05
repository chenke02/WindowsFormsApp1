using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 业务量情况和性能情况
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form2_Load(object sender, EventArgs e)
        {
            DataTable dt1 = new DataTable();  //实例化数据集
            Form1.DB d = new Form1.dataset();
            //调用空口流量求和项
            DataTable w1 = d.Num("空口流量", "2019/12/1"); 
            DataTable w2 = d.Num("空口流量", "2019/12/2");
            DataTable w3 = d.Num("空口流量", "2019/12/3");
            DataTable w4 = d.Num("空口流量", "2019/12/4");
            DataTable w5 = d.Num("空口流量", "2019/12/5");
            DataTable w6 = d.Num("空口流量", "2019/12/6");
            DataTable w7 = d.Num("空口流量", "2019/12/7");
            //求和项转化整型
            int n1 = Convert.ToInt32(w1.Rows[0][0]);
            int n2 = Convert.ToInt32(w2.Rows[0][0]);
            int n3 = Convert.ToInt32(w3.Rows[0][0]);
            int n4 = Convert.ToInt32(w4.Rows[0][0]);
            int n5 = Convert.ToInt32(w5.Rows[0][0]);
            int n6 = Convert.ToInt32(w6.Rows[0][0]);
            int n7 = Convert.ToInt32(w7.Rows[0][0]);
            //调用RRC连接数最大项求和项
            DataTable r1 = d.Num("RRC连接最大数", "2019/12/1");
            DataTable r2 = d.Num("RRC连接最大数", "2019/12/2");
            DataTable r3 = d.Num("RRC连接最大数", "2019/12/3");
            DataTable r4 = d.Num("RRC连接最大数", "2019/12/4");
            DataTable r5 = d.Num("RRC连接最大数", "2019/12/5");
            DataTable r6 = d.Num("RRC连接最大数", "2019/12/6");
            DataTable r7 = d.Num("RRC连接最大数", "2019/12/7");
            //求和项转化整型
            int p1 = Convert.ToInt32(r1.Rows[0][0]);
            int p2 = Convert.ToInt32(r2.Rows[0][0]);
            int p3 = Convert.ToInt32(r3.Rows[0][0]);
            int p4 = Convert.ToInt32(r4.Rows[0][0]);
            int p5 = Convert.ToInt32(r5.Rows[0][0]);
            int p6 = Convert.ToInt32(r6.Rows[0][0]);
            int p7 = Convert.ToInt32(r7.Rows[0][0]);   
            List<string> xData = new List<string>() { @"2019/12/1", @"2019/12/2", @"2019/12/3", @"2019/12/4", @"2019/12/5", @"2019/12/6", @"2019/12/7" };//横轴
            List<long> yData = new List<long>() { n1, n2, n3, n4, n5, n6, n7 }; //第一纵轴
            List<long> yData2 = new List<long>() { p1, p2, p3, p4, p5, p6, p7 }; //第二纵轴
            dt1.Columns.Add("X", Type.GetType("System.String")); //添加X项
            dt1.Columns.Add("Y1", Type.GetType("System.String")); //添加Y项
            dt1.Columns.Add("Y2", Type.GetType("System.String")); //添加Y2项
            DataRow newRow; //实例化空数据集
            //空数据集添加X,Y1和Y2数据
            for (int i = 0; i < 7; i++)
            {
                newRow = dt1.NewRow();
                newRow["X"] = xData[i];
                newRow["Y1"] = yData[i];
                newRow["Y2"] = yData2[i] * (850000 - 740000) / 2500000 + 740000;
                dt1.Rows.Add(newRow);
            }
            chart1.DataSource = dt1;   //绑定数据
            chart1.Series[0].XValueMember = "X";  //X绑定到横轴
            chart1.Series[0].YValueMembers = "Y1"; //Y1绑定到纵轴
            chart1.Series[1].XValueMember = "X";  //X绑定到横轴
            chart1.Series[1].YValueMembers = "Y2"; //Y2绑定到纵轴
            chart1.DataBind();   
            //求各项平均值
            DataTable q1 = d.AVG("无线接通率");
            DataTable q2 = d.AVG("无线掉线率");
            DataTable q3 = d.AVG("切换成功率");
            DataTable q4 = d.AVG("MR覆盖率");
            //平均值转化浮点数
            float m1 = (float)Convert.ToSingle(q1.Rows[0][0].ToString())* 100;
            float m2 = (float)Convert.ToSingle(q2.Rows[0][0].ToString()) * 100;
            float m3 = (float)Convert.ToSingle(q3.Rows[0][0].ToString()) * 100;
            float m4 = (float)Convert.ToSingle(q4.Rows[0][0].ToString()) * 100;
            int t1 = (int)m1;
            int t2 = (int)m2;
            int t3 = (int)m3;
            int t4 = (int)m4;
            //各项指标列表
            List<int> y2Data1 = new List<int>() { t1, t1, t1, t1, t1, t1, t1 };
            List<int> y2Data2 = new List<int>() { t2, t2, t2, t2, t2, t2, t2 };
            List<int> y2Data3 = new List<int>() { t3, t3, t3, t3, t3, t3, t3 };
            List<int> y2Data4 = new List<int>() { t4, t4, t4, t4, t4, t4, t4 };
            chart1.Series[0].Color = Color.Blue;  //线条颜色
            chart1.Series[0].BorderWidth = 5;   //线条粗细
            chart1.ChartAreas[0].AxisY.Maximum = 850000;//设定y轴的最大值
            chart1.ChartAreas[0].AxisY.Minimum = 740000;//设定y轴的最大值
            chart1.ChartAreas[0].AxisY.Interval = 10000;//设定y轴间隔
            chart1.ChartAreas[0].AxisY2.Maximum = 2500000;//设定y轴的最大值
            chart1.ChartAreas[0].AxisY2.Minimum = 0;//设定y轴的最大值
            chart1.ChartAreas[0].AxisY2.Interval = 500000;//设定y轴间隔
            for (int i = 0; i < 4; i++)
            {
                chart2.Series[i].BorderWidth = 5;  //线条粗细
            }
            chart2.ChartAreas[0].AxisY.Maximum = 100;//设定y轴的最大值
            chart2.ChartAreas[0].AxisY.Minimum = 0;//设定y轴的最大值
            chart2.ChartAreas[0].AxisY.Interval = 20;//设定y轴间隔
            chart2.Series[0].Points.DataBindXY(xData, y2Data1);
            chart2.Series[1].Points.DataBindXY(xData, y2Data2);
            chart2.Series[2].Points.DataBindXY(xData, y2Data3);
            chart2.Series[3].Points.DataBindXY(xData, y2Data4);

        }
    }
}
