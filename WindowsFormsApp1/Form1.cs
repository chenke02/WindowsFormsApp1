using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
        DataTable ds;
        public Form1()
        {


            InitializeComponent();


        }

        /// <summary>
        /// 导入数据功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 导入数据ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string fileName = "";   //文件名称


            //选中导入文件
            try
            {
                openFileDialog1.ShowDialog();  //打开文件列表
                fileName = openFileDialog1.FileName.ToString();   //获取CSV文件名称
            }
            catch (Exception ee)
            {
                MessageBox.Show("打开文件错误!" + ee.Message.ToString());  //提示文件出错
            }

            //导入文件存放到数据集
            try
            {
                if (fileName != "")
                {
                    DB d = new dataset();  //实例化数据集
                    ds = d.ds(fileName);//导入文件存放到数据集
                    this.dataGridView1.DataSource = ds;//从csv导入数据到表格中
                    //SQLsave(ds);//保存到数据库
                }
            }

            catch (Exception)
            {
                MessageBox.Show("CSV未导入或打开");  //提示CSV文件导入未成功
            }
        }


        private void 默认日报分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();

        }

        private void 最大无线接通率ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ds = GetConnectTOP20();
            this.dataGridView1.DataSource = GetConnectTOP20();
        }

        private void 无线掉线率ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ds = GetDisTOP20();
            this.dataGridView1.DataSource = GetDisTOP20();
        }

        private void 切换成功率ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ds = GetSuccessTOP20();
            this.dataGridView1.DataSource = GetSuccessTOP20();
        }

        private void 空口流量ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ds = GetFlowTOP20();
            this.dataGridView1.DataSource = GetFlowTOP20();
        }

        private void RRC最大数ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ds = GetRRCTOP20();
            this.dataGridView1.DataSource = GetRRCTOP20();
        }

       


        /// <summary>
        /// 无线接通率
        /// </summary>
        /// <returns>方法返回DataTable对象</returns>
        private DataTable GetConnectTOP20()
        { 
            string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
            string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                @"SELECT TOP 20 * from ShiYan ORDER BY 9 DESC");
            SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                P_Str_SqlStr, P_Str_ConnectionStr);
            DataTable P_dt = new DataTable();//创建数据表
            P_SqlDataAdapter.Fill(P_dt);//填充数据表
            return P_dt;//返回数据表
        }


        /// <summary>
        /// 掉线率
        /// </summary>
        /// <returns></returns>
        private DataTable GetDisTOP20()
        {
            string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
            string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                @"SELECT TOP 20 * from ShiYan ORDER BY 11 DESC");
            SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                P_Str_SqlStr, P_Str_ConnectionStr);
            DataTable P_dt = new DataTable();//创建数据表
            P_SqlDataAdapter.Fill(P_dt);//填充数据表
            return P_dt;//返回数据表
        }


        /// <summary>
        /// 切换成功率
        /// </summary>
        /// <returns></returns>
        private DataTable GetSuccessTOP20()
        {
            string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
            string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                @"SELECT TOP 20 * from ShiYan ORDER BY 10 DESC");
            SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                P_Str_SqlStr, P_Str_ConnectionStr);
            DataTable P_dt = new DataTable();//创建数据表
            P_SqlDataAdapter.Fill(P_dt);//填充数据表
            return P_dt;//返回数据表
        }


        /// <summary>
        /// 空口流量
        /// </summary>
        /// <returns></returns>
        private DataTable GetFlowTOP20()
        {
            string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
            string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                @"SELECT TOP 20 * from ShiYan ORDER BY 13 DESC");
            SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                P_Str_SqlStr, P_Str_ConnectionStr);
            DataTable P_dt = new DataTable();//创建数据表
            P_SqlDataAdapter.Fill(P_dt);//填充数据表
            return P_dt;//返回数据表
        }


        /// <summary>
        /// RRC最大数
        /// </summary>
        /// <returns></returns>
        private DataTable GetRRCTOP20()
        {
            string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
            string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                @"SELECT TOP 20 * from ShiYan ORDER BY 14 DESC");
            SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                P_Str_SqlStr, P_Str_ConnectionStr);
            DataTable P_dt = new DataTable();//创建数据表
            P_SqlDataAdapter.Fill(P_dt);//填充数据表
            return P_dt;//返回数据表
        }


        /// <summary>
        /// 保存到数据库
        /// </summary>
        public static void SQLsave(DataTable dt)
        {
            //数据库变量
            string con = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True";
            //初始化连接数据库实例
            SqlConnection conn = new SqlConnection(con);
            //创建SQL命令
            SqlCommand com = conn.CreateCommand();
            //打开数据库连接
            conn.Open();

            
            
            try
            {
                com.CommandText += "DELETE FROM ShiYan";
                com.ExecuteNonQuery();

                /*存储数据库*/
                com.CommandText = "insert into ShiYan" + " values(@日期,@小区名,@CGI,@覆盖类型,@网络类型,@厂商,@维护部,@优化片区,@无线接通率,@无线掉线率,@切换成功率,@MR覆盖率,@空口流量,@RRC连接最大数)";

                //列名
                com.Parameters.Add(new SqlParameter("@日期", SqlDbType.NVarChar));//参数名，参数类型
                com.Parameters.Add(new SqlParameter("@小区名", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@CGI", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@覆盖类型", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@网络类型", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@厂商", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@维护部", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@优化片区", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@无线接通率", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@无线掉线率", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@切换成功率", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@MR覆盖率", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@空口流量", SqlDbType.NVarChar));
                com.Parameters.Add(new SqlParameter("@RRC连接最大数", SqlDbType.NVarChar));


                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    com.Parameters[0].Value = dt.Rows[i][0].ToString();
                    com.Parameters[1].Value = dt.Rows[i][1].ToString();
                    com.Parameters[2].Value = dt.Rows[i][2].ToString();
                    com.Parameters[3].Value = dt.Rows[i][3].ToString();
                    com.Parameters[4].Value = dt.Rows[i][4].ToString();
                    com.Parameters[5].Value = dt.Rows[i][5].ToString();
                    com.Parameters[6].Value = dt.Rows[i][6].ToString();
                    com.Parameters[7].Value = dt.Rows[i][7].ToString();
                    com.Parameters[8].Value = dt.Rows[i][8].ToString();
                    com.Parameters[9].Value = dt.Rows[i][9].ToString();
                    com.Parameters[10].Value = dt.Rows[i][10].ToString();
                    com.Parameters[11].Value = dt.Rows[i][11].ToString();
                    com.Parameters[12].Value = dt.Rows[i][12].ToString();
                    com.Parameters[13].Value = dt.Rows[i][13].ToString();
                    com.ExecuteNonQuery();
                }

                MessageBox.Show("已成功将数据导入到数据库中");
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message.ToString());
            }
            finally
            {
                conn.Close();
            }
        }

        public abstract class DB
        {
            public abstract DataTable ds(string fileName);  //导入Excel数据到数据集
            public abstract DataTable AVG(string str);//求平均值
            public abstract DataTable Num(string str,string date); //按日期求和
            public abstract DataTable DifDate();  //获取所有日期
            public abstract DataTable DifArea();  //获取所有区域
            public abstract DataTable DifManu();  //获取所有厂商
            public abstract DataTable AVGManu(string filed,string date,string manu);  //按厂商日期和厂商筛选求平均
            public abstract DataTable SUMManu(string filed, string date, string manu);  //按厂商日期和厂商筛选求和
            public abstract DataTable AVGArea(string filed,string date, string area); //按区域日期和区域筛选求平均
            public abstract DataTable SUMArea(string filed, string date, string area);  //按区域日期和区域筛选求和
            public abstract DataTable Date(string date, string filed);                  //覆盖类型
            public abstract DataTable Date1(string date, string filed);                  //厂商
            public abstract DataTable Date2(string date, string filed);                  //网络类型
        }


        /// <summary>
        /// 导入Excel中的数据到数据集中
        /// </summary>
        public class dataset:DB
        {
            public override DataTable ds(string fileName)
            {
                DataTable mytable;
                String myline;
                String[] Ary;
                mytable = new DataTable();
                mytable.Columns.Add("日期", Type.GetType("System.String")); //第一列
                mytable.Columns.Add("小区名", Type.GetType("System.String")); //第二列
                mytable.Columns.Add("CGI", Type.GetType("System.String")); //第三列
                mytable.Columns.Add("覆盖类型", Type.GetType("System.String")); //第四列
                mytable.Columns.Add("网络类型", Type.GetType("System.String")); //第五列
                mytable.Columns.Add("厂商", Type.GetType("System.String")); //第六列
                mytable.Columns.Add("维护部", Type.GetType("System.String")); //第七列
                mytable.Columns.Add("优化片区", Type.GetType("System.String")); //第八列
                mytable.Columns.Add("无线接通率", Type.GetType("System.String")); //第九列
                mytable.Columns.Add("无线掉线率", Type.GetType("System.String")); //第十列
                mytable.Columns.Add("切换成功率", Type.GetType("System.String")); //第十一列
                mytable.Columns.Add("MR覆盖率", Type.GetType("System.String")); //第十二列
                mytable.Columns.Add("空口流量", Type.GetType("System.String")); //第十三列
                mytable.Columns.Add("RRC连接最大数", Type.GetType("System.String")); //第十四列

                Dictionary<String, String> user = new Dictionary<string, string>();
                StreamReader myreader = new StreamReader(fileName, System.Text.Encoding.Default);
                while ((myline = myreader.ReadLine()) != null)
                {
                    Ary = myline.Split(new char[] { ',' });
                    DataRow dr = mytable.NewRow();
                    for (int i = 0; i < 14; i++)
                    {
                        String value = Ary[i];
                        dr[i] = value;
                    }
                    mytable.Rows.Add(dr);
                }
                mytable.Rows.RemoveAt(0);
                return mytable;
            }


            /// <summary>
            /// 平均值
            /// </summary>
            /// <param name="str">字段</param>
            /// <returns></returns>
            public override DataTable AVG(string str)
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"SELECT AVG("+ str +") FROM ShiYan");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }


            /// <summary>
            /// 按日期求和
            /// </summary>
            /// <param name="str">字段</param>
            /// <param name="date">日期</param>
            /// <returns></returns>
            public override DataTable Num(string str,string date)
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"SELECT SUM("+ str +") AS 求和 FROM ShiYan where 日期 = '"+date+"'");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

            public override DataTable DifDate()
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"SELECT distinct 日期 FROM ShiYan order by 1 asc");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

            public override DataTable DifArea()
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"SELECT distinct 优化片区 FROM ShiYan");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

            public override DataTable DifManu()
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"SELECT distinct 厂商 FROM ShiYan");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

            /// <summary>
            /// 按厂商日期和厂商筛选求平均
            /// </summary>
            /// <param name="date">日期</param>
            /// <param name="manu">厂商</param>
            /// <param name="filed">参数</param>
            /// <returns></returns>
            public override DataTable AVGManu(string filed,string date, string manu)
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"SELECT AVG("+filed+") FROM ShiYan where 日期 = '" + date+"' and 厂商 = N'"+manu+"'");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

            /// <summary>
            /// 按厂商日期和厂商筛选求求和
            /// </summary>
            /// <param name="date">日期</param>
            /// <param name="manu">厂商</param>
            /// <param name="filed">参数</param>
            /// <returns></returns>
            public override DataTable SUMManu(string filed, string date, string manu)
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"SELECT SUM(" + filed + ") FROM ShiYan where 日期 = '" + date + "' and 厂商 = N'" + manu + "'");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

            /// <summary>
            /// 按区域日期和区域筛选求平均
            /// </summary>
            /// <param name="date">日期</param>
            /// <param name="area">优化片区</param>
            /// <param name="filed">参数</param>
            /// <returns></returns>
            public override DataTable AVGArea(string filed,string date, string area)
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"SELECT AVG(" + filed + ") FROM ShiYan where 日期 = '" + date+"' and 优化片区 = N'"+area+"'");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

            /// <summary>
            /// 按区域日期和区域筛选求求和
            /// </summary>
            /// <param name="date">日期</param>
            /// <param name="manu">厂商</param>
            /// <param name="filed">参数</param>
            /// <returns></returns>
            public override DataTable SUMArea(string filed, string date, string area)
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"SELECT SUM(" + filed + ") FROM ShiYan where 日期 = '" + date + "' and 优化片区 = N'" + area + "'");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

            public override DataTable Date(string date, string filed)
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"select '"+date+"',N'"+filed+"', AVG(无线接通率),AVG(无线掉线率),AVG(切换成功率),SUM(空口流量),SUM(RRC连接最大数) FROM ShiYan WHERE 日期 = '"+date+"' and 覆盖类型 = N'"+filed+"'");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

            public override DataTable Date1(string date, string filed)
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"select '" + date + "',N'" + filed + "', AVG(无线接通率),AVG(无线掉线率),AVG(切换成功率),SUM(空口流量),SUM(RRC连接最大数) FROM ShiYan WHERE 日期 = '" + date + "' and 厂商 = N'" + filed + "'");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

            public override DataTable Date2(string date, string filed)
            {
                string P_Str_ConnectionStr = string.Format(//创建数据库连接字符串
                @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=test;Integrated Security=True");
                string P_Str_SqlStr = string.Format(//创建SQL查询字符串
                    @"select '" + date + "',N'" + filed + "', AVG(无线接通率),AVG(无线掉线率),AVG(切换成功率),SUM(空口流量),SUM(RRC连接最大数) FROM ShiYan WHERE 日期 = '" + date + "' and 优化片区 = N'" + filed + "'");
                SqlDataAdapter P_SqlDataAdapter = new SqlDataAdapter(//创建数据适配器
                    P_Str_SqlStr, P_Str_ConnectionStr);
                DataTable P_dt = new DataTable();//创建数据表
                P_SqlDataAdapter.Fill(P_dt);//填充数据表
                return P_dt;//返回数据表
            }

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            
            DataView dv = ds.DefaultView;
            dv.RowFilter = "日期 = '2019/12/1'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "日期 = '2019/12/2'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;

        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "日期 = '2019/12/3'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;

        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "日期 = '2019/12/4'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;

        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "日期 = '2019/12/5'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;

        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "日期 = '2019/12/6'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;

        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "日期 = '2019/12/7'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;

        }

        private void 室内ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "覆盖类型 = '室内'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        private void 室外ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "覆盖类型 = '室外'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "网络类型 = '1'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "网络类型 = '2'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        private void gToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "网络类型 = '5G'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        private void fDDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "网络类型 = 'FDD'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        private void nBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "网络类型 = 'NB'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        private void tDDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "网络类型 = 'TDD'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        private void 华为ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "厂商 = '华为'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        private void 诺基亚ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataView dv = ds.DefaultView;
            dv.RowFilter = "厂商 = '诺基亚'";
            ds = dv.ToTable();
            this.dataGridView1.DataSource = dv;
        }

        /// <summary>
        /// 分厂商分析功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 无线接通率ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DB d = new dataset(); //日期列表
            DataTable dd = d.DifDate();    //所有日期数据集
            List<string> df = new List<string>(); //实例化日期列表
            foreach (DataRow dr in dd.Rows)
            {
                df.Add(dr["日期"].ToString());  //日期数据集添加到日期列表
            }

            //厂商列表
            DataTable dm = d.DifManu();  //厂商数据集
            List<string> ss = new List<string>(); //实例化厂商列表
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["厂商"].ToString());  //厂商数据集添加到厂商列表
            }        
            DataTable dt = new DataTable(); //创建空数据集
            dt.Columns.Add("日期", Type.GetType("System.String")); //添加日期项
            dt.Columns.Add("厂商", Type.GetType("System.String")); //添加厂商项
            dt.Columns.Add("无线接通率", Type.GetType("System.String")); //添加无线接通率项
            foreach (string i in df)
            {
                foreach (string j in ss)
                {
                    DataTable Am = d.AVGManu("无线接通率", i, j);  //无线接通率平均值数据集
                    string am = Am.Rows[0][0].ToString();   //获取无线接通率的值
                    DataRow dr = dt.NewRow(); //数据行
                    dr[0] = i;  //日期数据
                    dr[1] = j;  //厂商数据
                    dr[2] = am;  // 无线接通率数据
                    dt.Rows.Add(dr); //绑定数据行
                }
            }
            ds = dt;  //赋值新的数据集
            this.dataGridView1.DataSource = dt;   //绑定数据集
        }

        private void 无线掉线率ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //日期列表
            DB d = new dataset();
            DataTable dd = d.DifDate();
            List<string> df = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dd.Rows)
            {
                df.Add(dr["日期"].ToString());
            }

            //厂商列表
            DataTable dm = d.DifManu();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["厂商"].ToString());
            }
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("厂商", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            foreach (string i in df)
            {
                foreach (string j in ss)
                {
                    DataTable Am = d.AVGManu("无线掉线率", i, j);
                    string am = Am.Rows[0][0].ToString();
                    DataRow dr = dt.NewRow();
                    dr[0] = i;
                    dr[1] = j;
                    dr[2] = am;
                    dt.Rows.Add(dr);
                }
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void 切换成功率ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //日期列表
            DB d = new dataset();
            DataTable dd = d.DifDate();
            List<string> df = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dd.Rows)
            {
                df.Add(dr["日期"].ToString());
            }

            //厂商列表
            DataTable dm = d.DifManu();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["厂商"].ToString());
            }
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("厂商", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            foreach (string i in df)
            {
                foreach (string j in ss)
                {
                    DataTable Am = d.AVGManu("切换成功率", i, j);
                    string am = Am.Rows[0][0].ToString();
                    DataRow dr = dt.NewRow();
                    dr[0] = i;
                    dr[1] = j;
                    dr[2] = am;
                    dt.Rows.Add(dr);
                }
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void 空口流量ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //日期列表
            DB d = new dataset();
            DataTable dd = d.DifDate();
            List<string> df = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dd.Rows)
            {
                df.Add(dr["日期"].ToString());
            }

            //厂商列表
            DataTable dm = d.DifManu();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["厂商"].ToString());
            }
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("厂商", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            foreach (string i in df)
            {
                foreach (string j in ss)
                {
                    DataTable Am = d.SUMManu("空口流量", i, j);
                    string am = Am.Rows[0][0].ToString();
                    DataRow dr = dt.NewRow();
                    dr[0] = i;
                    dr[1] = j;
                    dr[2] = am;
                    dt.Rows.Add(dr);
                }
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void rRC最大连接数ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //日期列表
            DB d = new dataset();
            DataTable dd = d.DifDate();
            List<string> df = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dd.Rows)
            {
                df.Add(dr["日期"].ToString());
            }

            //厂商列表
            DataTable dm = d.DifManu();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["厂商"].ToString());
            }
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("厂商", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));
            foreach (string i in df)
            {
                foreach (string j in ss)
                {
                    DataTable Am = d.SUMManu("RRC连接最大数", i, j);
                    string am = Am.Rows[0][0].ToString();
                    DataRow dr = dt.NewRow();
                    dr[0] = i;
                    dr[1] = j;
                    dr[2] = am;
                    dt.Rows.Add(dr);
                }
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        /// <summary>
        /// 地理区域分析功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 无线接通率ToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            
            DB d = new dataset(); //日期列表
            DataTable dd = d.DifDate();  //所有日期数据集
            List<string> df = new List<string>(); //实例化日期列表
            foreach (DataRow dr in dd.Rows)
            {
                df.Add(dr["日期"].ToString()); //日期数据集添加到日期列表
            }
            DataTable dm = d.DifArea();  //优化片区数据集
            List<string> ss = new List<string>(); //实例化优化片区列表
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());  //优化片区数据集添加到优化片区列表
            }

            DataTable dt = new DataTable();  //创建空数据集
            dt.Columns.Add("日期", Type.GetType("System.String"));   //添加日期项
            dt.Columns.Add("优化片区", Type.GetType("System.String"));   //添加优化片区项
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));  //添加无线接通率项
            foreach (string i in df)
            {
                foreach (string j in ss)
                {
                    DataTable Am = d.AVGArea("无线接通率", i, j);  //无线接通率平均值数据集
                    string am = Am.Rows[0][0].ToString();   //获取无线接通率的值
                    DataRow dr = dt.NewRow();  //数据行
                    dr[0] = i;  //日期数据
                    dr[1] = j;  //厂商数据
                    dr[2] = am;  // 无线接通率数据
                    dt.Rows.Add(dr);  //绑定数据行
                }
            }
            ds = dt;  //赋值新的数据集
            this.dataGridView1.DataSource = dt;  //绑定数据集
        }

        private void 无线掉线率ToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            //日期列表
            DB d = new dataset();
            DataTable dd = d.DifDate();
            List<string> df = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dd.Rows)
            {
                df.Add(dr["日期"].ToString());
            }

            //厂商列表
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("优化片区", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            foreach (string i in df)
            {
                foreach (string j in ss)
                {
                    DataTable Am = d.AVGArea("无线掉线率", i, j);
                    string am = Am.Rows[0][0].ToString();
                    DataRow dr = dt.NewRow();
                    dr[0] = i;
                    dr[1] = j;
                    dr[2] = am;
                    dt.Rows.Add(dr);
                }
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void 切换成功率ToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            //日期列表
            DB d = new dataset();
            DataTable dd = d.DifDate();
            List<string> df = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dd.Rows)
            {
                df.Add(dr["日期"].ToString());
            }

            //厂商列表
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("优化片区", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            foreach (string i in df)
            {
                foreach (string j in ss)
                {
                    DataTable Am = d.AVGArea("切换成功率", i, j);
                    string am = Am.Rows[0][0].ToString();
                    DataRow dr = dt.NewRow();
                    dr[0] = i;
                    dr[1] = j;
                    dr[2] = am;
                    dt.Rows.Add(dr);
                }
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void 空口流量ToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            //日期列表
            DB d = new dataset();
            DataTable dd = d.DifDate();
            List<string> df = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dd.Rows)
            {
                df.Add(dr["日期"].ToString());
            }

            //厂商列表
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("优化片区", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            foreach (string i in df)
            {
                foreach (string j in ss)
                {
                    DataTable Am = d.SUMArea("空口流量", i, j);
                    string am = Am.Rows[0][0].ToString();
                    DataRow dr = dt.NewRow();
                    dr[0] = i;
                    dr[1] = j;
                    dr[2] = am;
                    dt.Rows.Add(dr);
                }
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void rRC最大连接数ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //日期列表
            DB d = new dataset();
            DataTable dd = d.DifDate();
            List<string> df = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dd.Rows)
            {
                df.Add(dr["日期"].ToString());
            }

            //厂商列表
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("优化片区", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));
            foreach (string i in df)
            {
                foreach (string j in ss)
                {
                    DataTable Am = d.SUMArea("RRC连接最大数", i, j);
                    string am = Am.Rows[0][0].ToString();
                    DataRow dr = dt.NewRow();
                    dr[0] = i;
                    dr[1] = j;
                    dr[2] = am;
                    dt.Rows.Add(dr);
                }
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }



        /// <summary>
        /// 获取要保存的文件名称（含完整路径）
        /// </summary>
        /// <returns></returns>
        private static string GetSaveFilePath()
        {
            SaveFileDialog saveFileDig = new SaveFileDialog();
            saveFileDig.Filter = "Excel Office97-2003(*.xls)|.xls|Excel Office2007及以上(*.xlsx)|*.xlsx";
            saveFileDig.FilterIndex = 0;
            saveFileDig.OverwritePrompt = true;
            string filePath = null;
            if (saveFileDig.ShowDialog() == DialogResult.OK)
            {
                filePath = saveFileDig.FileName;
            }

            return filePath;
        }

        /// <summary>
        /// 判断是否为兼容模式
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static bool GetIsCompatible(string filePath)
        {
            return filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 创建工作薄
        /// </summary>
        /// <param name="isCompatible"></param>
        /// <returns></returns>
        private static IWorkbook CreateWorkbook(bool isCompatible)
        {
            if (isCompatible)
            {
                return new HSSFWorkbook();
            }
            else
            {
                return new XSSFWorkbook();
            }
        }

        /// <summary>
        /// 创建工作薄(依据文件流)
        /// </summary>
        /// <param name="isCompatible"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        private static IWorkbook CreateWorkbook(bool isCompatible, dynamic stream)
        {
            if (isCompatible)
            {
                return new HSSFWorkbook(stream);
            }
            else
            {
                return new XSSFWorkbook(stream);
            }
        }

        /// <summary>
        /// 创建表格头单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private static ICellStyle GetCellStyle(IWorkbook workbook)
        {
            ICellStyle style = workbook.CreateCellStyle();
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;

            return style;
        }


        /// <summary>
        /// 由DataTable导出Excel
        /// </summary>
        /// <param name="sourceTable">要导出数据的DataTable</param>
        /// <returns>Excel工作表</returns>
        public static string ExportToExcel(DataTable sourceTable, string sheetName = "result", string filePath = null)
        {
            if (sourceTable.Rows.Count <= 0) return null;

            if (string.IsNullOrEmpty(filePath))
            {
                filePath = GetSaveFilePath();   //保存的文件路径赋值filePath
            }

            if (string.IsNullOrEmpty(filePath)) return null;

            bool isCompatible = GetIsCompatible(filePath);

            IWorkbook workbook = CreateWorkbook(isCompatible);  //创建工作薄
            ICellStyle cellStyle = GetCellStyle(workbook);    //创建样式

            ISheet sheet = workbook.CreateSheet(sheetName);   //创建工作表
            IRow headerRow = sheet.CreateRow(0);   //工作表添加行
            // 添加表头
            foreach (DataColumn column in sourceTable.Columns)
            {
                ICell headerCell = headerRow.CreateCell(column.Ordinal);   
                headerCell.SetCellValue(column.ColumnName);
                headerCell.CellStyle = cellStyle;
            }

            // 添加数据
            int rowIndex = 1;

            foreach (DataRow row in sourceTable.Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);

                foreach (DataColumn column in sourceTable.Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue((row[column] ?? "").ToString());
                }

                rowIndex++;
            }
            FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            workbook.Write(fs); //保存文件
            fs.Dispose();

            sheet = null;
            headerRow = null;
            workbook = null;

            return filePath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExportToExcel(ds);
        }

        /// <summary>
        /// 自定义分析功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem2_Click_1(object sender, EventArgs e)
        { 
            DataTable dt = new DataTable(); //创建空数据集
            dt.Columns.Add("日期", Type.GetType("System.String")); //添加日期项
            dt.Columns.Add("覆盖类型", Type.GetType("System.String")); //添加覆盖类型
            dt.Columns.Add("无线接通率", Type.GetType("System.String")); //添加无线接通率
            dt.Columns.Add("无线掉线率", Type.GetType("System.String")); //添加无线掉线率
            dt.Columns.Add("切换成功率", Type.GetType("System.String")); //添加切换成功率
            dt.Columns.Add("空口流量", Type.GetType("System.String")); //添加空口流量
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String")); //添加RRC连接最大数
            DB d = new dataset(); //创建空数据集
            DataTable date1 = d.Date("2019/12/1","室外"); //获取2019/12/1的室外数据
            DataRow dr = dt.NewRow();  //创建数据集新行
            //赋值给数据集
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr); //数据集绑定新行
            DataTable date2 = d.Date("2019/12/1", "室内"); //获取2019/12/1的室内数据
            DataRow dr1 = dt.NewRow(); //创建数据集新行
            //赋值给数据集
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1); //数据集绑定新行
            ds = dt;  //赋值ds数据集
            this.dataGridView1.DataSource = dt;  //数据集赋值给表格控件
        }

        private void toolStripMenuItem3_Click_1(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date("2019/12/2", "室外");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date("2019/12/2", "室内");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;

        }

        private void toolStripMenuItem4_Click_1(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date("2019/12/3", "室外");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date("2019/12/3", "室内");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem5_Click_1(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date("2019/12/4", "室外");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date("2019/12/4", "室内");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem6_Click_1(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date("2019/12/5", "室外");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date("2019/12/5", "室内");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem7_Click_1(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date("2019/12/6", "室外");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date("2019/12/6", "室内");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem8_Click_1(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date("2019/12/7", "室外");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date("2019/12/7", "室内");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem9_Click_1(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date1("2019/12/1", "华为");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date1("2019/12/1", "诺基亚");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem10_Click_1(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date1("2019/12/2", "华为");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date1("2019/12/2", "诺基亚");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date1("2019/12/3", "华为");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date1("2019/12/3", "诺基亚");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date1("2019/12/4", "华为");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date1("2019/12/4", "诺基亚");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date1("2019/12/5", "华为");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date1("2019/12/5", "诺基亚");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date1("2019/12/6", "华为");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date1("2019/12/6", "诺基亚");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem15_Click(object sender, EventArgs e)
        {
            //创建DataTable空表
            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            DB d = new dataset();
            DataTable date1 = d.Date1("2019/12/7", "华为");
            DataRow dr = dt.NewRow();
            dr[0] = date1.Rows[0][0].ToString();
            dr[1] = date1.Rows[0][1].ToString();
            dr[2] = date1.Rows[0][2].ToString();
            dr[3] = date1.Rows[0][3].ToString();
            dr[4] = date1.Rows[0][4].ToString();
            dr[5] = date1.Rows[0][5].ToString();
            dr[6] = date1.Rows[0][6].ToString();
            dt.Rows.Add(dr);
            DataTable date2 = d.Date1("2019/12/7", "诺基亚");
            DataRow dr1 = dt.NewRow();
            dr1[0] = date2.Rows[0][0].ToString();
            dr1[1] = date2.Rows[0][1].ToString();
            dr1[2] = date2.Rows[0][2].ToString();
            dr1[3] = date2.Rows[0][3].ToString();
            dr1[4] = date2.Rows[0][4].ToString();
            dr1[5] = date2.Rows[0][5].ToString();
            dr1[6] = date2.Rows[0][6].ToString();
            dt.Rows.Add(dr1);

            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem16_Click(object sender, EventArgs e)
        {
            DB d = new dataset();
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            foreach (string j in ss)
            {
                DataTable date = d.Date2("2019/12/1", j);
                DataRow dr = dt.NewRow();
                dr[0] = date.Rows[0][0].ToString();
                dr[1] = date.Rows[0][1].ToString();
                dr[2] = date.Rows[0][2].ToString();
                dr[3] = date.Rows[0][3].ToString();
                dr[4] = date.Rows[0][4].ToString();
                dr[5] = date.Rows[0][5].ToString();
                dr[6] = date.Rows[0][6].ToString();
                dt.Rows.Add(dr);
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem17_Click(object sender, EventArgs e)
        {
            DB d = new dataset();
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            foreach (string j in ss)
            {
                DataTable date = d.Date2("2019/12/2", j);
                DataRow dr = dt.NewRow();
                dr[0] = date.Rows[0][0].ToString();
                dr[1] = date.Rows[0][1].ToString();
                dr[2] = date.Rows[0][2].ToString();
                dr[3] = date.Rows[0][3].ToString();
                dr[4] = date.Rows[0][4].ToString();
                dr[5] = date.Rows[0][5].ToString();
                dr[6] = date.Rows[0][6].ToString();
                dt.Rows.Add(dr);
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem18_Click(object sender, EventArgs e)
        {
            DB d = new dataset();
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            foreach (string j in ss)
            {
                DataTable date = d.Date2("2019/12/3", j);
                DataRow dr = dt.NewRow();
                dr[0] = date.Rows[0][0].ToString();
                dr[1] = date.Rows[0][1].ToString();
                dr[2] = date.Rows[0][2].ToString();
                dr[3] = date.Rows[0][3].ToString();
                dr[4] = date.Rows[0][4].ToString();
                dr[5] = date.Rows[0][5].ToString();
                dr[6] = date.Rows[0][6].ToString();
                dt.Rows.Add(dr);
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem19_Click(object sender, EventArgs e)
        {
            DB d = new dataset();
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            foreach (string j in ss)
            {
                DataTable date = d.Date2("2019/12/4", j);
                DataRow dr = dt.NewRow();
                dr[0] = date.Rows[0][0].ToString();
                dr[1] = date.Rows[0][1].ToString();
                dr[2] = date.Rows[0][2].ToString();
                dr[3] = date.Rows[0][3].ToString();
                dr[4] = date.Rows[0][4].ToString();
                dr[5] = date.Rows[0][5].ToString();
                dr[6] = date.Rows[0][6].ToString();
                dt.Rows.Add(dr);
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem20_Click(object sender, EventArgs e)
        {
            DB d = new dataset();
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            foreach (string j in ss)
            {
                DataTable date = d.Date2("2019/12/5", j);
                DataRow dr = dt.NewRow();
                dr[0] = date.Rows[0][0].ToString();
                dr[1] = date.Rows[0][1].ToString();
                dr[2] = date.Rows[0][2].ToString();
                dr[3] = date.Rows[0][3].ToString();
                dr[4] = date.Rows[0][4].ToString();
                dr[5] = date.Rows[0][5].ToString();
                dr[6] = date.Rows[0][6].ToString();
                dt.Rows.Add(dr);
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem21_Click(object sender, EventArgs e)
        {
            DB d = new dataset();
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            foreach (string j in ss)
            {
                DataTable date = d.Date2("2019/12/6", j);
                DataRow dr = dt.NewRow();
                dr[0] = date.Rows[0][0].ToString();
                dr[1] = date.Rows[0][1].ToString();
                dr[2] = date.Rows[0][2].ToString();
                dr[3] = date.Rows[0][3].ToString();
                dr[4] = date.Rows[0][4].ToString();
                dr[5] = date.Rows[0][5].ToString();
                dr[6] = date.Rows[0][6].ToString();
                dt.Rows.Add(dr);
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }

        private void toolStripMenuItem22_Click(object sender, EventArgs e)
        {
            DB d = new dataset();
            DataTable dm = d.DifArea();
            List<string> ss = new List<string>(); //存放你一整列所有的值
            foreach (DataRow dr in dm.Rows)
            {
                ss.Add(dr["优化片区"].ToString());
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("日期", Type.GetType("System.String"));
            dt.Columns.Add("覆盖类型", Type.GetType("System.String"));
            dt.Columns.Add("无线接通率", Type.GetType("System.String"));
            dt.Columns.Add("无线掉线率", Type.GetType("System.String"));
            dt.Columns.Add("切换成功率", Type.GetType("System.String"));
            dt.Columns.Add("空口流量", Type.GetType("System.String"));
            dt.Columns.Add("RRC连接最大数", Type.GetType("System.String"));

            foreach (string j in ss)
            {
                DataTable date = d.Date2("2019/12/7", j);
                DataRow dr = dt.NewRow();
                dr[0] = date.Rows[0][0].ToString();
                dr[1] = date.Rows[0][1].ToString();
                dr[2] = date.Rows[0][2].ToString();
                dr[3] = date.Rows[0][3].ToString();
                dr[4] = date.Rows[0][4].ToString();
                dr[5] = date.Rows[0][5].ToString();
                dr[6] = date.Rows[0][6].ToString();
                dt.Rows.Add(dr);
            }
            ds = dt;
            this.dataGridView1.DataSource = dt;
        }
    }
}
