using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace CHConfig
{
    public partial class FormStudyGetResult : Form
    {
        public int studyType;
        public FormMain formMain;
        public FormStudyGetResult()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            checkValid();
            switch (studyType)
            {
                case 1:
                    saveStudy1();
                    break;
                case 2:
                    saveStudy2();
                    break;
            }
        }

        /// <summary>
        /// 保存前检查有效性
        /// </summary>
        private void checkValid()
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value == null || cell.Value.ToString().Trim().Length == 0)
                    {
                        cell.Value = "0";
                    }
                }
            }
        }

        /// <summary>
        /// 保存学习1内容
        /// </summary>
        private byte[] saveStudy1()
        {

            string content = "";
            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    content += cell.Value.ToString() + " ";
                }
            }

            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\study1.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamWriter writer = new StreamWriter(p.FullName);
            writer.Write(content);
            writer.Dispose();
            writer.Close();
            Console.WriteLine("Edit study1: Save Ok");
            lblMsg2.Text = "学习1 内容保存成功 " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            string[] temp = content.Trim().Split(" ".ToCharArray());
            byte[] data = new byte[256]; //262-6
            for (int i = 0; i < temp.Length; i++)
            {
                try
                {
                    data[i] = byte.Parse(temp[i]);
                }
                catch (Exception e)
                {
                    formMain.msg3.Text = "学习1， index: " + i + " 无法转换，" + e.Message;
                    data[i] = 0;
                }
            }
            return data;
        }

        /// <summary>
        /// 保存学习2内容
        /// </summary>
        private byte[] saveStudy2()
        {

            string content = "";
            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    content += cell.Value.ToString() + " ";
                }
            }

            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\study2.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamWriter writer = new StreamWriter(p.FullName);
            writer.Write(content);
            writer.Dispose();
            writer.Close();
            Console.WriteLine("Edit study2: Save Ok");
            lblMsg2.Text = "学习2 内容保存成功 " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            string[] temp = content.Trim().Split(" ".ToCharArray());
            byte[] data = new byte[256]; //262-6
            for (int i = 0; i < temp.Length; i++)
            {
                try
                {
                    data[i] = byte.Parse(temp[i]);
                }
                catch (Exception e)
                {
                    formMain.msg3.Text = "学习1， index: " + i + " 无法转换，" + e.Message;
                    data[i] = 0;
                }
            }
            return data;
        }

        private void btnSendAndSave_Click(object sender, EventArgs e)
        {
            checkValid();
            byte[] data;
            switch (studyType)
            {
                case 1:
                    data = saveStudy1();
                    sendStudy1(data);
                    break;
                case 2:
                    data = saveStudy2();
                    sendStudy2(data);
                    break;
            }
        }

        /// <summary>
        /// 发送study1内容
        /// </summary>
        private void sendStudy1(byte[] data)
        {

            byte[] order = new byte[262];
            order[0] = 0xFF;
            order[1] = 0x5A;
            order[2] = 0x0C; //study1和study2的区别
            order[3] = 0x08;
            data.CopyTo(order, 4);

            byte[] orderForCheck = new byte[260];
            for (int i = 0; i < 260; i++)
            {
                orderForCheck[i] = order[i];
            }

            //for test
            string test = "";
            for (int i = 0; i < orderForCheck.Length; i++)
            {
                test += orderForCheck[i] + " ";
            }
            Console.WriteLine("OrderForCheck: " + test);

            byte[] crcResult = formMain.CRC16MODBUS(orderForCheck, orderForCheck.Length);
            formMain.checkMaster = crcResult; //此处不用

            order[260] = crcResult[0];
            order[261] = crcResult[1];

            //for test
            string testOrder = "";
            for (int i = 0; i < 262; i++)
            {
                testOrder += order[i].ToString("x2").ToUpper() + " ";
            }
            formMain.rtbxTest.Text = testOrder;

            formMain.currOp = FormMain.OP_STUDY1_SEND;

            formMain.initPort();
            formMain.msg.Text = null;
            formMain.msg2.Text = null;

            formMain.sp.Write(order, 0, order.Length);

        }

        /// <summary>
        /// 发送study2内容
        /// </summary>
        private void sendStudy2(byte[] data)
        {

            byte[] order = new byte[262];
            order[0] = 0xFF;
            order[1] = 0x5A;
            order[2] = 0x0D; //study1和study2的区别
            order[3] = 0x08;
            data.CopyTo(order, 4);

            byte[] orderForCheck = new byte[260];
            for (int i = 0; i < 260; i++)
            {
                orderForCheck[i] = order[i];
            }

            byte[] crcResult = formMain.CRC16MODBUS(orderForCheck, orderForCheck.Length);
            formMain.checkMaster = crcResult; //此处不用

            order[260] = crcResult[0];
            order[261] = crcResult[1];

            //for test
            string testOrder = "";
            for (int i = 0; i < 262; i++)
            {
                testOrder += order[i].ToString("x2").ToUpper() + " ";
            }
            formMain.rtbxTest.Text = testOrder;

            formMain.currOp = FormMain.OP_STUDY2_SEND;

            formMain.initPort();
            formMain.msg.Text = null;
            formMain.msg2.Text = null;

            formMain.sp.Write(order, 0, order.Length);

        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            //checkValid();

            string content = "";
            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value == null || cell.Value.ToString().Trim().Length == 0)
                    {
                        content += "0 ";
                    }
                    else
                    {
                        content += cell.Value.ToString() + " ";
                    }
                   
                }
            }

            string[] temp = content.Trim().Split(" ".ToCharArray());
            int[] study = new int[256];
            for (int i = 0; i < temp.Length; i ++ )
            {
                study[i] = int.Parse(temp[i]);
            }

            switch (studyType)
            {
                case 1:
                    for(int i = 4; i < 260; i ++){ //从第4个开始，到（262-2）个
                        formMain.resultStudy1[i] = study[i-4];
                    }
                    formMain.showStudy1Result();

                    break;
                case 2:
                    for(int i = 4; i < 260; i ++){ //从第4个开始，到（262-2）个
                        formMain.resultStudy2[i] = study[i-4];
                    }
                    formMain.showStudy2Result();

                    break;
            }

            

        }

    }
}
