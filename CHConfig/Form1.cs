using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Threading;
using System.IO;
using System.Text.RegularExpressions;

namespace CHConfig
{
    public partial class FormMain : Form
    {
        const string CFG_PORT = "serialport", CFG_BAUDRATE = "baudrate";
        const int PROTOCAL_ONCE_LENGHTH = 420;
        const int PROTOCAL_ROW_LENGHTH = 42;
        const int PROTOCAL_END_SEQ = 999; //结束行的序号，测试时位88，正式位999
        const int PROTOCAL_SYS_PARAM_LENGTH = 21;
        const int PROTOCAL_STUDY1_LENGTH = 262;
        const int PROTOCAL_STUDY2_LENGTH = 262;
        const int PROTOCAL_STUDY1_GET_LENGTH = 262;
        const int PROTOCAL_STUDY2_GET_LENGTH = 262; 
        bool transGoOn = true;
        bool transGoOnTest = true; //测试命令时的标志
        string port, baudrate;
        int page = 1; //当前页

        int line = 0; //当前行
     
        int resultLength;
        int[] result = new int[420]; //用来存放下位机返回数据的数组

        int resultLengthTest;
        string resultTestRow = ""; //一行(存储时以R开头的
        string resultTest = ""; //用来存放测试返回结果的字符串

        int resultLengthParam;
        int[] resultRaram = new int[22];

        int lenToRead; //找点时候用的

        int resultLengthStudy1;
        public int[] resultStudy1 = new int[262];

        int resultLengthStudy2;
        public int[] resultStudy2 = new int[262];

        int resultLengthStudy1Get;
        int[] resultStudy1Get = new int[262];

        int resultLengthStudy2Get;
        int[] resultStudy2Get = new int[262];

        public byte[] checkMaster = new byte[2]; //上位机生成的校验
        public byte[] checkSlave = new byte[2]; //下位机生成的校验
        StreamReader readerC; //在保存下发操作中需要用到的reader

        string[][] setting = new string[10][]; //用于将本业结果保存至文件 10*16

        string[] storedSeq = new string[10];

        public int currOp;
        public const int OP_REQ = 1, OP_SAVE_SEND = 2, OP_TEST = 3, OP_GET_SYS_PARAM = 4, OP_SET_SYS_PARAM = 5, OP_ZD = 6, OP_STUDY1 = 7, OP_STUDY2 = 8, OP_STUDY1_GET = 9, OP_STUDY2_GET = 10, OP_STUDY1_SEND = 11, OP_STUDY2_SEND = 12;

        bool flagWData = true; //标志data.txt当前能不能写入，应对连续点击的情况
        bool flagWSysParam = true; //标志sysparam.txt当前能不能写入，应对连续点击的情况

        private FormStudyResult formStudyResult;
        private FormStudyGetResult formStudyGetResult;

        public FormMain()
        {
            InitializeComponent();
            formStudyResult = new FormStudyResult();
            formStudyGetResult = new FormStudyGetResult();
            formStudyGetResult.formMain = this;
            for (int i = 0; i < 16; i ++ )
            {
                formStudyGetResult.dgv.Rows.Add();
                formStudyGetResult.dgv.Rows[formStudyGetResult.dgv.Rows.Count - 1].HeaderCell.Value = (i *16) + "";
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
            this.Height = Screen.FromControl(this).WorkingArea.Height - 50;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //先把data.txt中内容全部读出
            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\config.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamWriter writer = new StreamWriter(p.FullName);
            writer.WriteLine(cbxPort.SelectedItem.ToString());
            writer.WriteLine(cbxBaudrate.SelectedItem.ToString());
            writer.Close();

            msg.Text = "配置成功";
        }

        private void btnPagePrevious_Click(object sender, EventArgs e)
        {
            //检查每一格的内容合法性
            checkValid();

            //先保存当前页
            saveCurrPage();

            clearTable();

            if (page == 2)
            {
                btnPagePrevious.Enabled = false;
            }
            if (page == 20)
            {
                btnPageNext.Enabled = true;
            }

            page--;
            tbxPage.Text = page + "";

            showCurrPage();

        }

        private void btnPageNext_Click(object sender, EventArgs e)
        {
            //检查每一格的内容合法性
            checkValid();
           

            //先保存当前页
            saveCurrPage();

            clearTable();

            if (page == 1)
            {
                btnPagePrevious.Enabled = true;
            }
            if (page == 19)
            {
                btnPageNext.Enabled = false;
            }

            page++;
            tbxPage.Text = page + "";

            showCurrPage();

        }

        /// <summary>
        /// 根据当前页码组织命令并发送
        /// </summary>
        private void sendOrderReq()
        {
            string[] partPage = { page.ToString("x4").Substring(0, 2), page.ToString("x4").Substring(2, 2) };
       
            byte[] orderReq = { 0xFF, 0x5A, 0x01, 0x08, byte.Parse(Convert.ToInt32(partPage[0], 16) + ""), byte.Parse(Convert.ToInt32(partPage[1], 16) + ""), 0x00, 0x00 };

            byte[] orderForCheck = new byte[6];
            for (int i = 0; i < 6; i ++ )
            {
                orderForCheck[i] = orderReq[i];
            }
            byte[] crcCheck = CRC16MODBUS(orderForCheck, (byte)orderForCheck.Length);
            orderReq[6] = crcCheck[0];
            orderReq[7] = crcCheck[1];

            sp.Write(orderReq, 0, orderReq.Length);
        }

        /// <summary>
        /// 测试指令
        /// </summary>
        private void sendOrderTest()
        {
            currOp = OP_TEST;
            byte[] order = { 0xFF, 0x5A, 0x04, 0x08, 0x00, 0x01, 0x2B, 0x0D }; 
            sp.Write(order, 0, order.Length);
        }

        /// <summary>
        ///复位指令
        /// </summary>
        private void sendOrderReset()
        {
            byte[] order = { 0xFF, 0x5A, 0x04, 0x08, 0x00, 0x00, 0xEB, 0xCC }; //测试模式，校验待加
            sp.Write(order, 0, order.Length);
        }

        /// <summary>
        /// 获取系统参数
        /// </summary>
        private void sendOrderGetSysParam()
        {
            byte[] order = { 0xFF, 0x5A, 0x06, 0x08, 0x00, 0x00, 0x53, 0xCD }; //测试模式，校验待加
            sp.Write(order, 0, order.Length);
        }

        /// <summary>
        /// 保存并下发系统参数
        /// </summary>
        private void sendOrderSetSysParam(byte[] data)
        {
            byte[] order = new byte[21];
            order[0] = 0xFF;
            order[1] = 0x5A;
            order[2] = 0x05;
            order[3] = 0x15;
            data.CopyTo(order, 4);
           
            byte[] orderForCheck = new byte[19];
            for (int i = 0; i < 19; i ++ )
            {
                orderForCheck[i] = order[i];
            }

            byte[] crcResult = CRC16MODBUS(orderForCheck, (byte)orderForCheck.Length);
            checkMaster = crcResult; //此处不用

            order[19] = crcResult[0];
            order[20] = crcResult[1];
            
            //foreach(byte b in order){
            //    cbxOrder.Text += b.ToString("x2").ToUpper() + " ";
            //}

            sp.Write(order, 0, order.Length);
        }

        /// <summary>
        /// 找点的命令
        /// </summary>
        private void sendOrderZd()
        {
            byte[] order = { 0xFF, 0x5A, 0x07, 0x08, 0x00, 0x00, 0xAF, 0xCC };
            sp.Write(order, 0, order.Length);
        }

        /// <summary>
        /// 学习1的命令
        /// </summary>
        private void sendOrderStudy1()
        {
            byte[] order = { 0xFF, 0x5A, 0x08, 0x08, 0x00, 0x00, 0xBB, 0xCF };
            sp.Write(order, 0, order.Length);
        }

        /// <summary>
        /// 学习2的命令
        /// </summary>
        private void sendOrderStudy2()
        {
            byte[] order = { 0xFF, 0x5A, 0x09, 0x08, 0x00, 0x00, 0x47, 0xCE };
            sp.Write(order, 0, order.Length);
        }

        /// <summary>
        /// 获取学习1的命令
        /// </summary>
        private void sendOrderStudy1Get()
        {
            byte[] order = { 0xFF, 0x5A, 0x0A, 0x08, 0x00, 0x00, 0x03, 0xCE };
            sp.Write(order, 0, order.Length);
        }

        /// <summary>
        /// 获取学习2的命令
        /// </summary>
        private void sendOrderStudy2Get()
        {
            byte[] order = { 0xFF, 0x5A, 0x0B, 0x08, 0x00, 0x00, 0xFF, 0xCF };
            sp.Write(order, 0, order.Length);
        }

        private void sp_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e) //接受到1个或多个字节，会反复调用改方法
        {
           switch(currOp){
               case OP_REQ: 
                   this.Invoke(new EventHandler(getResponse)); 
                   break;
               case OP_SAVE_SEND:
                   this.Invoke(new EventHandler(getResponse2)); 
                   break;
               case OP_TEST:
                   this.Invoke(new EventHandler(getResponse3));
                   break;
               case OP_GET_SYS_PARAM:
                   this.Invoke(new EventHandler(getResponse4));
                   break;
               case OP_SET_SYS_PARAM:
                   this.Invoke(new EventHandler(getResponse5));
                   break;
               case OP_ZD:
                   this.Invoke(new EventHandler(getResponse6));
                   break;
               case OP_STUDY1:
                   this.Invoke(new EventHandler(getResponse7));
                   break;
               case OP_STUDY2:
                   this.Invoke(new EventHandler(getResponse8));
                   break;
               case OP_STUDY1_GET:
                   this.Invoke(new EventHandler(getResponse9));
                   break;
               case OP_STUDY2_GET:
                   this.Invoke(new EventHandler(getResponse10));
                   break;
               case OP_STUDY1_SEND:
                   this.Invoke(new EventHandler(getResponse11));
                   break;
               case OP_STUDY2_SEND:
                   this.Invoke(new EventHandler(getResponse12));
                   break;
           }
           
        }

        /// <summary>
        /// get the data send by machine
        /// </summary>
        private void getResponse(object sender, EventArgs e)
        {
            int length = sp.BytesToRead;
            tbxDataLength.Text = (resultLength += length) + "";
            for (int i = 0; i < length; i++ )
            {
                result[resultLength - length + i] = sp.ReadByte();
            }
            if (resultLength == PROTOCAL_ONCE_LENGHTH)
            {
                resultLength = 0;
                handleResult();
                if (transGoOn)
                {
                    if (page < 20)
                    {
                        progress.Value += progress.Step;
                        tbxDataLength.Text = null;
                        page++;
                        sendOrderReq();
                    }
                    else
                    {
                        
                    }
                }
                else //传输完毕，调整界面元素
                {
                    page = 1;
                    tbxPage.Text = page + "";
                    btnPageNext.Enabled = true;
                    btnSaveAndSend.Enabled = true;
                    showCurrPage();
                   
                }
               
               
            }
        }

        

        /// <summary>
        /// get the data send by machine
        /// </summary>
        private void getResponse2(object sender, EventArgs e)
        {
            //MessageBox.Show("resultLength: " + resultLength);
            int length = sp.BytesToRead;
            //MessageBox.Show("可读的length: " + length);
            tbxDataLength.Text = (resultLength += length) + "";
            
            for (int i = 0; i < length; i++)
            {
                checkSlave[resultLength - length + i] = (byte)sp.ReadByte();
            }
            
            if (resultLength == 2) //两位的校验
            {
                resultLength = 0;
                //比较校验值
                if (checkMaster[0] == checkSlave[0] && checkMaster[1] == checkSlave[1])
                {
                    msg.Text = "保存成功";
                    line++;
                    msg.Text = "已发送 " + line + " 行";
                    progress.Value += progress.Step;
                    
                    //继续读取并发送
                    string rowStr = "";
                    if ((rowStr = readerC.ReadLine()) != null)
                    {
                        byte[] order = ConvertLineToOrderDown(rowStr);
                        sp.Write(order, 0, order.Length);
                    }
                    else
                    {
                        readerC.Close();
                        line = 0;
                        msg.Text = "命令已读取完毕";
                        
                    }
                }
                else
                {
                    msg.Text = "保存出错：预期 " + checkMaster[0].ToString("x2").ToUpper() + " " + checkMaster[1].ToString("x2").ToUpper() +
                        "   实际 " + checkSlave[0].ToString("x2").ToUpper() + " " + checkSlave[1].ToString("x2").ToUpper();

                    readerC.Close();
                    line = 0;
                    
                }

              
                
            }
            
        }

        /// <summary>
        /// “测试”命令的返回内容处理， 返回的内容和命令1是一样的
        /// 但是，这里由于不确定会返回多少位，因此无法用byte[]来收
        /// 所以要一位一位读，存在一个大字符串中，直到读取序号88那行为止
        /// </summary>
        private void getResponse3(object sender, EventArgs e)
        {
            int length = sp.BytesToRead;
            
            for (int i = 0; i < length; i++)
            {
                resultTestRow += sp.ReadByte() + " ";
                tbxDataLength.Text = ((resultLengthTest ++) + 1) + "";

                if ((resultLengthTest % PROTOCAL_ROW_LENGHTH) == 0) //读完一行
                {
                    rtbxTest.Text += resultTestRow + "\r\n";

                    string[] temp = resultTestRow.Split(" ".ToCharArray()); //这是包含头和校验的完整内容，不是实际需要存储的部分

                    //MessageBox.Show(resultTestRow + "\r\n temp: " + temp.Length);
                    int b1 = Convert.ToInt32(temp[6]);
                    int b2 = Convert.ToInt32(temp[7]);
                    if ((b1 * 256 + b2) == PROTOCAL_END_SEQ)
                    {
                        transGoOnTest = false;
                    }

                    string rowStr = "";
                    for (int k = 0; k < 36; k++)
                    {
                        rowStr += temp[k + 4] + " ";
                    }

                    resultTest += "R" + rowStr + "\r\n";

                    line++;
                    msg.Text = "已测试 " + line + " 行";
                    msg3.Text = "    测试结果：" + resultTestRow;

                    resultTestRow = "";

                }

            }
            //!!! 测试阶段，判断，读到序号为88的时候停止（88这行是要存储的）
            //不能直接取余为0，有时比如一下子读取420行也是取余为0，就无法当一行处理了
            

            if(! transGoOnTest){ //传输完毕
                    handleResultTest();
                    resultTest = "";
                    resultLengthTest = 0;
                    page = 1;
                    line = 0;
                    tbxPage.Text = page + "";
                    btnPageNext.Enabled = true;
                    btnSaveAndSend.Enabled = true;
                    showCurrPage();
                   
             }

            
        }

        /// <summary>
        /// 获取系统参数
        /// </summary>
        private void getResponse4(object sender, EventArgs e)
        {
            int length = sp.BytesToRead;

            for (int i = 0; i < length; i++)
            {
                tbxDataLength.Text = ((resultLengthParam++) + 1) + "";
                if (resultLengthParam == PROTOCAL_SYS_PARAM_LENGTH)
                {
                    int[] sysParam = new int[PROTOCAL_SYS_PARAM_LENGTH];
                    for (int k = 0; k < PROTOCAL_SYS_PARAM_LENGTH; k ++ )
                    {
                        sysParam[k] = sp.ReadByte();
                        msg3.Text += sysParam[k].ToString("x2").ToUpper() + " ";
                    }

                    clearSysParam();

                    tvSpVersion.Text = sysParam[4] + "";
                    tvSp485.Text = sysParam[5] + "";
                    tvSpDS.Text = (sysParam[6] * 256 + sysParam[7]) + "";
                    tvSpOSFZ.Text = (sysParam[8] * 256 + sysParam[9]) + "";
                    tvSpZDFZ.Text = (sysParam[10] * 256 + sysParam[11]) + "";
                    tvSpJT.Text = sysParam[12] + "";
                    tvSpJZH.Text += (char)sysParam[13] ;
                    tvSpJZH.Text += (char)sysParam[14];
                    tvSpJZH.Text += (char)sysParam[15];
                    tvSpJZH.Text += (char)sysParam[16];
                    tvSpJZH.Text += (char)sysParam[17];
                    tvSpJZH.Text += (char)sysParam[18];
                                         
                    string sysParamData = "";
                    for (int k = 4; k < 19; k ++ )
                    {
                        sysParamData += sysParam[k] + " ";
                    }

                    //先把data.txt中内容全部读出
                    string dataPath = System.Environment.CurrentDirectory;
                    if (!Directory.Exists(dataPath))
                    {
                        Directory.CreateDirectory(dataPath);
                    }
                    FileInfo p = new FileInfo(dataPath + "\\sysparam.txt");
                    if (!p.Exists)
                    {
                        p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
                    }
                    StreamWriter writer = new StreamWriter(p.FullName);
                    writer.WriteLine(sysParamData);
                    writer.Close();

                    msg.Text = "参数获取完成 " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    break;
                }

            }

        }

        /// <summary>
        /// 接收保存系统参数之后的校验
        /// </summary>
        private void getResponse5(object sender, EventArgs e)
        {
            int length = sp.BytesToRead;
            tbxDataLength.Text = (resultLength += length) + "";

            for (int i = 0; i < length; i++)
            {
                if (resultLength == 2) //两位的校验
                {
                    resultLength = 0;

                    checkSlave[0] = (byte)sp.ReadByte();
                    checkSlave[1] = (byte)sp.ReadByte();

                    //比较校验值
                    if (checkMaster[0] == checkSlave[0] && checkMaster[1] == checkSlave[1])
                    {
                        msg.Text = "系统参数保存成功 " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    else
                    {
                        msg.Text = "系统参数保存出错：预期 " + checkMaster[0].ToString("x2").ToUpper() + " " + checkMaster[1].ToString("x2").ToUpper() +
                            "   实际 " + checkSlave[0].ToString("x2").ToUpper() + " " + checkSlave[1].ToString("x2").ToUpper();
                        
                    }
                    //btnSaveSysParam.Enabled = true;
                 
                }

            }

        }

        /// <summary>
        /// 找点命令返回的数据
        /// </summary>
        private void getResponse6(object sender, EventArgs e)
        {
            rtbxTest.Text += "\r\n串口事件*******\r\n";

            int length = sp.BytesToRead;
           
            tbxDataLength.Text = (resultLength += length) + "";

            if (resultLength > 4 && lenToRead == 0) //第四位是数据长度，根据这个来判断
            {
                rtbxTest.Text += ">=4 当前收到 " + resultLength + " 位\r\n";
               
                byte[] partHead = new byte[4];
                sp.Read(partHead, 0, 4);

                //foreach(byte b in partHead){
                 //   msg.Text += b.ToString("x2").ToUpper() + " ";
               // }

                int foundLen = partHead[3];
                rtbxTest.Text += "总长度：" + foundLen + "\r\n";
                if (foundLen > 6) //说明找到了（长度是偶数）
                {
                    lenToRead = foundLen - 6;
                  //cbxOrder.Text = "找到 " + lenToRead/2 + " 点，正在读取...";
                }
                else //说明没找到 
                {
                    cbxOrder.Text = "未找到";
                }
            }

            else if(resultLength >= (6 + lenToRead) && lenToRead > 0){

                byte[] partData = new byte[lenToRead];
                int[] points = new int[lenToRead / 2]; //一个点占两位
                rtbxTest.Text += "-- 点数 " + points.Length + " 位\r\n";

               // MessageBox.Show("要读取的数量 " + lenToRead);

                sp.Read(partData, 0, lenToRead);

                //foreach (byte b in partData)
                //{
                //    msg.Text += b.ToString("x2").ToUpper() + " ";
                //}

                rtbxTest.Text += "-- 可读 " + sp.BytesToRead + " 位\r\n";

                //MessageBox.Show("-- 点数 " + points.Length);
                for (int i = 0; i < partData.Length; i ++ )
                {
                    //MessageBox.Show("i =  " +  i + "   i%2=" + i%2);
                    if (i % 2 == 0)
                    {
                        points[i/2] = partData[i] * 256 + partData[i + 1];
                    }
                    
                }
                //MessageBox.Show("-- 打印点 ");
                foreach(int p in points){
                    rtbxTest.Text += "\r\n打印点：\r\n" + p + " ";
                    if(p == 65535){
                        cbxOrder.Text += "...";
                    }else{
                        cbxOrder.Text += p + ",";
                    }
                }

                if(cbxOrder.Text.EndsWith(",")){
                    cbxOrder.Text = cbxOrder.Text.Substring(0, cbxOrder.Text.Length - 1);
                }

                int rest = sp.BytesToRead;
                for (int i = 0; i < rest; i ++ )
                {
                    sp.ReadByte();
                }

            }

        }

        /// <summary>
        /// 学习1命令返回的数据
        /// </summary>
        private void getResponse7(object sender, EventArgs e)
        {
            int length = sp.BytesToRead;
            tbxDataLength.Text = (resultLengthStudy1 += length) + "";
            for (int i = 0; i < length; i++)
            {
                resultStudy1[resultLengthStudy1 - length + i] = sp.ReadByte();
            }
            if (resultLengthStudy1 >= PROTOCAL_STUDY1_LENGTH)
            {
                resultLengthStudy1 = 0;
                showStudy1Result(); 
            }

        }

        /// <summary>
        /// 学习2命令返回的数据
        /// </summary>
        private void getResponse8(object sender, EventArgs e)
        {
            int length = sp.BytesToRead;
            tbxDataLength.Text = (resultLengthStudy2 += length) + "";
            for (int i = 0; i < length; i++)
            {
                resultStudy2[resultLengthStudy2 - length + i] = sp.ReadByte();
            }
            if (resultLengthStudy2 >= PROTOCAL_STUDY2_LENGTH)
            {
                resultLengthStudy2 = 0;
                showStudy2Result();
            }

        }

        /// <summary>
        /// 获取学习1命令返回的数据
        /// </summary>
        private void getResponse9(object sender, EventArgs e)
        {
            int length = sp.BytesToRead;
            tbxDataLength.Text = (resultLengthStudy1Get += length) + "";
            for (int i = 0; i < length; i++)
            {
                resultStudy1Get[resultLengthStudy1Get - length + i] = sp.ReadByte();
            }
            if (resultLengthStudy1Get >= PROTOCAL_STUDY1_GET_LENGTH)
            {
                resultLengthStudy1Get = 0;
                saveStudy1GetResult();
                //showStudy1GetResult();
            }

        }

        /// <summary>
        /// 获取学习2命令返回的数据
        /// </summary>
        private void getResponse10(object sender, EventArgs e)
        {
            int length = sp.BytesToRead;
            tbxDataLength.Text = (resultLengthStudy2Get += length) + "";
            for (int i = 0; i < length; i++)
            {
                resultStudy2Get[resultLengthStudy2Get - length + i] = sp.ReadByte();
            }
            if (resultLengthStudy2Get >= PROTOCAL_STUDY2_GET_LENGTH)
            {
                resultLengthStudy2Get = 0;
                saveStudy2GetResult();
                //showStudy2GetResult();
            }

        }

        /// <summary>
        /// 发送学习1内容之后的校验
        /// </summary>
        private void getResponse11(object sender, EventArgs e)
        {
            int length = sp.BytesToRead;
            tbxDataLength.Text = (resultLength += length) + "";

            for (int i = 0; i < length; i++)
            {
                if (resultLength == 2) //两位的校验
                {
                    resultLength = 0;

                    checkSlave[0] = (byte)sp.ReadByte();
                    checkSlave[1] = (byte)sp.ReadByte();

                    //比较校验值
                    if (checkMaster[0] == checkSlave[0] && checkMaster[1] == checkSlave[1])
                    {
                        msg.Text = "学习1 保存成功 " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    else
                    {
                        msg.Text = "学习1 保存出错：预期 " + checkMaster[0].ToString("x2").ToUpper() + " " + checkMaster[1].ToString("x2").ToUpper() +
                            "   实际 " + checkSlave[0].ToString("x2").ToUpper() + " " + checkSlave[1].ToString("x2").ToUpper();
                    }

                }

            }

        }

        /// <summary>
        /// 发送学习2内容之后的校验
        /// </summary>
        private void getResponse12(object sender, EventArgs e)
        {
            int length = sp.BytesToRead;
            tbxDataLength.Text = (resultLength += length) + "";

            for (int i = 0; i < length; i++)
            {
                if (resultLength == 2) //两位的校验
                {
                    resultLength = 0;

                    checkSlave[0] = (byte)sp.ReadByte();
                    checkSlave[1] = (byte)sp.ReadByte();

                    //比较校验值
                    if (checkMaster[0] == checkSlave[0] && checkMaster[1] == checkSlave[1])
                    {
                        msg.Text = "学习2 保存成功 " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    else
                    {
                        msg.Text = "学习2 保存出错：预期 " + checkMaster[0].ToString("x2").ToUpper() + " " + checkMaster[1].ToString("x2").ToUpper() +
                            "   实际 " + checkSlave[0].ToString("x2").ToUpper() + " " + checkSlave[1].ToString("x2").ToUpper();
                    }

                }

            }

        }

        /// <summary>
        /// 把收到的当前页数据存入data文件，仅此而已
        /// </summary>
        private void handleResult ()
        {
            int[] R0 = new int[42]; int[] R1 = new int[42]; int[] R2 = new int[42]; int[] R3 = new int[42]; int[] R4 = new int[42];
            int[] R5 = new int[42]; int[] R6 = new int[42]; int[] R7 = new int[42]; int[] R8 = new int[42]; int[] R9 = new int[42];

            for (int i = 0; i < 42; i ++) { R0[i] = result[0 + i]; }
            for (int i = 0; i < 42; i ++) { R1[i] = result[42 + i]; }
            for (int i = 0; i < 42; i ++) { R2[i] = result[84 + i]; }
            for (int i = 0; i < 42; i ++) { R3[i] = result[126 + i]; }
            for (int i = 0; i < 42; i ++) { R4[i] = result[168 + i]; }
            for (int i = 0; i < 42; i ++) { R5[i] = result[210 + i]; }
            for (int i = 0; i < 42; i ++) { R6[i] = result[252 + i]; }
            for (int i = 0; i < 42; i ++) { R7[i] = result[294 + i]; }
            for (int i = 0; i < 42; i ++) { R8[i] = result[336 + i]; }
            for (int i = 0; i < 42; i ++) { R9[i] = result[378 + i]; }
           
            //保存到当前页对应的文件中
            string dataPath = System.Environment.CurrentDirectory;
            if(! Directory.Exists(dataPath)){
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\data.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamWriter writer = new StreamWriter(p.FullName, true);
            string table = "";
            int[][] Rs = {R0, R1, R2, R3, R4, R5, R6, R7, R8, R9};
            foreach(int[] RX in Rs){
               //MessageBox.Show(R1[4] + "");
                if (RX[4] == 255) { transGoOn = false; break; } //遇到行号是255表示数据终止
                table += "R";
                for (int i = 4; i < 40; i++)
                {
                    table += RX[i] + " ";
                }
                table += "\r\n";
                
            }
            writer.Write(table);
            writer.Close();

        }

        /// <summary>
        /// 把收到的测试数据存入data文件
        /// </summary>
        private void handleResultTest()
        {
            //保存到当前页对应的文件中
            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\data.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamWriter writer = new StreamWriter(p.FullName, true);
           
            writer.Write(resultTest);
            writer.Close();
        }

        /// <summary>
        /// 把当前页对应的数据展示出来
        /// </summary>
        private void showCurrPage()
        {
            R0C0.BackColor = Color.White; R1C0.BackColor = Color.White; R2C0.BackColor = Color.White; R3C0.BackColor = Color.White; R4C0.BackColor = Color.White;
            R5C0.BackColor = Color.White; R6C0.BackColor = Color.White; R7C0.BackColor = Color.White; R8C0.BackColor = Color.White; R9C0.BackColor = Color.White; 

            //先把data.txt中内容全部读出
            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\data.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamReader reader = new StreamReader(p.FullName);
            string data = reader.ReadToEnd();
            reader.Close();

            for (int i = 0; i < 10; i++ )
            {
                int rowNum = (page - 1) * 10 + i + 1;
                string temp = (rowNum / 256) + " " +(rowNum % 256);
                Regex rg = new Regex("R" + temp + ".+?\r\n", RegexOptions.Singleline);
                string row = rg.Match(data).Value;       
                if(row == null || row.Trim().Length == 0){
                    break;
                }
                row = row.Substring(1).Trim();
                string[] tempArray = row.Split(" ".ToCharArray());
                int[] R = new int[36];
                for (int k = 0; k < 36; k ++ )
                {
                    R[k] = Convert.ToInt32(tempArray[k]);
                }

                switch(i){

                    case 0:
                        R0C0.Text = "" + (R[2] * 256 + R[3]);
                        R0C1.Text = "" + (char)R[4] + (char)R[5] + (char)R[6] + (char)R[7];
                        R0C2.Text = "" + (R[8] * 256 + R[9]);
                        R0C3.Text = "" + (R[10] * 256 + R[11]); //不能直接加，第一位要乘256
                        R0C4.Text = "" + (R[12] * 256 + R[13]);
                        R0C5.Text = "" + (R[14] * 256 + R[15]);
                        R0C6.Text = "" + (R[16] * 256 + R[17]);
                        R0C7.Text = "" + (R[18] * 256 + R[19]);
                        R0C8.Text = "" + (R[20] * 256 + R[21]);
                        R0C9.Text = "" + (char)R[22] + (char)R[23] + (char)R[24] + (char)R[25];
                        R0C10.Text = "" + R[26];
                        R0C11.Text = "" + (sbyte)R[27]; //转换成有符号整数
                        R0C12.Text = "" + (R[28] * 256 + R[29]);
                        R0C13.Text = "" + (R[30] * 256 + R[31]);
                        R0C14.Text = "" + (R[32] * 256 + R[33]);
                        R0C15.Text = "" + (R[34] * 256 + R[35]);
                        break;

                    case 1:
                        R1C0.Text = "" + (R[2] * 256 + R[3]);
                        R1C1.Text = "" + (char)R[4] + (char)R[5] + (char)R[6] + (char)R[7];
                        R1C2.Text = "" + (R[8] * 256 + R[9]);
                        R1C3.Text = "" + (R[10] * 256 + R[11]); //不能直接加，第一位要乘256
                        R1C4.Text = "" + (R[12] * 256 + R[13]);
                        R1C5.Text = "" + (R[14] * 256 + R[15]);
                        R1C6.Text = "" + (R[16] * 256 + R[17]);
                        R1C7.Text = "" + (R[18] * 256 + R[19]);
                        R1C8.Text = "" + (R[20] * 256 + R[21]);
                        R1C9.Text = "" + (char)R[22] + (char)R[23] + (char)R[24] + (char)R[25];
                        R1C10.Text = "" + R[26];
                        R1C11.Text = "" + (sbyte)R[27];
                        R1C12.Text = "" + (R[28] * 256 + R[29]);
                        R1C13.Text = "" + (R[30] * 256 + R[31]);
                        R1C14.Text = "" + (R[32] * 256 + R[33]);
                        R1C15.Text = "" + (R[34] * 256 + R[35]);
                        break;

                    case 2:
                        R2C0.Text = "" + (R[2] * 256 + R[3]);
                        R2C1.Text = "" + (char)R[4] + (char)R[5] + (char)R[6] + (char)R[7];
                        R2C2.Text = "" + (R[8] * 256 + R[9]);
                        R2C3.Text = "" + (R[10] * 256 + R[11]); //不能直接加，第一位要乘256
                        R2C4.Text = "" + (R[12] * 256 + R[13]);
                        R2C5.Text = "" + (R[14] * 256 + R[15]);
                        R2C6.Text = "" + (R[16] * 256 + R[17]);
                        R2C7.Text = "" + (R[18] * 256 + R[19]);
                        R2C8.Text = "" + (R[20] * 256 + R[21]);
                        R2C9.Text = "" + (char)R[22] + (char)R[23] + (char)R[24] + (char)R[25];
                        R2C10.Text = "" + R[26];
                        R2C11.Text = "" + (sbyte)R[27];
                        R2C12.Text = "" + (R[28] * 256 + R[29]);
                        R2C13.Text = "" + (R[30] * 256 + R[31]);
                        R2C14.Text = "" + (R[32] * 256 + R[33]);
                        R2C15.Text = "" + (R[34] * 256 + R[35]);
                        break;

                    case 3:
                        R3C0.Text = "" + (R[2] * 256 + R[3]);
                        R3C1.Text = "" + (char)R[4] + (char)R[5] + (char)R[6] + (char)R[7];
                        R3C2.Text = "" + (R[8] * 256 + R[9]);
                        R3C3.Text = "" + (R[10] * 256 + R[11]); //不能直接加，第一位要乘256
                        R3C4.Text = "" + (R[12] * 256 + R[13]);
                        R3C5.Text = "" + (R[14] * 256 + R[15]);
                        R3C6.Text = "" + (R[16] * 256 + R[17]);
                        R3C7.Text = "" + (R[18] * 256 + R[19]);
                        R3C8.Text = "" + (R[20] * 256 + R[21]);
                        R3C9.Text = "" + (char)R[22] + (char)R[23] + (char)R[24] + (char)R[25];
                        R3C10.Text = "" + R[26];
                        R3C11.Text = "" + (sbyte)R[27];
                        R3C12.Text = "" + (R[28] * 256 + R[29]);
                        R3C13.Text = "" + (R[30] * 256 + R[31]);
                        R3C14.Text = "" + (R[32] * 256 + R[33]);
                        R3C15.Text = "" + (R[34] * 256 + R[35]);
                        break;

                    case 4:
                        R4C0.Text = "" + (R[2] * 256 + R[3]);
                        R4C1.Text = "" + (char)R[4] + (char)R[5] + (char)R[6] + (char)R[7];
                        R4C2.Text = "" + (R[8] * 256 + R[9]);
                        R4C3.Text = "" + (R[10] * 256 + R[11]); //不能直接加，第一位要乘256
                        R4C4.Text = "" + (R[12] * 256 + R[13]);
                        R4C5.Text = "" + (R[14] * 256 + R[15]);
                        R4C6.Text = "" + (R[16] * 256 + R[17]);
                        R4C7.Text = "" + (R[18] * 256 + R[19]);
                        R4C8.Text = "" + (R[20] * 256 + R[21]);
                        R4C9.Text = "" + (char)R[22] + (char)R[23] + (char)R[24] + (char)R[25];
                        R4C10.Text = "" + R[26];
                        R4C11.Text = "" + (sbyte)R[27];
                        R4C12.Text = "" + (R[28] * 256 + R[29]);
                        R4C13.Text = "" + (R[30] * 256 + R[31]);
                        R4C14.Text = "" + (R[32] * 256 + R[33]);
                        R4C15.Text = "" + (R[34] * 256 + R[35]);
                        break;

                    case 5:
                        R5C0.Text = "" + (R[2] * 256 + R[3]);
                        R5C1.Text = "" + (char)R[4] + (char)R[5] + (char)R[6] + (char)R[7];
                        R5C2.Text = "" + (R[8] * 256 + R[9]);
                        R5C3.Text = "" + (R[10] * 256 + R[11]); //不能直接加，第一位要乘256
                        R5C4.Text = "" + (R[12] * 256 + R[13]);
                        R5C5.Text = "" + (R[14] * 256 + R[15]);
                        R5C6.Text = "" + (R[16] * 256 + R[17]);
                        R5C7.Text = "" + (R[18] * 256 + R[19]);
                        R5C8.Text = "" + (R[20] * 256 + R[21]);
                        R5C9.Text = "" + (char)R[22] + (char)R[23] + (char)R[24] + (char)R[25];
                        R5C10.Text = "" + R[26];
                        R5C11.Text = "" + (sbyte)R[27];
                        R5C12.Text = "" + (R[28] * 256 + R[29]);
                        R5C13.Text = "" + (R[30] * 256 + R[31]);
                        R5C14.Text = "" + (R[32] * 256 + R[33]);
                        R5C15.Text = "" + (R[34] * 256 + R[35]);
                        break;

                    case 6:
                        R6C0.Text = "" + (R[2] * 256 + R[3]);
                        R6C1.Text = "" + (char)R[4] + (char)R[5] + (char)R[6] + (char)R[7];
                        R6C2.Text = "" + (R[8] * 256 + R[9]);
                        R6C3.Text = "" + (R[10] * 256 + R[11]); //不能直接加，第一位要乘256
                        R6C4.Text = "" + (R[12] * 256 + R[13]);
                        R6C5.Text = "" + (R[14] * 256 + R[15]);
                        R6C6.Text = "" + (R[16] * 256 + R[17]);
                        R6C7.Text = "" + (R[18] * 256 + R[19]);
                        R6C8.Text = "" + (R[20] * 256 + R[21]);
                        R6C9.Text = "" + (char)R[22] + (char)R[23] + (char)R[24] + (char)R[25];
                        R6C10.Text = "" + R[26];
                        R6C11.Text = "" + (sbyte)R[27];
                        R6C12.Text = "" + (R[28] * 256 + R[29]);
                        R6C13.Text = "" + (R[30] * 256 + R[31]);
                        R6C14.Text = "" + (R[32] * 256 + R[33]);
                        R6C15.Text = "" + (R[34] * 256 + R[35]);
                        break;

                    case 7:
                        R7C0.Text = "" + (R[2] * 256 + R[3]);
                        R7C1.Text = "" + (char)R[4] + (char)R[5] + (char)R[6] + (char)R[7];
                        R7C2.Text = "" + (R[8] * 256 + R[9]);
                        R7C3.Text = "" + (R[10] * 256 + R[11]); //不能直接加，第一位要乘256
                        R7C4.Text = "" + (R[12] * 256 + R[13]);
                        R7C5.Text = "" + (R[14] * 256 + R[15]);
                        R7C6.Text = "" + (R[16] * 256 + R[17]);
                        R7C7.Text = "" + (R[18] * 256 + R[19]);
                        R7C8.Text = "" + (R[20] * 256 + R[21]);
                        R7C9.Text = "" + (char)R[22] + (char)R[23] + (char)R[24] + (char)R[25];
                        R7C10.Text = "" + R[26];
                        R7C11.Text = "" + (sbyte)R[27];
                        R7C12.Text = "" + (R[28] * 256 + R[29]);
                        R7C13.Text = "" + (R[30] * 256 + R[31]);
                        R7C14.Text = "" + (R[32] * 256 + R[33]);
                        R7C15.Text = "" + (R[34] * 256 + R[35]);
                        break;

                    case 8:
                        R8C0.Text = "" + (R[2] * 256 + R[3]);
                        R8C1.Text = "" + (char)R[4] + (char)R[5] + (char)R[6] + (char)R[7];
                        R8C2.Text = "" + (R[8] * 256 + R[9]);
                        R8C3.Text = "" + (R[10] * 256 + R[11]); //不能直接加，第一位要乘256
                        R8C4.Text = "" + (R[12] * 256 + R[13]);
                        R8C5.Text = "" + (R[14] * 256 + R[15]);
                        R8C6.Text = "" + (R[16] * 256 + R[17]);
                        R8C7.Text = "" + (R[18] * 256 + R[19]);
                        R8C8.Text = "" + (R[20] * 256 + R[21]);
                        R8C9.Text = "" + (char)R[22] + (char)R[23] + (char)R[24] + (char)R[25];
                        R8C10.Text = "" + R[26];
                        R8C11.Text = "" + (sbyte)R[27];
                        R8C12.Text = "" + (R[28] * 256 + R[29]);
                        R8C13.Text = "" + (R[30] * 256 + R[31]);
                        R8C14.Text = "" + (R[32] * 256 + R[33]);
                        R8C15.Text = "" + (R[34] * 256 + R[35]);
                        break;

                    case 9:
                        R9C0.Text = "" + (R[2] * 256 + R[3]);
                        R9C1.Text = "" + (char)R[4] + (char)R[5] + (char)R[6] + (char)R[7];
                        R9C2.Text = "" + (R[8] * 256 + R[9]);
                        R9C3.Text = "" + (R[10] * 256 + R[11]); //不能直接加，第一位要乘256
                        R9C4.Text = "" + (R[12] * 256 + R[13]);
                        R9C5.Text = "" + (R[14] * 256 + R[15]);
                        R9C6.Text = "" + (R[16] * 256 + R[17]);
                        R9C7.Text = "" + (R[18] * 256 + R[19]);
                        R9C8.Text = "" + (R[20] * 256 + R[21]);
                        R9C9.Text = "" + (char)R[22] + (char)R[23] + (char)R[24] + (char)R[25];
                        R9C10.Text = "" + R[26];
                        R9C11.Text = "" + (sbyte)R[27];
                        R9C12.Text = "" + (R[28] * 256 + R[29]);
                        R9C13.Text = "" + (R[30] * 256 + R[31]);
                        R9C14.Text = "" + (R[32] * 256 + R[33]);
                        R9C15.Text = "" + (R[34] * 256 + R[35]);
                        break;
                }

            }

            //检查序号有效性，若序号>1000，则表明改行有问题，该行背景变红
            checkSeqAfterShow();

        }

        /// <summary>
        /// 保存当前页面的数据
        /// </summary>
        private bool saveCurrPage()
        {
            flagWData = false;
            checkSeqBeforeSave();

            //先把data.txt中内容全部读出
            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\data.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamReader reader = new StreamReader(p.FullName);
            string data = reader.ReadToEnd();
            reader.Close();

            //把界面上的值写到R0 ~ R9中
            RichTextBox[] v0 = { R0C0, R0C1, R0C2, R0C3, R0C4, R0C5, R0C6, R0C7, R0C8, R0C9, R0C10, R0C11, R0C12, R0C13, R0C14, R0C15 };
            RichTextBox[] v1 = { R1C0, R1C1, R1C2, R1C3, R1C4, R1C5, R1C6, R1C7, R1C8, R1C9, R1C10, R1C11, R1C12, R1C13, R1C14, R1C15 };
            RichTextBox[] v2 = { R2C0, R2C1, R2C2, R2C3, R2C4, R2C5, R2C6, R2C7, R2C8, R2C9, R2C10, R2C11, R2C12, R2C13, R2C14, R2C15 };
            RichTextBox[] v3 = { R3C0, R3C1, R3C2, R3C3, R3C4, R3C5, R3C6, R3C7, R3C8, R3C9, R3C10, R3C11, R3C12, R3C13, R3C14, R3C15 };
            RichTextBox[] v4 = { R4C0, R4C1, R4C2, R4C3, R4C4, R4C5, R4C6, R4C7, R4C8, R4C9, R4C10, R4C11, R4C12, R4C13, R4C14, R4C15 };
            RichTextBox[] v5 = { R5C0, R5C1, R5C2, R5C3, R5C4, R5C5, R5C6, R5C7, R5C8, R5C9, R5C10, R5C11, R5C12, R5C13, R5C14, R5C15 };
            RichTextBox[] v6 = { R6C0, R6C1, R6C2, R6C3, R6C4, R6C5, R6C6, R6C7, R6C8, R6C9, R6C10, R6C11, R6C12, R6C13, R6C14, R6C15, };
            RichTextBox[] v7 = { R7C0, R7C1, R7C2, R7C3, R7C4, R7C5, R7C6, R7C7, R7C8, R7C9, R7C10, R7C11, R7C12, R7C13, R7C14, R7C15 };
            RichTextBox[] v8 = { R8C0, R8C1, R8C2, R8C3, R8C4, R8C5, R8C6, R8C7, R8C8, R8C9, R8C10, R8C11, R8C12, R8C13, R8C14, R8C15 };
            RichTextBox[] v9 = { R9C0, R9C1, R9C2, R9C3, R9C4, R9C5, R9C6, R9C7, R9C8, R9C9, R9C10, R9C11, R9C12, R9C13, R9C14, R9C15 };
            RichTextBox[][] views = { v0, v1, v2, v3, v4, v5, v6, v7, v8, v9 };

            int r = 0; //行号
            foreach(RichTextBox[] v in views) 
            {
                string strRow = "R";
                int valueInt;
                string valueStr;
                int[] cell;
                r++;
                string rowNum = "";
               //每个单元格单独处理，不写通用的转换算法了
                
                //2位 表格行号
                cell = new int[2];

                valueInt = ((page - 1) * 10) + r;
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";     rowNum = cell[0] + " " + cell[1] + " "; //行号记住，后面用来进行正则替换
          
                //2位 序号
                cell = new int[2];
                valueInt = int.Parse(v[0].Text); if (valueInt == 65535) { continue; } //如果序号被置为65535，则改行不予保存
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //4位 步骤名
                cell = new int[4];
                valueStr = v[1].Text;
                char[] chars = valueStr.ToCharArray();
                cell[0] = (int)chars[0]; cell[1] = (int)chars[1]; cell[2] = (int)chars[2]; cell[3] = (int)chars[3];
                strRow += cell[0] + " " + cell[1] + " " + cell[2] + " " + cell[3] + " ";

                //2位 标准值
                cell = new int[2];
                valueInt = int.Parse(v[2].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //2位 测量值
                cell = new int[2];
                valueInt = int.Parse(v[3].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //2位 上百分% 
                cell = new int[2];
                valueInt = int.Parse(v[4].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //2位 下百分%
                cell = new int[2];
                valueInt = int.Parse(v[5].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //2位 高点 
                cell = new int[2];
                valueInt = int.Parse(v[6].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //2位 低点
                cell = new int[2];
                valueInt = int.Parse(v[7].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //2位 延时 (0 ~ 9990)
                //2位 测量值
                cell = new int[2];
                valueInt = int.Parse(v[8].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //4位 模式
                cell = new int[4];
                valueStr = v[9].Text;
                char[] chars2 = valueStr.ToCharArray();
                cell[0] = (int)chars2[0]; cell[1] = (int)chars2[1]; cell[2] = (int)chars2[2]; cell[3] = (int)chars2[3];
                strRow += cell[0] + " " + cell[1] + " " + cell[2] + " " + cell[3] + " ";

                //1位 比例K
                cell = new int[1];
                valueInt = int.Parse(v[10].Text);
                cell[0] = valueInt;
                strRow += cell[0] + " ";

                //1位 偏移B
                cell = new int[1];
                valueInt = int.Parse(v[11].Text);
                cell[0] = valueInt;
                strRow += cell[0] + " ";

                //2位 上百分%2
                cell = new int[2];
                valueInt = int.Parse(v[12].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //2位 下百分%2
                cell = new int[2];
                valueInt = int.Parse(v[13].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //2位 高点2
                cell = new int[2];
                valueInt = int.Parse(v[14].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                //2位 低点2
                cell = new int[2];
                valueInt = int.Parse(v[15].Text);
                cell[0] = valueInt / 256; cell[1] = valueInt % 256;
                strRow += cell[0] + " " + cell[1] + " ";

                strRow += "\r\n";
                //用strRow 替换单行
                  
                Regex rg1 = new Regex("R"+rowNum+".+?\r\n", RegexOptions.Singleline);

                Match m = rg1.Match(data);
                if (!m.Success)
                {
                    data += strRow;
                }
                
                data = rg1.Replace(data, strRow);

            }

            try
            {
                StreamWriter writer = new StreamWriter(p.FullName);
                writer.Write(data);
                writer.Close();

                msg.Text = "第 " + page + " 页保存成功";
                
                return true;
            }catch(IOException ex){
                readerC.Close();
                MessageBox.Show("本页保存失败，请勿过快点击 \r\n" + ex.Message);
                return false;
            }
            
        }

        /// <summary>
        /// 显示当页数据之后检查每行序号的有效性
        /// </summary>
        private void checkSeqAfterShow() {
            RichTextBox[] views = { R0C0, R1C0, R2C0, R3C0, R4C0, R5C0, R6C0, R7C0, R8C0, R9C0 };
            for(int i = 0; i < views.Length; i ++ ){
                RichTextBox v = views[i];
                 
                try
                {
                     int seq = int.Parse(v.Text);
                     storedSeq[i] = seq.ToString();
                     if (seq > 1000)
                     {
                         storedSeq[i] = seq.ToString(); //这个数组只存异常的序号
                         v.Text = (seq % 1000) + "";
                         v.BackColor = Color.Red;
                     }
                     else
                     {
                         storedSeq[i] = null; //如果序号正常或者是空行，就存null
                     }
                }catch(Exception ex){
                    storedSeq[i] = null;
                    if(v.Text.Trim().Length != 0){
                        msg.Text = "序号转换出错：" + ex.Message + "  (" + v.Text + ")";
                        v.BackColor = Color.Red;
                    }
                    
                }
               
            }
            
        }

        /// <summary>
        /// 保存当页数据之前检查每行序号的有效性
        /// </summary>
        private void checkSeqBeforeSave()
        {
            RichTextBox[] views = { R0C0, R1C0, R2C0, R3C0, R4C0, R5C0, R6C0, R7C0, R8C0, R9C0 };
            for (int i = 0; i < views.Length; i++)
            {
                RichTextBox v = views[i];
                if(storedSeq[i] != null){ //说明收到的时候是异常的行
                    //比较现在的框里的值和显示的时候是否一样
                    int seq = int.Parse(v.Text);
                    if (seq == (int.Parse(storedSeq[i])) % 1000) //如果一样，说明没改，还是保存原异常值
                    {
                        v.Text = storedSeq[i];
                    }
                }
            }

        }

        /// <summary>
        /// 发送命令时检查命令中的序号，如果大于1000，就取余
        /// </summary>
        private void checkSeqBeforeSend(byte[] order)
        {
            //[6] [7] 两位是序号
            int seq = order[6]*256 + order[7];
            if(seq > 1000){
                int newSeq = seq % 1000;
                order[6] = (byte)(newSeq / 256);
                order[7] = (byte)(newSeq % 256);
            }
        }

         /// <summary>
        /// 显示系统参数
        /// </summary>
        private void showSysParam()
        {
            //先把data.txt中内容全部读出
            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\sysparam.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamReader reader = new StreamReader(p.FullName);
            string data = reader.ReadToEnd();
            reader.Close();
            if(data != null && data.Trim().Length > 0){
                string[] temp = data.Trim().Split(" ".ToCharArray());
                if (temp.Length < 15)
                {
                    MessageBox.Show("数据长度不足15，请确保sysparam配置文件内容完整");
                    return;
                }
                else
                {
                    int[] param = new int[15];
                    for (int i = 0; i < 15; i ++ )
                    {
                        param[i] = Convert.ToInt32(temp[i]);
                    }

                    tvSpVersion.Text = param[0] + "";
                    tvSp485.Text = param[1] + "";
                    tvSpDS.Text = (param[2] * 256 + param[3]) + "";
                    tvSpOSFZ.Text = (param[4] * 256 + param[5]) + "";
                    tvSpZDFZ.Text = (param[6] * 256 + param[7]) + "";
                    tvSpJT.Text = param[8] + "";
                    tvSpJZH.Text += (char)param[9];
                    tvSpJZH.Text += (char)param[10];
                    tvSpJZH.Text += (char)param[11];
                    tvSpJZH.Text += (char)param[12];
                    tvSpJZH.Text += (char)param[13];
                    tvSpJZH.Text += (char)param[14];
                }

            }


        }
        /// <summary>
        /// 清空界面
        /// </summary>
        private void clearTable()
        {
            RichTextBox[] views = {
            R0C0, R0C1, R0C2, R0C3, R0C4, R0C5, R0C6, R0C7, R0C8, R0C9, R0C10, R0C11, R0C12, R0C13, R0C14, R0C15,
            R1C0, R1C1, R1C2, R1C3, R1C4, R1C5, R1C6, R1C7, R1C8, R1C9, R1C10, R1C11, R1C12, R1C13, R1C14, R1C15,
            R2C0, R2C1, R2C2, R2C3, R2C4, R2C5, R2C6, R2C7, R2C8, R2C9, R2C10, R2C11, R2C12, R2C13, R2C14, R2C15,
            R3C0, R3C1, R3C2, R3C3, R3C4, R3C5, R3C6, R3C7, R3C8, R3C9, R3C10, R3C11, R3C12, R3C13, R3C14, R3C15,
            R4C0, R4C1, R4C2, R4C3, R4C4, R4C5, R4C6, R4C7, R4C8, R4C9, R4C10, R4C11, R4C12, R4C13, R4C14, R4C15,
            R5C0, R5C1, R5C2, R5C3, R5C4, R5C5, R5C6, R5C7, R5C8, R5C9, R5C10, R5C11, R5C12, R5C13, R5C14, R5C15,
            R6C0, R6C1, R6C2, R6C3, R6C4, R6C5, R6C6, R6C7, R6C8, R6C9, R6C10, R6C11, R6C12, R6C13, R6C14, R6C15,
            R7C0, R7C1, R7C2, R7C3, R7C4, R7C5, R7C6, R7C7, R7C8, R7C9, R7C10, R7C11, R7C12, R7C13, R7C14, R7C15,
            R8C0, R8C1, R8C2, R8C3, R8C4, R8C5, R8C6, R8C7, R8C8, R8C9, R8C10, R8C11, R8C12, R8C13, R8C14, R8C15,
            R9C0, R9C1, R9C2, R9C3, R9C4, R9C5, R9C6, R9C7, R9C8, R9C9, R9C10, R9C11, R9C12, R9C13, R9C14, R9C15  };

            foreach (RichTextBox v in views) { v.Text = null; v.BackColor = Color.White; }

            RichTextBox[] views2 = { R0C3, R1C3, R2C3, R3C3, R4C3, R5C3, R6C3, R7C3, R8C3, R9C3, };

            foreach (RichTextBox v in views2) { v.BackColor = Color.LemonChiffon; }

        }

        /// <summary>
        /// 清空界面上的系统参数
        /// </summary>
        private void clearSysParam()
        {
            tvSpVersion.Text = null;
            tvSp485.Text = null;
            tvSpDS.Text = null;
            tvSpOSFZ.Text = null;
            tvSpZDFZ.Text = null;
            tvSpJT.Text = null;
            tvSpJZH.Text = null;
        }

        /// <summary>
        /// 重新准备一下串口
        /// </summary>
        public void initPort()
        {
            resultLength = 0;

            if (sp.IsOpen)
            {
                sp.DiscardInBuffer();
                sp.DiscardOutBuffer();
                sp.Close();
            }
            try
            {
                sp.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开串口出错：\r\n" + ex.Message);
            }
            
        }

        /// <summary>
        /// 检查每一格的数据合法性（针对每一列进行检查）
        /// </summary>
        private void checkValid()
        {
            RichTextBox[] col0 = { R0C0, R1C0, R2C0, R3C0, R4C0, R5C0, R6C0, R7C0, R8C0, R9C0 };
            RichTextBox[] col1 = { R0C1, R1C1, R2C1, R3C1, R4C1, R5C1, R6C1, R7C1, R8C1, R9C1 };
            RichTextBox[] col2 = { R0C2, R1C2, R2C2, R3C2, R4C2, R5C2, R6C2, R7C2, R8C2, R9C2 };
            RichTextBox[] col3 = { R0C3, R1C3, R2C3, R3C3, R4C3, R5C3, R6C3, R7C3, R8C3, R9C3 };
            RichTextBox[] col4 = { R0C4, R1C4, R2C4, R3C4, R4C4, R5C4, R6C4, R7C4, R8C4, R9C4 };
            RichTextBox[] col5 = { R0C5, R1C5, R2C5, R3C5, R4C5, R5C5, R6C5, R7C5, R8C5, R9C5 };
            RichTextBox[] col6 = { R0C6, R1C6, R2C6, R3C6, R4C6, R5C6, R6C6, R7C6, R8C6, R9C6 };
            RichTextBox[] col7 = { R0C7, R1C7, R2C7, R3C7, R4C7, R5C7, R6C7, R7C7, R8C7, R9C7 };
            RichTextBox[] col8 = { R0C8, R1C8, R2C8, R3C8, R4C8, R5C8, R6C8, R7C8, R8C8, R9C8 };
            RichTextBox[] col9 = { R0C9, R1C9, R2C9, R3C9, R4C9, R5C9, R6C9, R7C9, R8C9, R9C9 };
            RichTextBox[] col10 = { R0C10, R1C10, R2C10, R3C10, R4C10, R5C10, R6C10, R7C10, R8C10, R9C10 };
            RichTextBox[] col11 = { R0C11, R1C11, R2C11, R3C11, R4C11, R5C11, R6C11, R7C11, R8C11, R9C11 };
            RichTextBox[] col12 = { R0C12, R1C12, R2C12, R3C12, R4C12, R5C12, R6C12, R7C12, R8C12, R9C12 };
            RichTextBox[] col13 = { R0C13, R1C13, R2C13, R3C13, R4C13, R5C13, R6C13, R7C13, R8C13, R9C13 };
            RichTextBox[] col14 = { R0C14, R1C14, R2C14, R3C14, R4C14, R5C14, R6C14, R7C14, R8C14, R9C14 };
            RichTextBox[] col15 = { R0C15, R1C15, R2C15, R3C15, R4C15, R5C15, R6C15, R7C15, R8C15, R9C15 };

            //序号
            foreach (RichTextBox v in col0) { 
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //步骤名
            foreach (RichTextBox v in col1) { 
                if (v.Text.Trim().Equals("")) { v.Text = "0000"; } 
            }
            //标准值
            foreach (RichTextBox v in col2) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //测量值
            foreach (RichTextBox v in col3) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //上百分
            foreach (RichTextBox v in col4) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //下百分
            foreach (RichTextBox v in col5) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //高点
            foreach (RichTextBox v in col6) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //低点
            foreach (RichTextBox v in col7) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //延时
            foreach (RichTextBox v in col8) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //模式
            foreach (RichTextBox v in col9) { 
                if (v.Text.Trim().Equals("")) { v.Text = "0000"; } 
            }
            //比例k
            foreach (RichTextBox v in col10) { 
                if (v.Text.Trim().Equals("")) { v.Text = 255 + ""; } 
            }
            //迁移b
            foreach (RichTextBox v in col11) { 
                if (v.Text.Trim().Equals("")) { v.Text = 255 + ""; } 
            }
            //上百分2
            foreach (RichTextBox v in col12) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //下百分2
            foreach (RichTextBox v in col13) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //高点2
            foreach (RichTextBox v in col14) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }
            //低点2
            foreach (RichTextBox v in col15) {
                if (v.Text.Trim().Equals("")) { v.Text = 65535 + ""; } 
            }

        }

        /// <summary>
        /// 检查系统参数部分的合法性
        /// </summary>
        private bool checkValidSysparam()
        {
          string version =  tvSpVersion.Text.Trim(); //1
          string sp485 = tvSp485.Text.Trim(); //1
          string ds = tvSpDS.Text.Trim(); //2
          string osfz = tvSpOSFZ.Text.Trim(); //2
          string zdfz = tvSpZDFZ.Text.Trim(); //2
          string jt = tvSpJT.Text.Trim(); //1
          string jzh = tvSpJZH.Text.Trim(); //6

          int temp;
          try { temp = Convert.ToInt32(version); if (temp > 255) { msg.Text = "版本号 数值应小于255"; return false; } }
          catch (Exception ex) { msg.Text = "版本号 应为数字"; return false; }

          try { temp = Convert.ToInt32(sp485); if (temp > 255) { msg.Text = "485地址 数值应小于255"; return false; } }
          catch (Exception ex) { msg.Text = "485地址 应为数字"; return false; }

          try { temp = Convert.ToInt32(ds); if (temp > 255) { msg.Text = "点数 数值应小于65535"; return false; } }
          catch (Exception ex) { msg.Text = "点数 应为数字"; return false; }

          try { temp = Convert.ToInt32(osfz); if (temp > 255) { msg.Text = "OS阈值 数值应小于65535"; return false; } }
          catch (Exception ex) { msg.Text = "OS阈值 应为数字"; return false; }

          try { temp = Convert.ToInt32(zdfz); if (temp > 255) { msg.Text = "找点阈值 数值应小于65535"; return false; } }
          catch (Exception ex) { msg.Text = "找点阈值 应为数字"; return false; }

          try { temp = Convert.ToInt32(jt); if (temp > 255) { msg.Text = "机台 数值应小于255"; return false; } }
          catch (Exception ex) { msg.Text = "机台 应为数字"; return false; }

          if (jzh.Length > 6) { //截取6位
              tvSpJZH.Text = jzh.Substring(0, 6);
          }
          if (jzh.Length < 6) //用空格补齐6位
          {
              for (int i = jzh.Length; i < 6; i ++ )
              {
                  jzh += " ";
              }
              tvSpJZH.Text = jzh;
          }

          return true;
        }

        /// <summary>
        /// 检查每一格的数据合法性（针对每一列进行检查）
        /// </summary>
        private void clearTableAfterCheckValid()
        {
            RichTextBox[] col0 = { R0C0, R1C0, R2C0, R3C0, R4C0, R5C0, R6C0, R7C0, R8C0, R9C0 };
            RichTextBox[] col1 = { R0C1, R1C1, R2C1, R3C1, R4C1, R5C1, R6C1, R7C1, R8C1, R9C1 };
            RichTextBox[] col2 = { R0C2, R1C2, R2C2, R3C2, R4C2, R5C2, R6C2, R7C2, R8C2, R9C2 };
            RichTextBox[] col3 = { R0C3, R1C3, R2C3, R3C3, R4C3, R5C3, R6C3, R7C3, R8C3, R9C3 };
            RichTextBox[] col4 = { R0C4, R1C4, R2C4, R3C4, R4C4, R5C4, R6C4, R7C4, R8C4, R9C4 };
            RichTextBox[] col5 = { R0C5, R1C5, R2C5, R3C5, R4C5, R5C5, R6C5, R7C5, R8C5, R9C5 };
            RichTextBox[] col6 = { R0C6, R1C6, R2C6, R3C6, R4C6, R5C6, R6C6, R7C6, R8C6, R9C6 };
            RichTextBox[] col7 = { R0C7, R1C7, R2C7, R3C7, R4C7, R5C7, R6C7, R7C7, R8C7, R9C7 };
            RichTextBox[] col8 = { R0C8, R1C8, R2C8, R3C8, R4C8, R5C8, R6C8, R7C8, R8C8, R9C8 };
            RichTextBox[] col9 = { R0C9, R1C9, R2C9, R3C9, R4C9, R5C9, R6C9, R7C9, R8C9, R9C9 };
            RichTextBox[] col10 = { R0C10, R1C10, R2C10, R3C10, R4C10, R5C10, R6C10, R7C10, R8C10, R9C10 };
            RichTextBox[] col11 = { R0C11, R1C11, R2C11, R3C11, R4C11, R5C11, R6C11, R7C11, R8C11, R9C11 };
            RichTextBox[] col12 = { R0C12, R1C12, R2C12, R3C12, R4C12, R5C12, R6C12, R7C12, R8C12, R9C12 };
            RichTextBox[] col13 = { R0C13, R1C13, R2C13, R3C13, R4C13, R5C13, R6C13, R7C13, R8C13, R9C13 };
            RichTextBox[] col14 = { R0C14, R1C14, R2C14, R3C14, R4C14, R5C14, R6C14, R7C14, R8C14, R9C14 };
            RichTextBox[] col15 = { R0C15, R1C15, R2C15, R3C15, R4C15, R5C15, R6C15, R7C15, R8C15, R9C15 };

            //序号
            foreach (RichTextBox v in col0)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = ""; }
            }
            //步骤名
            foreach (RichTextBox v in col1)
            {
                if (v.Text.Trim().Equals("")) { v.Text = 0 + ""; }
            }
            //标准值
            foreach (RichTextBox v in col2)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }
            //测量值
            foreach (RichTextBox v in col3)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }
            //上百分
            foreach (RichTextBox v in col4)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }
            //下百分
            foreach (RichTextBox v in col5)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }
            //高点
            foreach (RichTextBox v in col6)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }
            //低点
            foreach (RichTextBox v in col7)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }
            //延时
            foreach (RichTextBox v in col8)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }
            //模式
            foreach (RichTextBox v in col9)
            {
                if (v.Text.Trim().Equals("")) { v.Text = 0 + ""; }
            }
            //比例k
            foreach (RichTextBox v in col10)
            {
                if (v.Text.Trim().Equals("255")) { v.Text = 255 + ""; }
            }
            //迁移b
            foreach (RichTextBox v in col11)
            {
                if (v.Text.Trim().Equals("255")) { v.Text = 255 + ""; }
            }
            //上百分2
            foreach (RichTextBox v in col12)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }
            //下百分2
            foreach (RichTextBox v in col13)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }
            //高点2
            foreach (RichTextBox v in col14)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }
            //低点2
            foreach (RichTextBox v in col15)
            {
                if (v.Text.Trim().Equals("65535")) { v.Text = 65535 + ""; }
            }

        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            if (tabControl1.SelectedIndex == 1)
            {
                string dataPath = System.Environment.CurrentDirectory;
                if (!Directory.Exists(dataPath))
                {
                    Directory.CreateDirectory(dataPath);
                }
                FileInfo p = new FileInfo(dataPath + "\\config.txt");
                if (!p.Exists)
                {
                    p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
                }
                StreamReader reader = new StreamReader(p.FullName);
                
                port = reader.ReadLine();
                baudrate = reader.ReadLine();

                if(sp.IsOpen){
                    sp.Close();
                }
                sp.PortName = port;
                sp.BaudRate = Convert.ToInt32(baudrate);

                reader.Close();

                try {
                    sp.Open();
                    msg.Text = "串口" + sp.PortName + "可用";
                    sp.Close();
                }
                catch(Exception ex){
                    MessageBox.Show("串口无法打开，请检查串口设置\r\n" + ex.Message);
                    //for test
                    tabControl1.SelectedIndex = 0;
                    return;
                }

            }

        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            sp.Close();
        }

        private void tbxPage_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            double x =
            Math.Pow(16, 0) +
            Math.Pow(16, 1) +
            Math.Pow(16, 2) +
            Math.Pow(16, 3) +
            Math.Pow(16, 4) +
            Math.Pow(16, 5) +
            Math.Pow(16, 6) +
            Math.Pow(16, 7); 



            //MessageBox.Show(x*15 + "");
            double a = 4294967295;

            
        }

        private void btnSaveAndSend_Click(object sender, EventArgs e)
        {
  
            //btnSaveAndSend.Enabled = false;
            //0. 检查格式（是否数字、大小是否超范围、格式、长度等）
            checkValid();
            //1. 保存当前页的10行数据
            if(!saveCurrPage()){
                //msg.Text = "无法保存，停止发送";
                clearTable();
                showCurrPage();
                return;
            }
            

            //2. 将当前页保存的数据发给下位机,一行一行来，遇到某行不对就停止
            //先把data.txt中内容全部读出
            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\data.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            //此处先获取数据的总行数，用作进度条的显示
            int totalRows = 0;
            StreamReader readerTemp = new StreamReader(p.FullName);
            while(readerTemp.ReadLine() != null){
                totalRows++;
            }
            readerTemp.Close();

            progress.Value = 0;
            progress.ForeColor = Color.Green;
            progress.Maximum = totalRows;

            readerC = new StreamReader(p.FullName);
            
            currOp = OP_SAVE_SEND;

            string rowStr = "";
            if ((rowStr = readerC.ReadLine()) != null)
            {
                byte[] order = ConvertLineToOrderDown(rowStr);

                initPort();

                sp.Write(order, 0, order.Length);
            }
            else
            {
                readerC.Close();
                msg.Text = "命令已读取完毕";
                btnSaveAndSend.Enabled = true;
            }

            clearTable();
            showCurrPage();
        }

        /// <summary>
        /// 把一行数据转换为下发保存的命令
        /// </summary>
        private byte[] ConvertLineToOrderDown(string rowStr)
        {
            rowStr = rowStr.Substring(1).Trim();
            string[] temp = rowStr.Split(" ".ToCharArray());
            byte[] content = new byte[36];
            for (int i = 0; i < 36; i++)
            {
                content[i] = (byte)Convert.ToInt32(temp[i]);
            }

            byte[] order = new byte[42];
            order[0] = 0xFF;
            order[1] = 0x5A;
            order[2] = 0x03;
            order[3] = 0x2A;
            content.CopyTo(order, 4);
            checkSeqBeforeSend(order); //校验之前就去掉异常序号

            byte[] orderForCheck = new byte[40];
            orderForCheck[0] = 0xFF;
            orderForCheck[1] = 0x5A;
            orderForCheck[2] = 0x03;
            orderForCheck[3] = 0x2A;
            content.CopyTo(orderForCheck, 4);
            checkSeqBeforeSend(orderForCheck); //校验之前就去掉异常序号
            byte[] crcResult = CRC16MODBUS(orderForCheck, (byte)orderForCheck.Length);
            checkMaster = crcResult; //此处不用

            order[40] = crcResult[0]; //校验L
            order[41] = crcResult[1]; //校验H

            string str = "";
            foreach (byte i in order)
            {
                str += i.ToString("x2").ToUpper() + " ";
            }

            rtbxTest.Text += str + "\r\n";

            return order;
        }

        /// <summary>
        /// 分二十次逐页获取数据
        /// </summary>
        private void btnReqData_Click(object sender, EventArgs e)
        {
            clearTable();
            tbxPage.Text = null;
            progress.ForeColor = SystemColors.Highlight;
            progress.Value = 0;

            page = 1;
            transGoOn = true;

            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\data.txt");
            if (p.Exists)
            {
                try
                {
                    p.Delete();
                }catch(Exception ex){
                    readerC.Close();
                    p.Delete();
                }
            }

            initPort();
       
            currOp = OP_REQ;
            Thread.Sleep(500);
            sendOrderReq();

        }

        /// <summary>
        /// Good ! byte版本
        /// </summary>
        public byte[] CRC16MODBUS(byte[] dataBuff, byte dataLen)
        {
            Console.WriteLine("CRC16MODBUS长度：" + dataLen);

            byte CRC16High = 0;
            byte CRC16Low = 0;

            int CRCResult = 0xFFFF;
            for (int i = 0; i < dataLen; i++)
            {
                CRCResult = CRCResult ^ dataBuff[i];
                for (int j = 0; j < 8; j++)
                {
                    if ((CRCResult & 1) == 1)
                        CRCResult = (CRCResult >> 1) ^ 0xA001;
                    else
                        CRCResult >>= 1;
                }
            }
            CRC16High = Convert.ToByte(CRCResult & 0xff);
            CRC16Low = Convert.ToByte(CRCResult >> 8);

            byte[] result = new byte[2];
            result[0] = CRC16High;
            result[1] = CRC16Low;
            return result;

        }

        /// <summary>
        /// Good ! dataLen为int版本, 原来的先不动
        /// </summary>
        public byte[] CRC16MODBUS(byte[] dataBuff, int dataLen)
        {
            Console.WriteLine("CRC16MODBUS长度：" + dataLen);

            byte CRC16High = 0;
            byte CRC16Low = 0;

            int CRCResult = 0xFFFF;
            for (int i = 0; i < dataLen; i++)
            {
                CRCResult = CRCResult ^ dataBuff[i];
                for (int j = 0; j < 8; j++)
                {
                    if ((CRCResult & 1) == 1)
                        CRCResult = (CRCResult >> 1) ^ 0xA001;
                    else
                        CRCResult >>= 1;
                }
            }
            CRC16High = Convert.ToByte(CRCResult & 0xff);
            CRC16Low = Convert.ToByte(CRCResult >> 8);

            byte[] result = new byte[2];
            result[0] = CRC16High;
            result[1] = CRC16Low;
            return result;

        }

        private void btnTestData_Click_1(object sender, EventArgs e)
        {
            clearTable();
            tbxPage.Text = null;
            progress.ForeColor = SystemColors.Highlight;
            progress.Value = 0;

            page = 1;
            transGoOnTest = true;

            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\data.txt");
            if (p.Exists)
            {
                p.Delete();
            }

            line = 0;
            initPort();
            currOp = OP_TEST;
            Thread.Sleep(500);
            sendOrderTest();
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
             initPort();
             Thread.Sleep(500);
             sendOrderReset();
        }

        private void btnGetSysParam_Click(object sender, EventArgs e)
        {
             clearSysParam();
             resultLengthParam = 0;
             initPort();
             currOp = OP_GET_SYS_PARAM;
             Thread.Sleep(500);
             sendOrderGetSysParam();
        }

        private void btnSaveSysParam_Click(object sender, EventArgs e)
        {
            if (checkValidSysparam())
            {

                int[] sysParam = new int[15];

                sysParam[0] = Convert.ToInt32(tvSpVersion.Text);
                sysParam[1] = Convert.ToInt32(tvSp485.Text);
                sysParam[2] = Convert.ToInt32(tvSpDS.Text) / 256;
                sysParam[3] = Convert.ToInt32(tvSpDS.Text) % 256;
                sysParam[4] = Convert.ToInt32(tvSpOSFZ.Text) / 256; ;
                sysParam[5] = Convert.ToInt32(tvSpOSFZ.Text) % 256;
                sysParam[6] = Convert.ToInt32(tvSpZDFZ.Text) / 256;
                sysParam[7] = Convert.ToInt32(tvSpZDFZ.Text) % 256;
                sysParam[8] = Convert.ToInt32(tvSpJT.Text);

                byte[] temp = ASCIIEncoding.ASCII.GetBytes(tvSpJZH.Text.ToCharArray());

                sysParam[9] = temp[0];
                sysParam[10] = temp[1];
                sysParam[11] = temp[2];
                sysParam[12] = temp[3];
                sysParam[13] = temp[4];
                sysParam[14] = temp[5];

                string sysParamData = "";
                for (int i = 0; i < 15; i++)
                {
                    sysParamData += sysParam[i] + " ";
                }

                //先把data.txt中内容全部读出
                string dataPath = System.Environment.CurrentDirectory;
                if (!Directory.Exists(dataPath))
                {
                    Directory.CreateDirectory(dataPath);
                }
                FileInfo p = new FileInfo(dataPath + "\\sysparam.txt");
                if (!p.Exists)
                {
                    p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
                }
                StreamWriter writer = new StreamWriter(p.FullName);
                writer.WriteLine(sysParamData);
                writer.Close();

                byte[] data = new byte[15];
                for (int i = 0; i < 15; i++)
                {
                    data[i] = (byte)sysParam[i];
                }

                currOp = OP_SET_SYS_PARAM;
                //btnSaveSysParam.Enabled = false;

                initPort();
                Thread.Sleep(500);
                sendOrderSetSysParam(data);
            }
          
        }

        /// <summary>
        /// 获取系统参数
        /// </summary>
        private int[] getSysParam()
        {
            //先把data.txt中内容全部读出
            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
           
            FileInfo p = new FileInfo(dataPath + "\\sysparam.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamReader reader = new StreamReader(p.FullName);
            string str = reader.ReadToEnd().Trim();
            reader.Close();

            string[] temp = str.Split(" ".ToCharArray());
        
            int[] sysparam = new int[temp.Length];
            for (int i = 0; i < temp.Length; i++ )
            {
                sysparam[i] = int.Parse(temp[i]);
            }

            return sysparam;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //先把data.txt中内容全部读出
            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            FileInfo p = new FileInfo(dataPath + "\\data.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamReader reader = new StreamReader(p.FullName);
            string strData = reader.ReadToEnd();
            reader.Close();

            p = new FileInfo(dataPath + "\\sysparam.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            reader = new StreamReader(p.FullName);
            string strSysParam = reader.ReadToEnd();
            reader.Close();

            p = new FileInfo(dataPath + "\\study1.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            reader = new StreamReader(p.FullName);
            string study1 = reader.ReadToEnd().Trim();
            reader.Close();

            p = new FileInfo(dataPath + "\\study2.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            reader = new StreamReader(p.FullName);
            string study2 = reader.ReadToEnd().Trim();
            reader.Close();

            string export = strData + "--" + strSysParam + "--" + study1 + "--" + study2;

            p = new FileInfo("c:\\export.txt");
            if (!p.Exists)
            {
                p.Create().Close(); //【写blog 这里必须加Close() ，否则下一步会提示被占用】
            }
            StreamWriter writer = new StreamWriter(p.FullName);
            writer.Write(export);
            writer.Close();

            msg.Text = "数据已导出至 C:\\export.txt " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }

        private void btnImport_Click(object sender, EventArgs e)
        {

            clearTable();
            clearSysParam();

            StreamReader reader = null;
            StreamWriter writer = null;

            fileDlg = new OpenFileDialog();
            fileDlg.InitialDirectory = "c://";
            fileDlg.Filter = "txt files (*.txt)|*.txt";
            fileDlg.FilterIndex = 1;
            fileDlg.RestoreDirectory = true;
            if (fileDlg.ShowDialog() == DialogResult.OK)
            {
                string filePath = fileDlg.FileName;

                try
                {
                    reader = new StreamReader(filePath);
                    string temp = reader.ReadToEnd();
                    reader.Close();

                    string[] content = temp.Split("--".ToCharArray());

                    string dataPath = System.Environment.CurrentDirectory;
                    if (!Directory.Exists(dataPath))
                    {
                        Directory.CreateDirectory(dataPath);
                    }
                    FileInfo p = new FileInfo(dataPath + "\\data.txt");
                    if (!p.Exists)
                    {
                        p.Create().Close();
                    }

                    writer = new StreamWriter(p.FullName);
                    writer.Write(content[0]);
                    writer.Close();

                    p = new FileInfo(dataPath + "\\sysparam.txt");
                    if (!p.Exists)
                    {
                        p.Create().Close();
                    }

                    writer = new StreamWriter(p.FullName);
                    writer.Write(content[2]);
                    writer.Close();

                    p = new FileInfo(dataPath + "\\study1.txt");
                    if (!p.Exists)
                    {
                        p.Create().Close();
                    }

                    writer = new StreamWriter(p.FullName);
                    writer.Write(content[4]);
                    writer.Close();

                    p = new FileInfo(dataPath + "\\study2.txt");
                    if (!p.Exists)
                    {
                        p.Create().Close();
                    }

                    writer = new StreamWriter(p.FullName);
                    writer.Write(content[6]);
                    writer.Close();

                    msg.Text = "数据导入完成 " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    btnSaveAndSend.Enabled = true;
                    btnSaveSysParam.Enabled = true;

                    //显示数据
                    page = 1;
                    tbxPage.Text = page + "";
                    btnPageNext.Enabled = true;
                    btnPagePrevious.Enabled = true;

                    showCurrPage();
                    showSysParam();

                }catch(Exception ex){
                    if(reader != null){
                        reader.Close();
                    }
                    if(writer != null){
                        writer.Close();
                    }

                    MessageBox.Show("出错：" + ex.Message);
                }  
            }
            
        }

        private void btnFindPoint_Click(object sender, EventArgs e)
        {
            currOp = OP_ZD;
            lenToRead = 0;
            initPort();
            msg.Text = null;
            msg2.Text = null;
            cbxOrder.Text = null;
            sendOrderZd();
        }

        private void btnStudy1_Click(object sender, EventArgs e)
        {
            //for test
            //showStudy1Result();
            //....
            //return;

            currOp = OP_STUDY1;
            resultLengthStudy1 = 0;
            initPort();
            msg.Text = null;
            msg2.Text = null;
            cbxOrder.Text = null;
            sendOrderStudy1();
        }

        //for test
        public void readStudy1Data(int[] temp)
        {
            //先把data.txt中内容全部读出
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
            StreamReader reader = new StreamReader(p.FullName);
            string str = reader.ReadToEnd().Trim();
            reader.Close();

            string[] tempStrs = str.Split(" ".ToCharArray());

            int[] tempData = new int[tempStrs.Length];
            for (int i = 0; i < tempStrs.Length; i++)
            {
                temp[i] = int.Parse(tempStrs[i]);
                Console.Write(temp[i] + " ");
            }

           // return sysparam;
        }

        /// <summary>
        /// 展示学习1的结果
        /// </summary>
        public void showStudy1Result()
        {
            //获取系统参数，要判断点数
            int pointNum = 256;
            int[] sysparam = getSysParam();
            if (sysparam.Length >= 4)
            {
                pointNum = sysparam[2] * 256 + sysparam[3];
            }

            //int[] temp = new int[256]; //262-6
            //暂时注释 for test
            int[] temp = new int[256]; //262-6
            for(int i = 4; i < 260; i ++){ //从第4个开始，到（262-2）个
                temp[i-4] = resultStudy1[i]; //******【resultStudy[] 是关键】
            }
            
            //for test
            //readStudy1Data(temp);

            Console.WriteLine("temp[0]: " + temp[0]);
            //展示temp
          
            formStudyResult.lv.Items.Clear();
            Console.WriteLine("开始组织");
            //计数器

            int startIndex = 0; //本次循环的起始位置
            bool nextIndexReady = false;
            int nextIndex = 0; //下次循环的起始位置

            string passIndexs = ""; //存放已经处理过的index，后面遇到就跳过

          //while (startIndex < 255)
            while (startIndex < (pointNum - 1))
            {

                nextIndexReady = false;

                //第一次的时候不执行？
                startIndex = nextIndex;

                string line = "( ";
                line += (startIndex + 1) + ", ";

                Console.Write("本次 startIndex: " + startIndex + "  ");

                for (int i = startIndex; i < temp.Length; i++)
                {
                    if (temp[i] == temp[startIndex])
                    {
                        if(i != startIndex){
                            line += (i + 1) + ", "; //？ 同时，这个i要去除，下次循环的时候遇到就跳过
                            passIndexs += i + "_";
                        }
                    }
                    else
                    {
                        if (!nextIndexReady)
                        {
                            nextIndex = (i);
                            nextIndexReady = true;
                            Console.WriteLine(" nextIndex: " + nextIndex + "  ");
                        }
                        
                    }
                }

            
                line = line.Substring(0, line.Length - 2);

                line += " )";
                Console.WriteLine(" line: " + line);
                if (!passIndexs.Contains(startIndex + "") || startIndex == 0)
                {
                    formStudyResult.lv.Items.Add(line);
                }
               

            }

            for (int i = 0; i < formStudyResult.lv.Items.Count; i ++ )
            {

                if (formStudyResult.lv.Items[i].ToString().Contains(","))
                {
                    formStudyResult.lv.Items[i] = "  # " + formStudyResult.lv.Items[i];
                }
                else
                {
                    formStudyResult.lv.Items[i] = "     " + formStudyResult.lv.Items[i];
                }
            }

            formStudyGetResult.lblMsg2.Text = null;
            formStudyResult.lblStudyResultType.Text = "学习1";
            
            formStudyResult.Show();

        }

        /// <summary>
        /// 展示学习2的结果
        /// </summary>
        public void showStudy2Result()
        {
            //获取系统参数，要判断点数
            int pointNum = 256;
            int[] sysparam = getSysParam();
            if (sysparam.Length >= 4)
            {
                pointNum = sysparam[2] * 256 + sysparam[3];
            }

            int[] temp = new int[256]; //262-6
            for (int i = 4; i < 260; i++)
            { //从第4个开始，到（262-2）个
                temp[i - 4] = resultStudy2[i];
                //Console.WriteLine((i-4) + ": " + temp[i-4]);

            }

            Console.WriteLine("temp[0]: " + temp[0]);
            //展示temp

            formStudyResult.lv.Items.Clear();
            Console.WriteLine("开始组织");
            //计数器

            int startIndex = 0; //本次循环的起始位置
            bool nextIndexReady = false;
            int nextIndex = 0; //下次循环的起始位置

            string passIndexs = ""; //存放已经处理过的index，后面遇到就跳过

            //while (startIndex < 255)
            while (startIndex < (pointNum - 1))
            {

                nextIndexReady = false;

                //第一次的时候不执行？
                startIndex = nextIndex;

                string line = "( ";
                line += (startIndex + 1) + ", ";

                Console.Write("本次 startIndex: " + startIndex + "  ");

                for (int i = startIndex; i < temp.Length; i++)
                {
                    if (temp[i] == temp[startIndex])
                    {
                        if (i != startIndex && !passIndexs.Contains(i + ""))
                        {
                            line += (i + 1) + ", "; //？ 同时，这个i要去除，下次循环的时候遇到就跳过
                            passIndexs += i + "_";
                        }
                    }
                    else
                    {
                        if (!nextIndexReady)
                        {
                            nextIndex = (i);
                            nextIndexReady = true;
                            Console.WriteLine(" nextIndex: " + nextIndex + "  ");
                        }

                    }
                }


                line = line.Substring(0, line.Length - 2);

                line += " )";
                formStudyResult.lv.Items.Add(line);

            }

            for (int i = 0; i < formStudyResult.lv.Items.Count; i++)
            {
              
                if (formStudyResult.lv.Items[i].ToString().Contains(","))
                {
                    formStudyResult.lv.Items[i] = "  # " + formStudyResult.lv.Items[i];
                }
                else
                {
                    formStudyResult.lv.Items[i] = "     " + formStudyResult.lv.Items[i];
                }

            }

            formStudyGetResult.lblMsg2.Text = null;
            formStudyResult.lblStudyResultType.Text = "学习2";
            formStudyResult.Show();

        }

        /// <summary>
        /// 保存获取学习1的结果
        /// </summary>
        private void saveStudy1GetResult()
        {
            string content = "";
            int[] temp = new int[256]; //262-6
            for (int i = 4; i < 260; i++)
            { //从第4个开始，到（262-2）个
                content += resultStudy1Get[i] + " ";
            }
            Console.WriteLine("study1: " + content);
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
            writer.Close();
            Console.WriteLine("study1: Save Ok");
            msg.Text = "学习1内容保存成功 " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }

        /// <summary>
        /// 保存获取学习2的结果
        /// </summary>
        private void saveStudy2GetResult()
        {
            string content = "";
            int[] temp = new int[256]; //262-6
            for (int i = 4; i < 260; i++)
            { //从第4个开始，到（262-2）个
                content += resultStudy2Get[i] + " ";
            }
            Console.WriteLine("study2: " + content);
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
            writer.Close();
            Console.WriteLine("study2: Save Ok");
            msg.Text = "学习2内容保存成功 " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }

        /// <summary>
        /// 展示获取学习1的结果
        /// </summary>
        private void showStudy1GetResult()
        {
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
            StreamReader reader = new StreamReader(p.FullName);
            string study1 = reader.ReadToEnd();
            reader.Close();

            string[] data = study1.Trim().Split(" ".ToCharArray());

            for (int i = 0; i < formStudyGetResult.dgv.Rows.Count; i ++ )
            {
                 if(data.Length > (i * 16)){
                     for (int k = 0; k < 16; k++)
                     {
                         if (data.Length > (i * 16 + k))
                         {
                             formStudyGetResult.dgv.Rows[i].Cells[k].Value = data[i * 16 + k];
                         }
                     }
                 }
            }

            formStudyGetResult.studyType = 1;
            formStudyGetResult.lblStudyResultType.Text = "学习1 内容";
            formStudyGetResult.Show();

        }

        /// <summary>
        /// 展示获取学习2的结果
        /// </summary>
        private void showStudy2GetResult()
        {
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
            StreamReader reader = new StreamReader(p.FullName);
            string study1 = reader.ReadToEnd();
            reader.Close();

            string[] data = study1.Trim().Split(" ".ToCharArray());

            for (int i = 0; i < formStudyGetResult.dgv.Rows.Count; i++)
            {
                if (data.Length > (i * 16))
                {
                    for (int k = 0; k < 16; k++)
                    {
                        if (data.Length > (i * 16 + k))
                        {
                            formStudyGetResult.dgv.Rows[i].Cells[k].Value = data[i * 16 + k];
                        }
                    }
                }
            }

            formStudyGetResult.studyType = 2;
            formStudyGetResult.lblStudyResultType.Text = "学习2 内容";
            formStudyGetResult.Show();

        }

        private void btnStudy2_Click(object sender, EventArgs e)
        {
            currOp = OP_STUDY2;
            resultLengthStudy2 = 0;
            initPort();
            msg.Text = null;
            msg2.Text = null;
            cbxOrder.Text = null;
            sendOrderStudy2();
        }

        private void btnStudy1Get_Click(object sender, EventArgs e)
        {
            currOp = OP_STUDY1_GET;
            resultLengthStudy1Get = 0;
            initPort();
            msg.Text = null;
            msg2.Text = null;
            cbxOrder.Text = null;
            sendOrderStudy1Get();
        }

        private void btnStudy2Get_Click(object sender, EventArgs e)
        {
            currOp = OP_STUDY2_GET;
            resultLengthStudy2Get = 0;
            initPort();
            msg.Text = null;
            msg2.Text = null;
            cbxOrder.Text = null;
            sendOrderStudy2Get();
        }

        private void btnStudy1Edit_Click(object sender, EventArgs e)
        {
            showStudy1GetResult();
        }

        private void btnStudy2Edit_Click(object sender, EventArgs e)
        {
            showStudy2GetResult();
        }

        /// <summary>
        /// 下位机单机测试，软件只要接收即可
        /// </summary>
        private void btnTestData2_Click(object sender, EventArgs e)
        {
            clearTable();
            tbxPage.Text = null;
            progress.ForeColor = SystemColors.Highlight;
            progress.Value = 0;

            page = 1;
            transGoOnTest = true;

            string dataPath = System.Environment.CurrentDirectory;
            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }
            /*
            FileInfo p = new FileInfo(dataPath + "\\data.txt");
            if (p.Exists)
            {
                p.Delete();
            }
            */
            line = 0;
            initPort();
            currOp = OP_TEST;

        }
       
    }
}
