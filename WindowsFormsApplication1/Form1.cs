using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Timers;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            aTimer = new System.Timers.Timer(100);
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Enabled = true;
        }

        public struct ResolvedNetObj
        {   
            public string NetObjName;
            public List<List<string>> ResolvedObj;
			public bool IsObjectNegated;
            public string FDQN;

            public ResolvedNetObj(string p)
            {
                NetObjName = p;
                ResolvedObj = ResolveNetObj(p);
                IsObjectNegated = false;
                FDQN = "";
            }

            public ResolvedNetObj(string p, bool n)
            {
                NetObjName = p;
                ResolvedObj = ResolveNetObj(p);
                IsObjectNegated = n;
                FDQN = "";
            }

        }

        public static List<List<string>> CurrentErrorObjResolved = new List<List<string>>();

        public static List<string> AllWSNames = new List<string>();

        public static List<List<string>> Rules = new List<List<string>>();
        public static List<List<IEnumerable<string>>> NetworkObj = new List<List<IEnumerable<string>>>();
        public static List<List<string>> Services = new List<List<string>>();
        public static List<ResolvedNetObj> ToNodeList = new List<ResolvedNetObj>();
        public static System.Timers.Timer aTimer;
        
        public static string returnNet(string netInPointDelimeter)
        {
            string result = "";
            switch (netInPointDelimeter)
            {
                case "255.255.255.255":
                    result = "";
                    break;
                case "255.255.255.254":
                    result = "/31";
                    break;
                case "255.255.255.252":
                    result = "/30";
                    break;
                case "255.255.255.248":
                    result = "/29";
                    break;
                case "255.255.255.240":
                    result = "/28";
                    break;
                case "255.255.255.224":
                    result = "/27";
                    break;
                case "255.255.255.192":
                    result = "/26";
                    break;
                case "255.255.255.128":
                    result = "/25";
                    break;
                case "255.255.255.0":
                    result = "/24";
                    break;
                case "255.255.254.0":
                    result = "/23";
                    break;
                case "255.255.252.0":
                    result = "/22";
                    break;
                case "255.255.248.0":
                    result = "/21";
                    break;
                case "255.255.240.0":
                    result = "/20";
                    break;
                case "255.255.224.0":
                    result = "/19";
                    break;
                case "255.255.192.0":
                    result = "/18";
                    break;
                case "255.255.128.0":
                    result = "/17";
                    break;
                case "255.255.0.0":
                    result = "/16";
                    break;
                case "255.254.0.0":
                    result = "/15";
                    break;
                case "255.252.0.0":
                    result = "/14";
                    break;
                case "255.248.0.0":
                    result = "/13";
                    break;
                case "255.240.0.0":
                    result = "/12";
                    break;
                case "255.224.0.0":
                    result = "/11";
                    break;
                case "255.192.0.0":
                    result = "/10";
                    break;
                case "255.128.0.0":
                    result = "/9";
                    break;
                case "255.0.0.0":
                    result = "/8";
                    break;
                case "254.0.0.0":
                    result = "/7";
                    break;
                case "252.0.0.0":
                    result = "/6";
                    break;
                case "248.0.0.0":
                    result = "/5";
                    break;
                case "240.0.0.0":
                    result = "/4";
                    break;
                case "224.0.0.0":
                    result = "/3";
                    break;
                case "192.0.0.0":
                    result = "/2";
                    break;
                case "128.0.0.0":
                    result = "/1";
                    break;
                case "0.0.0.0":
                    result = "/0";
                    break;
                default:
                    result = "error";
                    break;
            }
            return result;
        }

        public static List<List<string>> ResolveNetObj(string NetObj)
        {
            List<List<string>> result = new List<List<string>>();
            string[] split = { "\r\n", "\r", "\n" };

            if (NetObj == "Any")
            {
                result.Add(new List<string> { "Any", "0.0.0.0", "/0"});
                return result;
            }

            try
            {
                List<IEnumerable<string>> GetRowWithResolvedInfo = NetworkObj.Single(OneRow => OneRow[0].ToList()[0].Trim().ToLower().Equals(NetObj.Trim().ToLower()));

                string tmp = GetRowWithResolvedInfo[1].ToList()[0];

                switch (GetRowWithResolvedInfo[1].ToList()[0])
                {
                    case "Host Node":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0], "" });
                        break;
                    case "Dynamic Object":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[0].ToList()[0], "" });
                        break;
                    case "Network":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0], returnNet(GetRowWithResolvedInfo[3].ToList()[0]) });
                        break;
                    case "Unknown Type":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0].Split(split, StringSplitOptions.RemoveEmptyEntries)[0].ToString(), "" });
                        break;
                    case "Check Point Host":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0], "" });
                        break;
                    case "Check Point Gateway":
                        string tmp3 = GetRowWithResolvedInfo[2].ToList()[0].Split(split, StringSplitOptions.RemoveEmptyEntries)[0].ToString();
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], tmp3, "" });
                        break;
                    case "Group":
                        foreach (string OneOfGroupNetObj in GetRowWithResolvedInfo[6].Distinct())
                        {
                            if (OneOfGroupNetObj == "")
                                continue;
                            List<List<string>> tmpresult = ResolveNetObj(OneOfGroupNetObj);
                            if (tmpresult.Count == 0)
                                continue;
                            result.AddRange(tmpresult);
                        }
                        break;
                    case "Cluster Member":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0].Split(split, StringSplitOptions.RemoveEmptyEntries)[0].ToString(), "" });
                        break;
                    case "Gateway Cluster":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0].Split(split, StringSplitOptions.RemoveEmptyEntries)[0].ToString(), "" });
                        break;
                    case "Interoperable Device":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0], "" });
                        break;
                    case "Address Range":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0], "" });
                        break;
                    case "Safe@Gateway":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0].Split(split, StringSplitOptions.RemoveEmptyEntries)[0].ToString(), "" });
                        break;
                    case "Connectra":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0].Split(split, StringSplitOptions.RemoveEmptyEntries)[0].ToString(), "" });
                        break;
                    default:
                        result.Add(new List<string> { "Неизвестный тип", GetRowWithResolvedInfo[1].ToList()[0], "Error" });
                        CurrentErrorObjResolved.Add(new List<string> { "Неизвестный тип", GetRowWithResolvedInfo[1].ToList()[0], "Error" });
                        break;
                }
                
            }
            catch (Exception resExc)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, resExc.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, resExc.Source);
                result.Add(new List<string> { NetObj + " Не найден", "", "" });
                CurrentErrorObjResolved.Add(new List<string> { NetObj + " Не найден", "", "" });
            }            
            return result;

        }

        public  List<List<string>> ResolvePorts(string NetObj)
        {
            List<List<string>> result = new List<List<string>>();
            string[] split = { "\r\n", "\r", "\n" };

            if (NetObj == "Any")
            {
                result.Add(new List<string> {"Any", "Any", "" });
                return result;
            }

            try
            {
                List<string> GetRowWithResolvedInfo = Services.Single(OneRow => OneRow[0].Equals(NetObj));

                switch (GetRowWithResolvedInfo[1].ToLower())
                {
                    case "other":
                        if (GetRowWithResolvedInfo[2] == "-1")
                            result.Add(new List<string> { GetRowWithResolvedInfo[0].Trim(), "Динамические порты", "" });
                        else
                            result.Add(new List<string> { GetRowWithResolvedInfo[0].Trim(), "Протокол #" + GetRowWithResolvedInfo[2].Trim(), "" });
                        break;
                    case "dcerpc":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].Trim(), GetRowWithResolvedInfo[0].Trim(), "DCE-RPC" });
                        break;
                    case "tcp":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].Trim(), "Tcp", GetRowWithResolvedInfo[2].Trim() });
                        break;
                    case "udp":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].Trim(), "Udp", GetRowWithResolvedInfo[2].Trim() });
                        break;
                    case "rpc":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0], "RPC", GetRowWithResolvedInfo[2].Trim() });
                        break;
                    case "icmp":
                        result.Add(new List<string> { "Icmp", "Icmp", "" });
                        break;
                    case "group":
                        foreach (string OneOfGroupPortObj in GetRowWithResolvedInfo[6].Split(split, StringSplitOptions.RemoveEmptyEntries).Distinct().ToList())
                        {
                            if (OneOfGroupPortObj == "")
                                continue;
                            List<List<string>> tmpresult = ResolvePorts(OneOfGroupPortObj);
                            if (tmpresult.Count == 0)
                                continue;
                            result.AddRange(tmpresult);
                        }
                        break;
                    default:
                        result.Add(new List<string> { "Error", "Error", "Error" });
                        break;
                }

            }
            catch (Exception resExc)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, resExc.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, resExc.Source);
                result.Add(new List<string> { "Exception Error", resExc.Message, resExc.Source });
            }
            return result;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            numericUpDown1.Enabled = false;
            numericUpDown2.Enabled = false;
            checkBox2.Enabled = false;
            checkBox1.Enabled = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            { 
                ///
                /// Копирую таблицы со структурой
                /// Правила - двумерная структура
                /// Сетевые объекты - трёхмерная (двумерная таблица с объектом-списком)
                /// Сервисы - двумерная таблица
                ///

                Rulesd_Delegate d = null;
                d = new Rulesd_Delegate(Rulesd);
                IAsyncResult R = null;
                R = d.BeginInvoke(openFileDialog1.FileName, null, null);                
                Rules = d.EndInvoke(R);

                /*
                 * Это N правил + 1 отрезаем кончающие правила 
                 * Далее начальные
                 */

                if (checkBox2.Checked &&
                    numericUpDown1.Value < numericUpDown2.Value &&
                    numericUpDown2.Value < Rules.Count)
                {
                    Rules.RemoveRange(Convert.ToInt32(numericUpDown2.Value) + 1, Rules.Count - Convert.ToInt32(numericUpDown2.Value) - 1);
                    Rules.RemoveRange(0, Convert.ToInt32(numericUpDown1.Value));
                }
                else
                    if (checkBox2.Checked)
                    {
                        MessageBox.Show("Проверьте правильно ли Вы выставили значения диапазона интересующих правил");
                        return;
                    }

                AddTextToStatus("Скопировал Таблицу правил");

                NetworkObjd_Delegate ND = null;
                ND = new NetworkObjd_Delegate(NetworkObjd);
                IAsyncResult NR = null;
                NR = ND.BeginInvoke(openFileDialog1.FileName, null, null);
                NetworkObj = ND.EndInvoke(NR);

                AddTextToStatus("Скопировал Таблицу Сетевых имён");                

                for (int k = 0; k < NetworkObj.Count; k++)
                {
                    for (int l = 0; l < NetworkObj[l].Count; l++)
                    {
                        NetworkObj[k][l] = NetworkObj[k][l].Distinct();
                    }
                }

                Servicesd_Delegate SD = null;
                SD = new Servicesd_Delegate(Servicesd);
                IAsyncResult SR = null;
                SR = SD.BeginInvoke(openFileDialog1.FileName, null, null);
                Services = SD.EndInvoke(SR);

                AddTextToStatus("Скопировал Таблицу сервисов");

                ///
                /// Открываем книгу Excel. 
                /// 
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Workbook WB; 
                Worksheet WS;
                WB = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                WS = (Worksheet)WB.Sheets[1];
                int i = 1;

                AddTextToStatus("Открыл ExCel");

                progressBar1.Maximum = Rules.Count;
                progressBar1.Value = 0;

                try
                {
                    foreach (List<string> row in Rules)
                    {
                        if (i == 1)
                        {
                            WS.Cells[i, 1].Value = "Сервис";
                            WS.Cells[i, 2].Value = "Сегмент источника";
                            WS.Cells[i, 3].Value = "IP Источника";
                            WS.Cells[i, 4].Value = "Сегмент назначения";
                            WS.Cells[i, 5].Value = "IP назначения";
                            WS.Cells[i, 6].Value = "Порты";
                            WS.Cells[i, 7].Value = "Решение";
                            WS.Cells[i, 8].Value = "Установлено";
                            WS.Cells[i, 9].Value = "Время";
                            WS.Cells[i, 10].Value = "Комментарий";
                            i++;
                            continue;
                        }

                        string[] split = { "\r\n", "\r", "\n" };

                        ///
                        /// Это для Source
                        /// Есть структура данных ResolvedNetObj, которая содержит
                        /// отрезолвленные члены группового объекта
                        /// Создадим список из структуры ResolvedNetObj который будет содержать
                        /// все объекты ячейки источник из политики
                        ///
                        
                        List<string> CurrentSourceNetObj = row[2].Split(split, StringSplitOptions.RemoveEmptyEntries).ToList();
                        List<ResolvedNetObj> CurrentSrc = GetResolvedNetObj(CurrentSourceNetObj);
                        AddTextToStatus("Отрезолвил Объекты Источники для " + i.ToString() + " правила");

                        foreach (ResolvedNetObj oneNetObj in CurrentSrc)
                            if (oneNetObj.ResolvedObj.Count > 1)
                                foreach (List<string> OneResNetObj in oneNetObj.ResolvedObj)
                                    if (ToNodeList.Where(n => n.NetObjName == OneResNetObj[0]).Count() == 0)
                                    {
                                        List<ResolvedNetObj> tmpone = GetResolvedNetObj(new List<string> { OneResNetObj[0] });
                                        ToNodeList.AddRange(tmpone);
                                        if (ToNodeList.Where(n => n.NetObjName == oneNetObj.NetObjName).Count() == 0)
                                            ToNodeList.Add(oneNetObj);
                                    }
                                    else
                                        continue;
                            else
                                if (ToNodeList.Where(n => n.NetObjName == oneNetObj.NetObjName).Count() == 0)                                
                                    ToNodeList.Add(oneNetObj);
                       
                        ///
                        /// Это для Destination
                        ///

                        List<string> CurrentDestNetObj = row[3].Split(split, StringSplitOptions.RemoveEmptyEntries).ToList();
                        List<ResolvedNetObj> CurrentDst = GetResolvedNetObj(CurrentDestNetObj);
                        AddTextToStatus("Отрезолвил Объекты Назначения для " + i.ToString() + " правила");

                        foreach (ResolvedNetObj oneNetObj in CurrentDst)
                            if (oneNetObj.ResolvedObj.Count > 1)
                                foreach (List<string> OneResNetObj in oneNetObj.ResolvedObj)
                                    if (ToNodeList.Where(n => n.NetObjName == OneResNetObj[0]).Count() == 0)
                                    {
                                        List<string> one = new List<string> { OneResNetObj[0] };
                                        List<ResolvedNetObj> tmpone = GetResolvedNetObj(one);
                                        ToNodeList.AddRange(tmpone);
										if (ToNodeList.Where(n => n.NetObjName == oneNetObj.NetObjName).Count() == 0)
											ToNodeList.Add(oneNetObj);
                                    }
                                    else
                                        continue;
                            else
                                if (ToNodeList.Where(n => n.NetObjName == oneNetObj.NetObjName).Count() == 0)
                                    ToNodeList.Add(oneNetObj);

                        ///
                        /// Это для портов
                        ///

                        List<string> CurrenPortObj = row[5].Split(split, StringSplitOptions.RemoveEmptyEntries).ToList();
                        List<List<string>> CurrentPortObjResolved = new List<List<string>>();

                        progressBar2.Maximum = CurrenPortObj.Count;
                        progressBar2.Value = 0;

                        foreach (string OnePortObj in CurrenPortObj)
                        {
                            string s = OnePortObj.Replace("Not ", "").Trim();
                            List<List<string>> tmp = ResolvePorts(s);
                            if (OnePortObj.Contains("Not "))
                                for (int b = 0; b < tmp.Count; b++)
                                    tmp[b][0] = "Not " + tmp[b][0];

                            CurrentPortObjResolved.AddRange(tmp);

                            if (progressBar2.InvokeRequired)
                                progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Value++));
                            else
                                progressBar2.Value++;
                            System.Windows.Forms.Application.DoEvents();
                        }

                        AddTextToStatus("Отрезолвил Объекты порты для " + i.ToString() + " правила");

                        for (int a = 0 ; a < row.Count ; a++)
                        {
                            if (a == 1)
                                WS.Cells[i, 1].Value = row[a];

                            if (a == 2)                            
                                AddNetworkObj(CurrentSrc, WB, WS, i, 2);
                            

                            if (a == 3)                            
                                AddNetworkObj(CurrentDst, WB, WS, i, 4);
                            
                            if (a == 5)
                            {
                                string ThisCell = "";
                                foreach (List<string> PortTableRow in CurrentPortObjResolved)
                                    ThisCell += PortTableRow[0] + " (" + PortTableRow[1] +  "/" + PortTableRow[2] + ")" + Environment.NewLine;
                                WS.Cells[i, 6].Value = ThisCell.Replace("Icmp (Icmp/)", "ICMP")
                                    .Replace("(ALL_DCE_RPC/DCE-RPC)", "(Все DCE-RPC)")
                                    .Replace("Any (Any/)", "Any")
                                    .Replace("l_ALL_DCE_RPC(l_ALL_DCE_RPC / DCE - RPC)", "l_ALL_DCE_RPC (Все DCE-RPC)")
                                    .Trim();
                            }

                            if (a == 6)
                                WS.Cells[i, 7].Value = row[a];

                            if (a == 8)
                                WS.Cells[i, 8].Value = row[a].Replace("\r\n", "\n").Replace("\n\r", "\n").Replace("\r\r", "\n").Replace("\n\n", "\n").Trim();

                            if (a == 9)
                                WS.Cells[i, 9].Value = row[a];

                            if (a == 10)
                                WS.Cells[i, 10].Value = row[a].Replace("&nbsp;", "").Trim();

                            if (a == 0 || a == 4 || a == 7)
                                continue;
                        }
                        
                        AddTextToStatus("Записал " + i.ToString() + " правило");       
                        i++;

                        if (progressBar1.InvokeRequired)
                            progressBar1.BeginInvoke(new MethodInvoker(() => progressBar1.Value++));
                        else
                            progressBar1.Value++;
                        System.Windows.Forms.Application.DoEvents();                           
                    }

                    AddTextToStatus("Добавляю лист Node");
                    AddNodeList(WB);

					if (checkBox3.Checked)	
						AddAllNodeList(WB);
				

                    if (CurrentErrorObjResolved.Count != 0)
                    {
                        PolycyConverter.Form3 form3 = new PolycyConverter.Form3();
                        form3.ShowDialog(this);
                    }
                    else
                    {
                        PolycyConverter.Form2 form2 = new PolycyConverter.Form2();
                        form2.ShowDialog(this);
                    }
                }
                catch (Exception theException)  
                {  
                    String errorMessage;  
                    errorMessage = "Error: ";  
                    errorMessage = String.Concat(errorMessage, theException.Message);  
                    errorMessage = String.Concat(errorMessage, " Line: ");  
                    errorMessage = String.Concat(errorMessage, theException.Source);  
                    
                    MessageBox.Show(errorMessage, "Error");  

                    ObjExcel.Columns.AutoFit();
                    ObjExcel.Rows.AutoFit();

                    PolycyConverter.Form3 form3 = new PolycyConverter.Form3();                    
                    form3.ShowDialog(this);
                    return;
                } 
                finally
                {
                    WS.Activate();
                    Microsoft.Office.Interop.Excel.Range c1 = WS.Cells[1, 2];
                    Microsoft.Office.Interop.Excel.Range c2 = WS.Cells[i+1, 10];
                    Range oRange = (Microsoft.Office.Interop.Excel.Range)WS.get_Range(c1, c2);

                    oRange.Columns.AutoFit();
                    oRange.Rows.AutoFit();

                    WS.Application.ActiveWindow.SplitRow = 1;
                    WS.Application.ActiveWindow.FreezePanes = true;

                    ObjExcel.Visible = true;
                    ObjExcel.UserControl = true;

                    Rules.Clear();
                    NetworkObj.Clear();
                    Services.Clear();

                    button1.Enabled = true;
                    numericUpDown1.Enabled = true;
                    numericUpDown2.Enabled = true;
                    checkBox2.Enabled = true;
                    checkBox1.Enabled = true;

                    Marshal.ReleaseComObject(WS);
                    WS = null;
                    Marshal.ReleaseComObject(WB);
                    WB = null;
                    Marshal.ReleaseComObject(ObjExcel);
                    ObjExcel = null;

                    GC.Collect();
                }
            }
                
        }

        private void OnTimedEvent(object sender, ElapsedEventArgs e)
        {
            System.Windows.Forms.Application.DoEvents();             
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                numericUpDown1.Enabled = true;
                numericUpDown2.Enabled = true;
            }
            else
            {
                numericUpDown1.Enabled = false;
                numericUpDown2.Enabled = false;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                numericUpDown3.Enabled = true;
            else
                numericUpDown3.Enabled = false;
        }
    }
}
