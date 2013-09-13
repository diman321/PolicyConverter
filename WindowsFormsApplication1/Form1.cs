using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Timers;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public struct ResolvedNetObj
        {   
            public string NetObjName;
            public List<List<string>> ResolvedObj;
            private bool IsObjectNegated;

            public ResolvedNetObj(string p)
            {
                NetObjName = p;
                ResolvedObj = ResolveNetObj(p);
                IsObjectNegated = false;
            }

            public ResolvedNetObj(string p, bool n)
            {
                NetObjName = p;
                ResolvedObj = ResolveNetObj(p);
                IsObjectNegated = n;
            }

        }

        public static List<List<string>> CurrentErrorObjResolved = new List<List<string>>();


        public static List<List<string>> Rules = new List<List<string>>();
        public static List<List<IEnumerable<string>>> NetworkObj = new List<List<IEnumerable<string>>>();
        public static List<List<string>> Services = new List<List<string>>();
        private static System.Timers.Timer aTimer;
        public bool stopped = false;
        
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
                        //result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[0].ToList()[0], "" });
                        break;
                    case "Check Point Host":
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], GetRowWithResolvedInfo[2].ToList()[0], "" });
                        break;
                    case "Check Point Gateway":
                        string tmp3 = GetRowWithResolvedInfo[2].ToList()[0].Split(split, StringSplitOptions.RemoveEmptyEntries)[0].ToString();
                        result.Add(new List<string> { GetRowWithResolvedInfo[0].ToList()[0], tmp3, "" });
                        break;
                    case "Group":                        foreach (string OneOfGroupNetObj in GetRowWithResolvedInfo[6].Distinct())
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
                //MessageBox.Show(errorMessage, "Error");
                result.Add(new List<string> { NetObj + " Не найден", resExc.Message, resExc.Source });
                CurrentErrorObjResolved.Add(new List<string> { NetObj + " Не найден", resExc.Message, resExc.Source });
            }
            
            return result;

        }

        public static List<List<string>> ResolvePorts(string NetObj)
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
                //MessageBox.Show(errorMessage, "Error");
                result.Add(new List<string> { "Exception Error", resExc.Message, resExc.Source });

            }

            return result;

        }

        public void BorderAround(Microsoft.Office.Interop.Excel.Range range, int colour)
        {
            Microsoft.Office.Interop.Excel.Borders borders = range.Borders;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders.Color = colour;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            borders = null;
        }

        public virtual void OnMyEvent()
        {
            Rules.Clear();
            NetworkObj.Clear();
            Services.Clear();
            GC.Collect();
            label1.Text = "Останавился и очистил память";
            alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);
            button1.Enabled = true;
            PolycyConverter.Form2 form2 = new PolycyConverter.Form2();
            form2.ShowDialog(this);
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            stopped = false;
            button1.Enabled = false;
            
            aTimer = new System.Timers.Timer(10000);

            // Hook up the Elapsed event for the timer.
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);

            aTimer.Interval = 200;
            aTimer.Enabled = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                HtmlWeb htmlWeb = new HtmlWeb();

                // Creates an HtmlDocument object from an URL
                HtmlAgilityPack.HtmlDocument document = htmlWeb.Load(openFileDialog1.FileName);

                ///
                /// Копирую таблицы со структурой
                /// Правила - двумерная структура
                /// Сетевые объекты - трёхмерная (двумерная таблица с объектом-списком)
                /// Сервисы - двумерная таблица
                ///
                Rules = document.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[2]")
                            .Descendants("tr")
                            .Skip(1)
                            .Where(tr => tr.Elements("td").Count() > 1)
                            .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                            .ToList();

                if (stopped) 
                {
                    OnMyEvent();
                    return;
                }

                label1.Text = "Скопировал Таблицу правил";
                alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);
                System.Windows.Forms.Application.DoEvents();

                NetworkObj = document.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[4]")
                            .Descendants("tr")
                            .Skip(1)
                            .Where(tr => tr.Elements("td").Count() > 1)
                            .Select(tr => tr.Elements("td").Distinct()
                            .Select(td => td.Descendants().Distinct()
                            .Where(br => br.InnerText != "").Distinct()
                            .Select(br => br.InnerText.Trim())).Distinct().ToList().Distinct().ToList()).Distinct()
                            .ToList();

                if (stopped)
                {
                    OnMyEvent();
                    return;
                }

                label1.Text = "Скопировал Таблицу Сетевых имён";
                alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);
                System.Windows.Forms.Application.DoEvents();

                for (int k = 0; k < NetworkObj.Count; k++)
                {
                    for (int l = 0; l < NetworkObj[l].Count; l++)
                    {
                        NetworkObj[k][l] = NetworkObj[k][l].Distinct();
                        if (stopped)
                        {
                            OnMyEvent();
                            return;
                        }
                    }
                }
                
                Services = document.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[6]")
                          .Descendants("tr")
                          .Skip(1)
                          .Where(tr => tr.Elements("td").Count() > 1)
                          .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                          .ToList();

                if (stopped)
                {
                    OnMyEvent();
                    return;
                }

                label1.Text = "Скопировал Таблицу сервисов";
                alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);
                System.Windows.Forms.Application.DoEvents();

                ///
                /// Открываем книгу Excel. 
                /// 
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Workbook WB; 
                Worksheet WS;
                WB = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                WS = (Worksheet)WB.Sheets[1];
                int i = 1;

                label1.Text = "Открыл ExCel";
                alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);
                System.Windows.Forms.Application.DoEvents();

                progressBar1.Maximum = Rules.Count;
                progressBar1.Value = 0;

                try
                {
                    foreach (List<string> row in Rules)
                    {
                        if (row[2] == "SOURCE")
                        {
                            WS.Cells[i, 1].Value = "Сервис";
                            WS.Cells[i, 2].Value = "Сегмент источника";
                            WS.Cells[i, 3].Value = "IP Источника";
                            WS.Cells[i, 4].Value = "Сегмент назначения";
                            WS.Cells[i, 5].Value = "IP назначения";
                            WS.Cells[i, 6].Value = "Порты";
                            WS.Cells[i, 7].Value = "Описание";
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
                        
                        progressBar2.Maximum = CurrentSourceNetObj.Count;
                        progressBar2.Value = 0;

                        List<ResolvedNetObj> CurrentSrc = new List<ResolvedNetObj>();

                        foreach (string OneNetObj in CurrentSourceNetObj)
                        {
                            string s = OneNetObj.Replace("Not ", "").Trim();
                            ResolvedNetObj TmpRNO = new ResolvedNetObj(s, OneNetObj.Contains("Not "));
                            CurrentSrc.Add(TmpRNO);

                            progressBar2.Value++;
                            System.Windows.Forms.Application.DoEvents();
                        }

                        label1.Text = "Отрезолвил Объекты источники для " + i.ToString() + " правила";
                        alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);
                        System.Windows.Forms.Application.DoEvents();

                        ///
                        /// Это для Destination
                        ///

                        List<string> CurrentDestNetObj = row[3].Split(split, StringSplitOptions.RemoveEmptyEntries).ToList();

                        progressBar2.Maximum = CurrentDestNetObj.Count;
                        progressBar2.Value = 0;

                        List<ResolvedNetObj> CurrentDst = new List<ResolvedNetObj>();

                        foreach (string OneNetObj in CurrentDestNetObj)
                        {
                            string s = OneNetObj.Replace("Not ", "").Trim();
                            ResolvedNetObj TmpRNO = new ResolvedNetObj(s, OneNetObj.Contains("Not "));
                            CurrentDst.Add(TmpRNO);

                            progressBar2.Value++;
                            System.Windows.Forms.Application.DoEvents();
                        }


                        label1.Text = "Отрезолвил Объекты назначения для " + i.ToString() + " правила";
                        alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);
                        System.Windows.Forms.Application.DoEvents();

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
                            if (stopped)
                            {
                                OnMyEvent();
                                WB.Close(false);
                                ObjExcel.Quit();
                                return;
                            }
                            CurrentPortObjResolved.AddRange(tmp);
                            progressBar2.Value++;
                            System.Windows.Forms.Application.DoEvents();
                        }

                        label1.Text = "Отрезолвил Объекты порты для " + i.ToString() + " правила";
                        alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);
                        System.Windows.Forms.Application.DoEvents();

                        for (int a = 0 ; a < row.Count ; a++)
                        {
                            if (stopped)
                            {
                                OnMyEvent();                                
                                Marshal.ReleaseComObject(WS);
                                WS = null;
                                Marshal.ReleaseComObject(WB);
                                WB.Close(false);
                                WB = null;
                                Marshal.ReleaseComObject(ObjExcel);
                                ObjExcel.Quit();
                                ObjExcel = null;
                                
                                return;
                            }

                            if (a == 0 || a == 4 || a == 7 || a == 8)
                                continue ;

                            if (a == 1)
                                WS.Cells[i, 1].Value = row[a];

                            if (a == 2)
                            {
                                string ThisCell = "";

                                foreach (ResolvedNetObj SingleNetObj in CurrentSrc)
                                    foreach (List<string> OneElement in SingleNetObj.ResolvedObj)
                                        ThisCell += OneElement[0] + " (" + OneElement[1] + OneElement[2] + ")" + Environment.NewLine;
                                WS.Cells[i, 2].Value = row[a].Replace("\r\n", "\n").Replace("\n\n", "\n").Replace("\r\n", "\n").Replace("\n\n", "\n").Trim();
                                WS.Cells[i, 3].Value = ThisCell.Trim();
                                continue;

                            }

                            if (a == 3)
                            {
                                string ThisCell = "";

                                foreach (ResolvedNetObj SingleNetObj in CurrentDst)
                                    foreach (List<string> OneElement in SingleNetObj.ResolvedObj)
                                        ThisCell += OneElement[0] + " (" + OneElement[1] + OneElement[2] + ")" + Environment.NewLine;
                                WS.Cells[i, 4].Value = row[a].Replace("\r\n", "\n").Replace("\n\n", "\n").Replace("\r\n", "\n").Replace("\n\n", "\n").Trim();
                                WS.Cells[i, 5].Value = ThisCell.Trim();
                                continue;
                            }

                            if (a == 5)
                            {
                                string ThisCell = "";
                                foreach (List<string> PortTableRow in CurrentPortObjResolved)
                                    ThisCell += PortTableRow[0] + " (" + PortTableRow[1] +  "/" + PortTableRow[2] + ")\n";
                                WS.Cells[i, 6].Value = ThisCell.Trim();
                            }

                            if (a == 6)
                                WS.Cells[i, 7].Value = row[a];
                        }
                        if (i % 2 == 0)
                        {
                            Microsoft.Office.Interop.Excel.Range chartRange;
                            chartRange = WS.get_Range("A" + i.ToString(), "G" + i.ToString());
                            chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                        }
                        i++;
                        progressBar1.Value++;
                        progressBar1.Update();
                        label1.Text = "Записал " + i.ToString() + " правило";
                        alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);

                        System.Windows.Forms.Application.DoEvents();                        
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

                    ObjExcel.Visible = true;
                    ObjExcel.UserControl = true;
                    
                    ObjExcel.Columns.AutoFit();
                    ObjExcel.Rows.AutoFit();
                    //tbad.Start();
                    
                    PolycyConverter.Form3 form3 = new PolycyConverter.Form3();                    
                    form3.ShowDialog(this);

                    return;
                    //ObjExcel.Quit();
                    //WB.Close(true, openFileDialog1.FileName.Replace(".html", ".xlsx"));
                } 
                finally
                {
                    if (!stopped)
                    {
                        ObjExcel.Columns.AutoFit();
                        ObjExcel.Rows.AutoFit();
                        ObjExcel.Visible = true;
                        ObjExcel.UserControl = true;

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
                    Rules.Clear();
                    NetworkObj.Clear();
                    Services.Clear();
                    GC.Collect();
                    button1.Enabled = true;

                    
                    //WB.Close(false, Type.Missing, Type.Missing);
                    Marshal.ReleaseComObject(WS);
                    WS = null;
                    Marshal.ReleaseComObject(WB);
                    WB = null;
                    Marshal.ReleaseComObject(ObjExcel);
                    ObjExcel = null;
                    //t.Start();
                    
                    //WB.Close(true, openFileDialog1.FileName.Replace(".html", ".xlsx"));
                    //ObjExcel.Quit();
                }
            }
                
        }

        private void OnTimedEvent(object sender, ElapsedEventArgs e)
        {
            System.Windows.Forms.Application.DoEvents();             
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.DoEvents();
            stopped = true;
            label1.Text = "Останавливаюсь";
            alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);
            System.Windows.Forms.Application.DoEvents();
        }

    }
}
