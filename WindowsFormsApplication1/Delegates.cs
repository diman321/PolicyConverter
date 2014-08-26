using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Globalization;

namespace WindowsFormsApplication1
{

    public partial class Form1 
    {

        public void AddTextToStatus(string s)
        {
            if (label1.InvokeRequired)
                label1.BeginInvoke(new MethodInvoker(() => label1.Text = s));
            else
                label1.Text = s;

            if (alphaBlendTextBox2.InvokeRequired)
                alphaBlendTextBox2.BeginInvoke(new MethodInvoker(() => alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine)));
            else
                alphaBlendTextBox2.AppendText(label1.Text + Environment.NewLine);

            System.Windows.Forms.Application.DoEvents();
        }

        public delegate List<List<string>> Rulesd_Delegate(string s);

        public List<List<string>> Rulesd(string s)
        {
            // Write your time consuming task here
            HtmlWeb htmlWeb = new HtmlWeb();

            // Creates an HtmlDocument object from an URL
            HtmlAgilityPack.HtmlDocument document = htmlWeb.Load(s);

            List<List<string>> rulesindelegate = new List<List<string>>();

            rulesindelegate = document.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[2]")
                        .Descendants("tr")
                        .Skip(1)
                        .Where(tr => tr.Elements("td").Count() > 1)
                        .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                        .ToList();

            return rulesindelegate;
        }

        public delegate List<List<IEnumerable<string>>> NetworkObjd_Delegate(string s);

        public List<List<IEnumerable<string>>> NetworkObjd(string s)
        {
            // Write your time consuming task here
            HtmlWeb htmlWeb = new HtmlWeb();

            // Creates an HtmlDocument object from an URL
            HtmlAgilityPack.HtmlDocument document = htmlWeb.Load(s);

            List<List<IEnumerable<string>>> NetworkObjindelegate = new List<List<IEnumerable<string>>>();

            NetworkObjindelegate = document.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[4]")
                            .Descendants("tr")
                            .Skip(1)
                            .Where(tr => tr.Elements("td").Count() > 1)
                            .Select(tr => tr.Elements("td").Distinct()
                            .Select(td => td.Descendants().Distinct()
                            .Where(br => br.InnerText != "").Distinct()
                            .Select(br => br.InnerText.Trim())).Distinct().ToList().Distinct().ToList()).Distinct()
                            .ToList();

            return NetworkObjindelegate;
        }


        public delegate List<List<string>> Servicesd_Delegate(string s);

        public List<List<string>> Servicesd(string s)
        {
            // Write your time consuming task here
            HtmlWeb htmlWeb = new HtmlWeb();

            // Creates an HtmlDocument object from an URL
            HtmlAgilityPack.HtmlDocument document = htmlWeb.Load(s);

            List<List<string>> Servicesindelegate = new List<List<string>>();

            Servicesindelegate = document.DocumentNode.SelectSingleNode("/html[1]/body[1]/table[6]")
                          .Descendants("tr")
                          .Skip(1)
                          .Where(tr => tr.Elements("td").Count() > 1)
                          .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                          .ToList();

            return Servicesindelegate;
        }

        public void AddWStoWB(ResolvedNetObj SingleNetObj, Workbook WB)
        {
            Worksheet tmpWS = (Worksheet)WB.Worksheets.Add();
            tmpWS.Name = SingleNetObj.NetObjName.Replace("?", "").Replace("!", "").Replace(";", "").Trim();
            //tmpWS.Tab.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            tmpWS.Cells[1, 1].Value = "Имя";
            tmpWS.Cells[1, 2].Value = "IP";
            tmpWS.Cells[1, 3].Value = "Маска";

            for (int j = 0; j < SingleNetObj.ResolvedObj.Count; j++)
            {
                tmpWS.Cells[j + 2, 1].Value = SingleNetObj.ResolvedObj[j][0];
                tmpWS.Cells[j + 2, 2].Value = SingleNetObj.ResolvedObj[j][1];
                tmpWS.Cells[j + 2, 3].Value = SingleNetObj.ResolvedObj[j][2];
            }

            tmpWS.Columns.AutoFit();
        }
        
        public void AddNetworkObj(List<ResolvedNetObj> Current, Workbook WB, Worksheet WS, int currentrow, int startplase)
        {
			string ThisCellResolve = "";
			string ThisCell = "";
            int allobj = 0;

            foreach (ResolvedNetObj SingleNetObj in Current)
                allobj += SingleNetObj.ResolvedObj.Count;


            foreach (ResolvedNetObj SingleNetObj in Current)
            {			
				string Negated = "";
				if (SingleNetObj.IsObjectNegated)
					Negated = "Not ";
                if (checkBox1.Checked == true && (((SingleNetObj.ResolvedObj.Count > Convert.ToInt32(numericUpDown3.Value))) || (allobj > 25 && SingleNetObj.ResolvedObj.Count > 2)))
                {
                    if (AllWSNames.Where(n => n.Contains(SingleNetObj.NetObjName)).ToList().Count == 0)
                    {
                        AddWStoWB(SingleNetObj, WB);
                        AllWSNames.Add(SingleNetObj.NetObjName);

						ThisCell += Negated + SingleNetObj.NetObjName + Environment.NewLine;
						ThisCellResolve += Negated + SingleNetObj.NetObjName + Environment.NewLine;
                    }
                    else
                    {
						ThisCell += Negated + SingleNetObj.NetObjName + Environment.NewLine;
						ThisCellResolve += Negated + SingleNetObj.NetObjName + Environment.NewLine;
                    }
                }
                else
                {
                    string Tmp4LittleGroupObj = "";
                    foreach (List<string> OneElement in SingleNetObj.ResolvedObj)
						Tmp4LittleGroupObj += Negated + OneElement[0] + " (" + OneElement[1] + OneElement[2] + ")" + Environment.NewLine;
					ThisCell += Negated + SingleNetObj.NetObjName + Environment.NewLine;
                    ThisCellResolve += Tmp4LittleGroupObj.Trim() + Environment.NewLine;              
                }
            }
            WS.Cells[currentrow, startplase].Value = ThisCell.Trim();
            WS.Cells[currentrow, startplase + 1].Value = ThisCellResolve.Replace(" Не найден ()", "").Trim();
        }

        public List<ResolvedNetObj> GetResolvedNetObj(List<string> CurrentNetObj)
        {
            if (progressBar2.InvokeRequired)
                progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Maximum = CurrentNetObj.Count));
            else
                progressBar2.Maximum = CurrentNetObj.Count;

            if (progressBar2.InvokeRequired)
                progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Value = 0));
            else
                progressBar2.Value = 0;

            List<ResolvedNetObj> CurrentResolvedNetObjs = new List<ResolvedNetObj>();

            foreach (string OneNetObj in CurrentNetObj)
            {
                string s = OneNetObj.Replace("Not ", "").Trim();
                ResolvedNetObj TmpRNO = new ResolvedNetObj(s, OneNetObj.Contains("Not "));

                CurrentResolvedNetObjs.Add(TmpRNO);

                if (progressBar2.InvokeRequired)
                    progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Value++));
                else
                    progressBar2.Value++;

                System.Windows.Forms.Application.DoEvents();
            }            
            return CurrentResolvedNetObjs;
        }

        public void AddNodeList(Workbook WB)
        {
            List<NetInfo> AllNets = GetNets();

            Worksheet tmpWS = (Worksheet)WB.Worksheets.Add();
            tmpWS.Activate();
            tmpWS.Name = "Node";
            
            tmpWS.Cells[2, 1].Value = "Динамический объект";
            tmpWS.Cells[2, 2].Value = "FQDN (если применимо)";
            tmpWS.Cells[2, 3].Value = "IP/MASK";
            tmpWS.Cells[2, 4].Value = "Включенные объекты (для групп)";
            tmpWS.Cells[2, 5].Value = "Система (сервис)";
            tmpWS.Cells[2, 6].Value = "Сегмент сети";
            tmpWS.Cells[2, 7].Value = "Категория обрабатываемой информации";
            tmpWS.Cells[2, 8].Value = "Примечание";

            int i = 3;

            if (progressBar2.InvokeRequired)
                progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Maximum = ToNodeList.Count));
            else
                progressBar2.Maximum = ToNodeList.Count;

            if (progressBar2.InvokeRequired)
                progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Value = 0));
            else
                progressBar2.Value = 0;

            foreach (ResolvedNetObj oneNetObj in ToNodeList)
            {
                tmpWS.Cells[i, 1].Value = oneNetObj.NetObjName;
                if (oneNetObj.ResolvedObj.Count > 1)
                {
                    string Entre4Cell = "";
                    foreach (List<string> OneResNetObj in oneNetObj.ResolvedObj)
                        Entre4Cell += OneResNetObj[0] + Environment.NewLine;
                    tmpWS.Cells[i, 4].Value = Entre4Cell.Trim();
                }
                else
                {
                    tmpWS.Cells[i, 3].Value = oneNetObj.ResolvedObj[0][1] + oneNetObj.ResolvedObj[0][2];
                    List<NetInfo> InfoAboutNetFromIPDB = AllNets.Where(n => n.Net.Replace("/32", "") == oneNetObj.ResolvedObj[0][1] + oneNetObj.ResolvedObj[0][2]).ToList();
                    if (InfoAboutNetFromIPDB.Count == 0 && oneNetObj.ResolvedObj[0][2] == "")
                        InfoAboutNetFromIPDB = AllNets.Where(n => IP2Int(oneNetObj.ResolvedObj[0][1]) >= n.StartIP && IP2Int(oneNetObj.ResolvedObj[0][1]) <= n.EndIP).ToList();
                    if (InfoAboutNetFromIPDB.Count > 0)
                        tmpWS.Cells[i, 6].Value = InfoAboutNetFromIPDB[0].Descr;
                }
                
                if (progressBar2.InvokeRequired)
                    progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Value++));
                else
                    progressBar2.Value++;
                System.Windows.Forms.Application.DoEvents();

                i++;
            }

            Microsoft.Office.Interop.Excel.Range c1 = tmpWS.Cells[2, 1];
            Microsoft.Office.Interop.Excel.Range c2 = tmpWS.Cells[2, 8];
            Range oRange = (Microsoft.Office.Interop.Excel.Range)tmpWS.get_Range(c1, c2);
            oRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
            oRange.Font.Bold = true;

            c1 = tmpWS.Cells[1, 1];
            c2 = tmpWS.Cells[i, 8];
            oRange = (Microsoft.Office.Interop.Excel.Range)tmpWS.get_Range(c1, c2);
            oRange.Columns.AutoFit();
            oRange.Rows.AutoFit();

            tmpWS.Cells[1, 1].Font.Size = 18;
            tmpWS.Cells[1, 1].Font.Bold = true;
            tmpWS.Cells[1, 1].Value = "Описание объектов";

            tmpWS.Application.ActiveWindow.SplitRow = 2;
            tmpWS.Application.ActiveWindow.FreezePanes = true;
        }

		public void AddAllNodeList(Workbook WB)
		{
			List<NetInfo> AllNets = GetNets();

			Worksheet tmpWS = (Worksheet)WB.Sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
			tmpWS.Activate();
			tmpWS.Name = "AllNode";

			tmpWS.Cells[2, 1].Value = "Динамический объект";
			tmpWS.Cells[2, 2].Value = "FQDN (если применимо)";
			tmpWS.Cells[2, 3].Value = "IP/MASK";
			tmpWS.Cells[2, 4].Value = "Включенные объекты (для групп)";
			tmpWS.Cells[2, 5].Value = "Система (сервис)";
			tmpWS.Cells[2, 6].Value = "Сегмент сети";
			tmpWS.Cells[2, 7].Value = "Категория обрабатываемой информации";
			tmpWS.Cells[2, 8].Value = "Примечание";

			int i = 3;

			if (progressBar2.InvokeRequired)
				progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Maximum = NetworkObj.Count * 2));
			else
				progressBar2.Maximum = NetworkObj.Count * 2;

			if (progressBar2.InvokeRequired)
				progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Value = 0));
			else
				progressBar2.Value = 0;

			List<ResolvedNetObj> AllNodesToNodeList = new List<ResolvedNetObj>();
			
			foreach (List<IEnumerable<string>> OneNetObj in NetworkObj)
			{	   
				ResolvedNetObj tmp = new ResolvedNetObj(OneNetObj[0].ToList()[0].Trim().ToLower());
				AllNodesToNodeList.Add(tmp);

				if (progressBar2.InvokeRequired)
					progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Value++));
				else
					progressBar2.Value++;
				System.Windows.Forms.Application.DoEvents();
			}

			foreach (ResolvedNetObj oneNetObj in AllNodesToNodeList)
			{	
				if (i == 237)
				{
					continue;
				}
				
				tmpWS.Cells[i, 1].Value = oneNetObj.NetObjName;
				if (oneNetObj.ResolvedObj.Count > 1)
				{
					string Entre4Cell = "";
					foreach (List<string> OneResNetObj in oneNetObj.ResolvedObj)
						Entre4Cell += OneResNetObj[0] + Environment.NewLine;
					tmpWS.Cells[i, 4].Value = Entre4Cell.Trim();
				}
				else
				{
					tmpWS.Cells[i, 3].Value = oneNetObj.ResolvedObj[0][1] + oneNetObj.ResolvedObj[0][2];
					List<NetInfo> InfoAboutNetFromIPDB = AllNets.Where(n => n.Net.Replace("/32", "") == oneNetObj.ResolvedObj[0][1] + oneNetObj.ResolvedObj[0][2]).ToList();
					if (InfoAboutNetFromIPDB.Count == 0 && oneNetObj.ResolvedObj[0][2] == "")
						InfoAboutNetFromIPDB = AllNets.Where(n => IP2Int(oneNetObj.ResolvedObj[0][1]) >= n.StartIP && IP2Int(oneNetObj.ResolvedObj[0][1]) <= n.EndIP).ToList();
					if (InfoAboutNetFromIPDB.Count > 0)
						tmpWS.Cells[i, 6].Value = InfoAboutNetFromIPDB[0].Descr.Replace("/", "").Replace("\\", "").Trim().ToString(new CultureInfo("en-US"));
				}

				if (progressBar2.InvokeRequired)
					progressBar2.BeginInvoke(new MethodInvoker(() => progressBar2.Value++));
				else
					progressBar2.Value++;
				System.Windows.Forms.Application.DoEvents();

				i++;
			}

			Microsoft.Office.Interop.Excel.Range c1 = tmpWS.Cells[2, 1];
			Microsoft.Office.Interop.Excel.Range c2 = tmpWS.Cells[2, 8];
			Range oRange = (Microsoft.Office.Interop.Excel.Range)tmpWS.get_Range(c1, c2);
			oRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
			oRange.Font.Bold = true;

			c1 = tmpWS.Cells[1, 1];
			c2 = tmpWS.Cells[i, 8];
			oRange = (Microsoft.Office.Interop.Excel.Range)tmpWS.get_Range(c1, c2);
			oRange.Columns.AutoFit();
			oRange.Rows.AutoFit();

			tmpWS.Cells[1, 1].Font.Size = 18;
			tmpWS.Cells[1, 1].Font.Bold = true;
			tmpWS.Cells[1, 1].Value = "Описание объектов";

			tmpWS.Application.ActiveWindow.SplitRow = 2;
			tmpWS.Application.ActiveWindow.FreezePanes = true;
		}
    }
}
