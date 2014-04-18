using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{

    public partial class Form1
    {
        public struct NetInfo
        {
            public string Net;
            public string Descr;
            public uint StartIP;
            public uint EndIP;

            public NetInfo(string tmpNet, string tmpDescr, uint SIP, uint EIP)
            {
                Net = tmpNet;
                Descr = tmpDescr;
                StartIP = SIP;
                EndIP = EIP;
            }
        }

        public struct IPandNet
        {
            public string IP;
            public NetInfo NI;

            public IPandNet(string tmpIP, NetInfo tmpNI)
            {
                IP = tmpIP;
                NI = tmpNI;
            }
        }

        public static uint IP2Int(string IPNumber)
        {
            uint ip = 0;
            string[] elements = IPNumber.Split(new Char[] { '.' });
            if (elements.Length == 4)
            {
                ip = Convert.ToUInt32(elements[0]) << 24;
                ip += Convert.ToUInt32(elements[1]) << 16;
                ip += Convert.ToUInt32(elements[2]) << 8;
                ip += Convert.ToUInt32(elements[3]);
            }
            return ip;
        }

        public void Network2IpRange(string sNetwork, out uint startIP, out uint endIP)
        {
            uint ip,		/* ip address */
                mask,		/* subnet mask */
                broadcast,	/* Broadcast address */
                network;	/* Network address */
            //usableIps;

            int bits;

            string[] elements = sNetwork.Split(new Char[] { '\\', '/' });



            ip = IP2Int(elements[0]);

            if (elements[1] == "32")
            {
                startIP = endIP = ip;
                return;
            }

            bits = Convert.ToInt32(elements[1]);

            mask = ~(0xffffffff >> bits);

            network = ip & mask;
            broadcast = network + ~mask;

            //usableIps = (bits > 30) ? 0 : (broadcast - network - 1);

            startIP = network;
            endIP = broadcast;

        }

        public List<NetInfo> GetNets()
        {
            MySqlConnection connection = new MySqlConnection("SERVER=10.46.48.180;" +
                                                 "DATABASE=ipdb;" +
                                                 "UID=reader;" +
                                                 "PASSWORD=reader;");

            MySqlConnection connectionRN = new MySqlConnection("SERVER=10.19.2.2;" +
                                                 "DATABASE=ipdb;" +
                                                 "UID=amrodchenko;" +
                                                 "PASSWORD=amr3588;");

            List<NetInfo> AllNets = new List<NetInfo>();

            try
            {
                connection.Open();

                MySqlCommand CmdReadNet = new MySqlCommand(
                    "SELECT CONCAT(CONVERT(INET_NTOA(ip) USING utf8), '/', CONVERT(prefix USING utf8)), descr FROM ipdb_sp.ipdb WHERE (descr is not NULL) AND (ip > 0) order by prefix desc;", connection);
                CmdReadNet.CommandTimeout = 120;

                MySqlDataReader dataReader = CmdReadNet.ExecuteReader();

                for (int i = 0; dataReader.Read(); i++)
                {
                    string tmpNet = dataReader.GetString(0).Trim();
                    string tmpDescr = dataReader.GetString(1).Trim();
                    uint startIP = 0, endIP = 0;
                    Network2IpRange(tmpNet, out startIP, out endIP);
                    NetInfo tmpItem = new NetInfo(tmpNet, tmpDescr, startIP, endIP);
                    AllNets.Add(tmpItem);
                }

                dataReader.Close();
                dataReader.Dispose();

                connection.Close();
                connection.Dispose();

                CmdReadNet.Dispose();

            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                connectionRN.Open();
                
                MySqlCommand CmdReadNet = new MySqlCommand(
                    "SELECT CONCAT(CONVERT(INET_NTOA(ip) USING utf8), '/', CONVERT(prefix USING utf8)), descr \n" +
                    "FROM ipdb.ipdb \n" +
                    "WHERE (descr is not NULL) AND NOT (\n" +
                    "\t(ip BETWEEN INET_ATON('10.32.2.0') AND INET_ATON('10.32.2.255')) OR\n" +
                    "\t(ip BETWEEN INET_ATON('10.44.0.0') AND INET_ATON('10.47.255.255')) OR\n" +
                    "\t(ip BETWEEN INET_ATON('10.121.0.0') AND INET_ATON('10.121.255.255')) OR\n" +
                    "\t(ip BETWEEN INET_ATON('10.223.0.0') AND INET_ATON('10.223.255.255')) OR\n" +
                    "\t(ip BETWEEN INET_ATON('10.252.0.0') AND INET_ATON('10.252.255.255')) OR\n" +
                    "\t(ip BETWEEN INET_ATON('109.236.254.0') AND INET_ATON('109.236.255.255')) OR\n" +
                    "\t(ip = INET_ATON('0.0.0.0'))\n" +
                    ")" +
                    "order by prefix desc;", connectionRN);
                CmdReadNet.CommandTimeout = 120;

                MySqlDataReader dataReader = CmdReadNet.ExecuteReader();


                for (int i = 0; dataReader.Read(); i++)
                {
                    string tmpNet = dataReader.GetString(0).Trim();
                    string tmpDescr = dataReader.GetString(1).Trim();
                    uint startIP = 0, endIP = 0;
                    Network2IpRange(tmpNet, out startIP, out endIP);
                    NetInfo tmpItem = new NetInfo(tmpNet, tmpDescr, startIP, endIP);
                    AllNets.Add(tmpItem);
                }

                dataReader.Close();
                dataReader.Dispose();

                connection.Close();
                connection.Dispose();

            }
            catch (MySqlException ex)
            {
                Console.WriteLine(ex);
            }

            List<NetInfo> sortedList = (from elem in AllNets
                                        orderby (elem.EndIP - elem.StartIP)
                                        select elem).ToList();


            return sortedList;
        }
    }
}