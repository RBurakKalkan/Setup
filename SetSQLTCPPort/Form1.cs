using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SqlServer.Management.Smo.Wmi;
using Microsoft.SqlServer.Management.Smo;
using NetFwTypeLib;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace ServerConfigManipulation.cs
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        const string instanceName = "SQLEXPRESS";

        void FireWallRule()
        {
            try
            {
                Type tNetFwPolicy2 = Type.GetTypeFromProgID("HNetCfg.FwPolicy2");
                INetFwPolicy2 fwPolicy2 = (INetFwPolicy2)Activator.CreateInstance(tNetFwPolicy2);

                // Let's create a new rule
                INetFwRule2 inboundRule = (INetFwRule2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule"));
                inboundRule.Enabled = true;
                //Allow through firewall
                inboundRule.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;

                //For all profile
                inboundRule.Profiles = (int)NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_ALL;

                //Using protocol TCP
                inboundRule.Protocol = 6; // TCP
                                          //Local Port 1433
                inboundRule.LocalPorts = "1433";
                //Name of rule
                inboundRule.Name = "SQLRule";

                // Now add the rule
                INetFwPolicy2 firewallPolicy = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
                firewallPolicy.Rules.Add(inboundRule);
                MessageBox.Show("Güvenlik Duvarından Yetki Verildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        [DllImport("kernel32.dll", SetLastError = true, CallingConvention = CallingConvention.Winapi)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWow64Process([In] IntPtr hProcess, [Out] out bool lpSystemInfo);

        public bool Is64Bit()
        {
            bool retVal;
            IsWow64Process(Process.GetCurrentProcess().Handle, out retVal);

            return retVal;
        }
        void TCPPortAçma()
        {
            try
            {
                var managedComputer = new ManagedComputer(Environment.MachineName);

                if (Is64Bit())
                {
                    managedComputer.ConnectionSettings.ProviderArchitecture = ProviderArchitecture.Use64bit;
                }
                else
                {
                    managedComputer.ConnectionSettings.ProviderArchitecture = ProviderArchitecture.Use32bit;
                }

                var serviceController = new ServiceController(string.Concat("MSSQL$", instanceName));

                var serverInstance = managedComputer.ServerInstances[instanceName];

                var serverProtocol = serverInstance?.ServerProtocols["Tcp"];

                var ipAddresses = serverProtocol?.IPAddresses;
                if (ipAddresses != null)
                {
                    for (var i = 0; i < ipAddresses?.Count; i++)
                    {
                        var ipAddress = ipAddresses[i];

                        if (!string.Equals(ipAddress.Name, "IPAll"))
                        {
                            MessageBox.Show(ipAddresses[i].Name);
                            continue;
                        }

                        if (serviceController.Status.Equals(ServiceControllerStatus.Running))
                        {
                            serviceController.Stop();

                            serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                        }

                        ipAddress.IPAddressProperties["TcpDynamicPorts"].Value = "0";
                        ipAddress.IPAddressProperties["TcpPort"].Value = "1433";

                        serverProtocol.Alter();

                        break;
                    }
                }

                if (serviceController.Status.Equals(ServiceControllerStatus.Running))
                {
                    return;
                }

                serviceController.Start();

                serviceController.WaitForStatus(ServiceControllerStatus.Running);

                MessageBox.Show("TCP Port Değiştirildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("PORT DEĞİŞMEDİ!!!. LÜTFEN MANUEL PORT ATAYINIZ. (1433)", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            TCPPortAçma();

            FireWallRule();

            Close();
        }
    }
}
