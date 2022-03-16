using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Management;

namespace Petit4Tablet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            UpdatePorts();
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            UpdatePorts();
        }

        private System.Collections.ArrayList GetDeviceNames()
        {
            var result = new System.Collections.ArrayList();
            var serialPorts = SerialPort.GetPortNames();
            ManagementClass mcPnPEntity = new ManagementClass("Win32_PnPEntity");
            foreach (ManagementObject manageObj in mcPnPEntity.GetInstances())
            {
                var propName = manageObj.GetPropertyValue("Name");
                if (propName == null) continue;
                var name = propName.ToString();
                foreach (var port in serialPorts)
                {
                    if (name.IndexOf(port) >= 0)
                    {
                        result.Add(port + ": " + name);
                        break;
                    }
                }
            }
            return result;
        }

        private void UpdatePorts()
        {
            var ports = GetDeviceNames().ToArray();
            string oldText = (string)comboBox1.SelectedItem;
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(ports);
            comboBox1.SelectedIndex = Array.IndexOf(ports, oldText);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string port = (string)comboBox1.SelectedItem;
            int speed = int.Parse((string)comboBox2.SelectedItem);
            var form2 = new Form2(port.Split(':')[0], speed);
            form2.ShowDialog();
        }
    }
}
