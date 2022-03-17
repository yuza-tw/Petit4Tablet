using System;
using System.Diagnostics;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms;

namespace Petit4Tablet
{
    public partial class Form2 : Form
    {
        private WintabDN.CWintabContext m_context = null;
        private WintabDN.CWintabData m_data;
        private int m_maxPressure = 0;
        private Thread m_sendThread = null;

        private SerialPort m_port;

        private bool m_available = false;
        private bool m_active = false;
        private bool m_formActive = true;
        private long m_valueToSend = 0, m_numSent = 0;
        private float m_zoom = 1.0f;

        private byte[] m_keycodes = {
            4, 5, 6, 7, 8, 9, 10, 11, 12, 13,
            14, 15, 16, 17, 18, 19, 20, 21, 22, 23,
            24, 25, 26, 27, 28, 29, 30, 31, 32, 33,
            34, 35, 36, 37, 38, 39, 40, 41, 42, 43,
            44, 45, 46, 47, 48, 49, 50, 51, 52, 54,
            55, 56, 73, 74, 75, 76, 77, 78, 79, 80,
            81, 82, 84, 85, 86, 87, 88, 89, 90, 91,
            92, 93, 94, 95, 96, 97, 98, 99, 100, 103,
            104, 105, 106, 107, 108, 109, 110, 111, 112, 113,
            114, 115, 116, 117, 118, 119, 120, 121, 122, 123,
            124, 125, 126, 127, 128, 129, 133, 134, 135, 137,
            155, 156, 157, 158, 159, 160, 161, 162, 163, 164,
            176, 177, 178, 179, 180, 181, 182, 183, 184, 185,
            186, 187, 188, 189, 190, 191, 192, 193, 194, 195,
            196, 197, 198, 199, 200, 201, 202, 203, 204, 205,
            206, 207, 208, 209, 210, 211, 212, 213, 214, 215,
            216, 217, 218, 219, 220, 221, 165, 166, 167, 168,
            169, 170, 171, 172, 173, 174, 175,
            232, 233, 234, 235, 236, 237, 238, 239, 240, 241,
            242, 243, 244, 245, 246, 247, 248, 249
        };


        public Form2(string portName, int baudRate)
        {
            InitializeComponent();
            Top = Screen.PrimaryScreen.Bounds.X;
            Left = Screen.PrimaryScreen.Bounds.Y;
            Width = Screen.PrimaryScreen.Bounds.Width;
            Height = Screen.PrimaryScreen.Bounds.Height;
            m_zoom = Math.Min(1280.0f / Width, 720.0f / Height);
            m_port = new SerialPort(portName, baudRate, Parity.None, 8, StopBits.One);
            m_port.DataReceived += port_DataReceived;
            m_port.ReadTimeout = 50;
            m_port.WriteTimeout = 50;
            m_port.DtrEnable = true;
            m_sendThread = new Thread(new ThreadStart(ThreadFunc));
            m_sendThread.Start();
        }

        private void ThreadFunc()
        {
            int releaseCounter = 0;
            long lastValueToSend = 0;
            while (true)
            {
                if (!m_port.IsOpen)
                {
                    Thread.Sleep(16);
                    continue;
                }

                if (m_valueToSend == 0)
                {
                    Thread.Sleep(4);
                    if (--releaseCounter == 0)
                    {
                        SendRelease();
                    }
                    continue;
                }

                if (m_valueToSend < 0) break;

                if (m_active && lastValueToSend != m_valueToSend)
                {
                    lastValueToSend = m_valueToSend;
                    try
                    {
                        m_port.Write(CreateP4SendPacket(m_valueToSend), 0, 12);
                        m_numSent++;
                        releaseCounter = 30;
                    }
                    catch
                    {
                         // Timeout
                    }
                }
                m_valueToSend = 0;
            }
            SendRelease();
        }

        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                byte[] buffer = new byte[m_port.BytesToRead];
                m_port.Read(buffer, 0, buffer.Length);
            }
            catch
            {
            }
        }

        private void SendRelease()
        {
            if (!m_port.IsOpen) return;
            try
            {
                byte[] relSeq = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                m_port.Write(relSeq, 0, 12);
            }
            catch { }
        }


        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) Close();
            if (e.KeyCode == Keys.Space) ToggleSend();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            tabletInput.Text = "";
            timer1.Start();

            try
            {
                m_port.Open();
            }
            catch { }

            if (!m_port.IsOpen)
            {
                MessageBox.Show("COM ポートが利用できません");
                Close();
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            while (m_valueToSend != 0) Thread.Sleep(8);
            m_valueToSend = -1;
            m_sendThread.Join();

            Disconnect();
            if (m_port.IsOpen)
            {
                m_port.DataReceived -= port_DataReceived;
                while (!(m_port.BytesToRead == 0 && m_port.BytesToWrite == 0))
                {
                    m_port.DiscardInBuffer();
                    m_port.DiscardOutBuffer();
                }
                m_port.Close();
            }
        }

        private void Connect()
        {
            if (m_context != null) return;

            m_context = WintabDN.CWintabInfo.GetDefaultDigitizingContext(
                WintabDN.ECTXOptionValues.CXO_MESSAGES
            );
            m_context.OutOrgX = 0;
            m_context.OutOrgY = 0;
            m_context.OutExtX = 1280;
            m_context.OutExtY = 720;
            m_context.PktRate = 60;
            m_maxPressure = WintabDN.CWintabInfo.GetMaxPressure();

            var status = m_context.Open();
            if (!status)
            {
                m_context = null;
                tabletInput.Text = "タブレットドライバを初期化できませんでした";
                return;
            }

            m_data = new WintabDN.CWintabData(m_context);
            m_data.SetWTPacketEventHandler(OnReceivedWTPacket);

            m_available = true;
            m_active = true;
        }

        private void Disconnect()
        {
            if (m_context == null) return;

            m_data = null;
            m_context.Close();
            m_context = null;
            m_available = false;
            m_active = false;
        }

        private void OnReceivedWTPacket(Object sender, WintabDN.MessageReceivedEventArgs eventArgs)
        {
            if (m_data == null) return;

            uint nPackets = 0;
            var packets = m_data.GetDataPackets(m_data.GetPacketQueueSize(), true, ref nPackets);
            if (nPackets == 0) return;

            var pkt = packets[packets.Length - 1];
            if (pkt.pkContext == 0) return;

            int p = pkt.pkNormalPressure.pkRelativeNormalPressure * 255 / m_maxPressure;
            int x = pkt.pkX, y = 719 - pkt.pkY;
            m_valueToSend = CreateHash(x, y, p, (int)pkt.pkButtons & 1, ((int)pkt.pkButtons >> 1) & 1);
            tabletInput.Text = string.Format("X:{1}, Y:{2}, Z:{3}, P:{4}, B:{5}\nPitch:{6}, Yaw:{7}, Roll:{8}\nAltitude:{9}, Azimuth:{10}, Twist:{11}",
                0, x, y, pkt.pkZ, p, pkt.pkButtons,
                pkt.pkRotation.rotPitch, pkt.pkRotation.rotYaw, pkt.pkRotation.rotRoll, 
                pkt.pkOrientation.orAltitude, pkt.pkOrientation.orAzimuth, pkt.pkOrientation.orTwist
            );
        }

        private long CreateHash(int tX, int tY, int tP, int b0, int b1)
        {
            return tX | (tY << 11) | (tP << 21) | (b0 << 29) | (b1 << 30);
        }

        private byte[] CreateP4SendPacket(long v)
        {
            long c = 1;
            int r = 6, n = r;
            int[] keyary = { 0, 0, 0, 0, 0, 0 };

            while (true)
            {
                long c2 = c * (n + 1) / (n + 1 - r);
                if (c2 > v) break;
                c = c2;
                n++;
            }

            while (n >= r && r > 0)
            {
                if (v >= c)
                {
                    v -= c;
                    c = c * r / (n - r + 1);
                    r--;
                    keyary[5 - r] = n;
                }
                c = c * (n - r) / n;
                n--;
            }

            while (r > 0)
            {
                r--;
                keyary[5 - r] = n;
                n--;
            }

            byte[] result = { 0, 0, 0, 0,  0, 0, 0, 0,  0, 0, 0, 0 };
            for (int i = 0; i < 6; i++)
                result[i + 1] = m_keycodes[keyary[i]];

            return result;
        }

        private void ToggleSend()
        {
            if (!m_available) return;
            m_active = !m_active;
        }

        private void Form2_Activated(object sender, EventArgs e)
        {
            m_formActive = true;
        }

        private void Form2_Deactivate(object sender, EventArgs e)
        {
            m_formActive = false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (WintabDN.CWintabInfo.IsWintabAvailable())
            {
                statusLabel.Text = m_active ? "送信中" : "送信停止中";
                Connect();
            }
            else
            {
                statusLabel.Text = "タブレット接続待機中";
                Disconnect();
            }
        }

    }
}
