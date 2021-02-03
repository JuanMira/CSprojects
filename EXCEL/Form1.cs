using System;
using System.Windows.Forms;
using System.IO.Ports;
using System.Data.OleDb;

namespace EXCEL
{
    public partial class Form1 : Form
    {
        string dato;
        string puertoSlect;        
        string fechaNow;
        string horaNow;
        string fechaTot;
        public Form1()
        {
            InitializeComponent();
            string[] puertos = SerialPort.GetPortNames();
            foreach (string mostrar in puertos) {
                comboBox1.Items.Add(mostrar);
            }
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            serialPort1.Close();
            serialPort1.Dispose();
            serialPort1.PortName = puertoSlect;

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            serialPort1.Close();
            serialPort1.Dispose();
            puertoSlect = comboBox1.Text;
            serialPort1.PortName = puertoSlect;
            serialPort1.Open();
            CheckForIllegalCrossThreadCalls = false;
            if (serialPort1.IsOpen == true) {
                connect.Text = "Conectado";
            }
            else
            {
                return;
            }
        }

        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                serialPort1.BaudRate = 9600;
                dato = serialPort1.ReadLine();
                char[] limitador = { ',', ' ' };
                string[] trozos = dato.Split(limitador);                
                fechaNow = DateTime.Now.ToString("dd/MM/yyyy");
                horaNow = DateTime.Now.ToString("h:mm:ss tt");
                fechaTot = fechaNow + " " + horaNow;
                excel(trozos,fechaTot);
            }
            catch (Exception ex) {
                MessageBox.Show("A ocurrido un error, revice el puerto Serial");
            }


        }

        private void excel(string[] data,string fechaInicial) {         
            string filename = @"E:\Documents\Reporte.xlsx";
            String connection = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
            string command = "insert into[sheet1$](fecha_inicial,autor,T1,T2,T3,B,AnalogRead,fecha_final) values("+fechaInicial+","+data[0]+"," + data[1] + "," + data[2] + "," + data[3] + ","+data[4]+","+data[5]+",NOW())";
            OleDbConnection con = new OleDbConnection(connection);
            con.Open();
            OleDbCommand cmd = new OleDbCommand(command,con);
            cmd.ExecuteNonQuery();
        }

        private void reset_Click(object sender, EventArgs e)
        {
            serialPort1.WriteLine("r"); 
        }  
    }
}
