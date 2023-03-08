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
using System.Threading;
using SpreadsheetLight;

namespace Proyecto_Tópicos_III
{
    public partial class Form1 : Form
    {
        private int status = 0;
        private int i = 0;
        private int j = 0;
        private int x = 0;
        private int y = 0;
        private byte[] datosentrada = new byte[9];
        private Queue<byte> recievedData = new Queue<byte>();
        private int n = 0;
        private int auto = 0;
        private int sola = 0;
        private int vacio = 1;
        private int activo = 0;
        private int activa = 0;
        private int ciclado = 0;
        private int espera = 0;
        private int detener = 0;
        private decimal tiempohoras = 0;
        private decimal tiempominutos = 0;
        private decimal tiemposegundos = 0;
        private decimal tiempotranscurrido = 0;
        private decimal tiempomuestra = 0;
        private decimal tiempofinal = 0;
        private decimal a = 0;
        private decimal b = 0;
        private decimal c = 0;
        private DateTime tiempoinicial = DateTime.Now;
        private int controlcolor = 0;
        private byte[] datossalida = new byte[9];
        public Form1()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            BotonConectar.Enabled = false;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void BotonEscanear_Click(object sender, EventArgs e)
        {
            string[] PuertosDisponibles = SerialPort.GetPortNames();

            ComboPuerto.Items.Clear();

            foreach (string puerto_simple in PuertosDisponibles)
            {
                ComboPuerto.Items.Add(puerto_simple);
            }

            if(ComboPuerto.Items.Count > 0)
            {
                ComboPuerto.SelectedIndex = 0;
                BotonConectar.Enabled = true;
                BotonConectar.BackColor = Color.Lime;
                ComboBaudrate.Enabled = true;
            }
            else
            {
                MessageBox.Show("No se encontraron puertos");
                ComboPuerto.Items.Clear();
                ComboPuerto.Text = "                    ";
                BotonConectar.Enabled = false;
            }
        }

        private void BotonConectar_Click(object sender, EventArgs e)
        {
            try
            {
                if (BotonConectar.Text == "Conectar") 
                {
                    PuertoSerial.BaudRate = Int32.Parse(ComboBaudrate.Text);
                    PuertoSerial.DataBits = 8;
                    PuertoSerial.Parity = Parity.None;
                    PuertoSerial.StopBits = StopBits.One;
                    PuertoSerial.Handshake = Handshake.None;
                    PuertoSerial.PortName = ComboPuerto.Text;
                    PuertoSerial.ReadTimeout = 500;
                    PuertoSerial.WriteTimeout = 500;
                    try
                    {
                        PuertoSerial.Open();
                        status = 1;
                        BotonConectar.Text = "Desconectar";
                        BotonConectar.BackColor = Color.Crimson;
                        NumericBase.Enabled = true;
                        NumericA1.Enabled = true;
                        NumericA2.Enabled = true;
                        NumericA3.Enabled = true;
                        NumericA4.Enabled = true;
                        NumericP.Enabled = true;
                        NumericV.Enabled = true;
                        BotonInsertar.Enabled = true;
                        BotonInsertarNormal.Enabled = true;
                        BotonGuardar.Enabled = true;
                        BotonCargar.Enabled = true;
                        BotonActiva.Enabled = true;
                        BotonActiva.BackColor = Color.Lime;
                        if (vacio == 0)
                        {
                            BotonArriba.Enabled = true;
                            BotonAbajo.Enabled = true;
                            BotonSolo.Enabled = true;
                            BotonAuto.Enabled = true;
                            BotonBorrar.Enabled = true;
                            BotonLoop.Enabled = true;
                        }
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show(exc.Message.ToString());
                    }
                }
                else if (BotonConectar.Text == "Desconectar")
                {
                    PuertoSerial.Close();
                    status = 0;
                    ciclado = 0;
                    auto = 0;
                    BotonConectar.Text = "Conectar";
                    BotonConectar.BackColor = Color.Lime;
                    label9.Text = "---";
                    label10.Text = "---";
                    label11.Text = "---";
                    label12.Text = "---";
                    label13.Text = "---";
                    label14.Text = "---";
                    label16.Text = "---";
                    NumericBase.Enabled = false;
                    NumericA1.Enabled = false;
                    NumericA2.Enabled = false;
                    NumericA3.Enabled = false;
                    NumericA4.Enabled = false;
                    NumericP.Enabled = false;
                    NumericV.Enabled = false;
                    BotonInsertar.Enabled = false;
                    BotonInsertarNormal.Enabled = false;
                    BotonArriba.Enabled = false;
                    BotonAbajo.Enabled = false;
                    BotonSolo.Enabled = false;
                    BotonAuto.Enabled = false;
                    BotonBorrar.Enabled = false;
                    BotonGuardar.Enabled = false;
                    BotonCargar.Enabled = false;
                    BotonActiva.Enabled = false;
                    BotonActiva.BackColor = Color.Silver;
                    BotonLoop.Enabled = false;
                }
            }
            catch (Exception exc)
            {
                if(ComboBaudrate.Text == "Seleccionar")
                {
                    MessageBox.Show("Seleccionar un valor en Tasa de Baudios");
                }
                MessageBox.Show(exc.Message.ToString());
            }
        }

        private void label3_Click_1(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            HoraLabel.Text = DateTime.Now.ToString("HH:mm");
            DiaLabel.Text = DateTime.Now.ToLongDateString();
            if (status == 1)
            {
                byte[] data = new byte[PuertoSerial.BytesToRead];
                PuertoSerial.Read(data, 0, data.Length);
                data.ToList().ForEach(b => recievedData.Enqueue(b));
                processData();
                if (data.Length == 9)
                {
                    if (data[0] == 253 & data[8] == 254)
                    {
                        label9.Text = Convert.ToString(data[1]);
                        label10.Text = Convert.ToString(data[2]);
                        label11.Text = Convert.ToString(data[3]);
                        label12.Text = Convert.ToString(data[4]);
                        label13.Text = Convert.ToString(data[5]);
                        label14.Text = Convert.ToString(data[6]);
                        label16.Text = Convert.ToString(data[7]);

                    }
                }
                if (auto == 1)
                {
                    if(data.Length == 9)
                    {
                        datossalida[0] = 253;
                        datossalida[1] = Convert.ToByte(TablaProgramacion.Rows[j].Cells[1].Value);
                        datossalida[2] = Convert.ToByte(TablaProgramacion.Rows[j].Cells[2].Value);
                        datossalida[3] = Convert.ToByte(TablaProgramacion.Rows[j].Cells[3].Value);
                        datossalida[4] = Convert.ToByte(TablaProgramacion.Rows[j].Cells[4].Value);
                        datossalida[5] = Convert.ToByte(TablaProgramacion.Rows[j].Cells[5].Value);
                        datossalida[6] = Convert.ToByte(TablaProgramacion.Rows[j].Cells[6].Value);
                        datossalida[7] = Convert.ToByte(TablaProgramacion.Rows[j].Cells[7].Value);
                        datossalida[8] = 254;

                        int contador = 0;

                        foreach (var elemento in data)
                        {
                            if (datossalida.Contains(elemento))
                            {
                                contador++;
                            }
                        }

                        int coincidencias = data.Intersect(datossalida).Count();

                        if (contador == 9)
                        {
                            Thread.Sleep(250);
                            espera++;

                            if (espera >=1)
                            {
                                espera = 0;
                                i = j + 1;
                                TablaProgramacion.Rows[j].DefaultCellStyle.BackColor = SystemColors.Window;
                                if (i <= n)
                                {
                                    datossalida[0] = 253;
                                    datossalida[1] = Convert.ToByte(TablaProgramacion.Rows[i].Cells[1].Value);
                                    datossalida[2] = Convert.ToByte(TablaProgramacion.Rows[i].Cells[2].Value);
                                    datossalida[3] = Convert.ToByte(TablaProgramacion.Rows[i].Cells[3].Value);
                                    datossalida[4] = Convert.ToByte(TablaProgramacion.Rows[i].Cells[4].Value);
                                    datossalida[5] = Convert.ToByte(TablaProgramacion.Rows[i].Cells[5].Value);
                                    datossalida[6] = Convert.ToByte(TablaProgramacion.Rows[i].Cells[6].Value);
                                    datossalida[7] = Convert.ToByte(TablaProgramacion.Rows[i].Cells[7].Value);
                                    datossalida[8] = 254;
                                    PuertoSerial.DiscardOutBuffer();
                                    PuertoSerial.Write(datossalida, 0, 9);
                                    TablaProgramacion.Rows[i].DefaultCellStyle.BackColor = Color.Lime;
                                    controlcolor++;
                                    j++;
                                }
                                if (i > n && ciclado == 1)
                                {
                                    TablaProgramacion.Rows[i].DefaultCellStyle.BackColor = SystemColors.Window;
                                    TablaProgramacion.Rows[0].DefaultCellStyle.BackColor = Color.Lime;
                                    controlcolor = 0;
                                    i = 0;
                                    j = 0;
                                    datossalida[0] = 253;
                                    datossalida[1] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[1].Value);
                                    datossalida[2] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[2].Value);
                                    datossalida[3] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[3].Value);
                                    datossalida[4] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[4].Value);
                                    datossalida[5] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[5].Value);
                                    datossalida[6] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[6].Value);
                                    datossalida[7] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[7].Value);
                                    datossalida[8] = 254;
                                    PuertoSerial.DiscardOutBuffer();
                                    PuertoSerial.Write(datossalida, 0, 9);
                                }
                                if (i > n && ciclado == 0)
                                {
                                    TablaProgramacion.Rows[i].DefaultCellStyle.BackColor = SystemColors.Window;
                                    TablaProgramacion.Rows[n].DefaultCellStyle.BackColor = Color.Lime;
                                    controlcolor = n;
                                    auto = 0;
                                    i = 0;
                                    j = 0;
                                    BotonAuto.Enabled = true;
                                    BotonDetener.Enabled = false;
                                    BotonInsertar.Enabled = true;
                                    BotonInsertarNormal.Enabled = true;
                                    BotonBorrar.Enabled = true;
                                    BotonArriba.Enabled = true;
                                    BotonAbajo.Enabled = true;
                                    BotonSolo.Enabled = true;
                                    BotonActiva.Enabled = true;
                                    BotonActiva.BackColor = Color.Lime;
                                    BotonActiva.Text = "Encender Programación Guiada";
                                    BotonLoop.Enabled = true;
                                    NumericBase.Enabled = true;
                                    NumericA1.Enabled = true;
                                    NumericA2.Enabled = true;
                                    NumericA3.Enabled = true;
                                    NumericA4.Enabled = true;
                                    NumericP.Enabled = true;
                                    NumericV.Enabled = true;
                                    BotonGuardar.Enabled = true;
                                    BotonCargar.Enabled = true;
                                }
                            }                           
                        }
                    } 
                }
                if(sola == 1)
                {
                    if (data.Length == 9)
                    {
                        datossalida[0] = 253;
                        datossalida[1] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[1].Value);
                        datossalida[2] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[2].Value);
                        datossalida[3] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[3].Value);
                        datossalida[4] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[4].Value);
                        datossalida[5] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[5].Value);
                        datossalida[6] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[6].Value);
                        datossalida[7] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[7].Value);
                        datossalida[8] = 254;

                        int contador2 = 0;

                        foreach (var elemento in data)
                        {
                            if (datossalida.Contains(elemento))
                            {
                                contador2++;
                            }
                        }

                        if (contador2 == 9)
                        {
                            TablaProgramacion.Rows[controlcolor].DefaultCellStyle.BackColor = Color.Lime;
                            sola = 0;
                            BotonAuto.Enabled = true;
                            BotonInsertar.Enabled = true;
                            BotonInsertarNormal.Enabled = true;
                            BotonBorrar.Enabled = true;
                            BotonArriba.Enabled = true;
                            BotonAbajo.Enabled = true;
                            BotonSolo.Enabled = true;
                            BotonActiva.Enabled = true;
                            BotonActiva.BackColor = Color.Lime;
                            BotonActiva.Text = "Encender Programación Guiada";
                            BotonLoop.Enabled = true;
                            NumericBase.Enabled = true;
                            NumericA1.Enabled = true;
                            NumericA2.Enabled = true;
                            NumericA3.Enabled = true;
                            NumericA4.Enabled = true;
                            NumericP.Enabled = true;
                            NumericV.Enabled = true;
                            BotonGuardar.Enabled = true;
                            BotonCargar.Enabled = true;
                        }
                    }
                }
                if(activa == 1)
                {
                    datossalida[0] = 253;
                    datossalida[1] = Convert.ToByte(NumericBase.Value);
                    datossalida[2] = Convert.ToByte(NumericA1.Value);
                    datossalida[3] = Convert.ToByte(NumericA2.Value);
                    datossalida[4] = Convert.ToByte(NumericA3.Value);
                    datossalida[5] = Convert.ToByte(NumericA4.Value);
                    datossalida[6] = Convert.ToByte(NumericP.Value);
                    datossalida[7] = Convert.ToByte(NumericV.Value);
                    datossalida[8] = 254;
                    PuertoSerial.DiscardOutBuffer();
                    PuertoSerial.Write(datossalida, 0, 9);
                }
                if(ciclado == 1)
                {
                    auto = 1;
                    DateTime tiempoactual = DateTime.Now;
                    long elapsedTicks = tiempoactual.Ticks - tiempoinicial.Ticks;
                    TimeSpan elapsedSpan = new TimeSpan(elapsedTicks);
                    tiempotranscurrido = Math.Truncate(Convert.ToDecimal(elapsedSpan.TotalSeconds));
                    tiempomuestra = tiempofinal - tiempotranscurrido;

                    if (tiempotranscurrido > tiempofinal)
                    {
                        ciclado = 0;
                        NumericHoras.Value = tiempohoras;
                        NumericMinutos.Value = tiempominutos;
                        NumericSegundos.Value = tiemposegundos;
                        BotonLoop.Enabled = true;
                        NumericHoras.Enabled = true;
                        NumericMinutos.Enabled = true;
                        NumericSegundos.Enabled = true;
                        TablaProgramacion.Rows[i].DefaultCellStyle.BackColor = SystemColors.Window;
                        TablaProgramacion.Rows[n].DefaultCellStyle.BackColor = Color.Lime;
                        controlcolor = n;
                        auto = 0;
                        i = 0;
                        j = 0;
                        BotonAuto.Enabled = true;
                        BotonInsertar.Enabled = true;
                        BotonInsertarNormal.Enabled = true;
                        BotonDetener.Enabled = false;
                        BotonBorrar.Enabled = true;
                        BotonArriba.Enabled = true;
                        BotonAbajo.Enabled = true;
                        BotonSolo.Enabled = true;
                        BotonActiva.Enabled = true;
                        BotonActiva.BackColor = Color.Lime;
                        BotonActiva.Text = "Encender Programación Guiada";
                        NumericBase.Enabled = true;
                        NumericA1.Enabled = true;
                        NumericA2.Enabled = true;
                        NumericA3.Enabled = true;
                        NumericA4.Enabled = true;
                        NumericP.Enabled = true;
                        NumericV.Enabled = true;
                        BotonGuardar.Enabled = true;
                        BotonCargar.Enabled = true;
                    }
                    else
                    {
                        a = Math.Truncate(tiempomuestra / 3600);
                        b = Math.Truncate(((tiempomuestra / 3600) - Math.Truncate(tiempomuestra / 3600)) * 60);
                        c = ((((tiempomuestra / 3600) - Math.Truncate(tiempomuestra / 3600)) * 60) - Math.Truncate(((tiempomuestra / 3600) - Math.Truncate(tiempomuestra / 3600)) * 60)) * 60;
                        c = Math.Round(c);
                        if (c > 59)
                        {
                            b ++;
                            c = 0;
                        }
                        NumericHoras.Value = a;
                        NumericMinutos.Value = b;
                        NumericSegundos.Value = c;
                    }
                }
                if (detener == 1)
                {
                    ciclado = 0;
                    detener = 0;
                    NumericHoras.Value = tiempohoras;
                    NumericMinutos.Value = tiempominutos;
                    NumericSegundos.Value = tiemposegundos;
                    BotonLoop.Enabled = true;
                    NumericHoras.Enabled = true;
                    NumericMinutos.Enabled = true;
                    NumericSegundos.Enabled = true;
                    TablaProgramacion.Rows[i].DefaultCellStyle.BackColor = SystemColors.Window;
                    TablaProgramacion.Rows[n].DefaultCellStyle.BackColor = Color.Lime;
                    controlcolor = n;
                    auto = 0;
                    i = 0;
                    j = 0;
                    BotonAuto.Enabled = true;
                    BotonDetener.Enabled = false;
                    BotonInsertar.Enabled = true;
                    BotonInsertarNormal.Enabled = true;
                    BotonBorrar.Enabled = true;
                    BotonArriba.Enabled = true;
                    BotonAbajo.Enabled = true;
                    BotonSolo.Enabled = true;
                    BotonActiva.Enabled = true;
                    BotonActiva.BackColor = Color.Lime;
                    BotonActiva.Text = "Encender Programación Guiada";
                    NumericBase.Enabled = true;
                    NumericA1.Enabled = true;
                    NumericA2.Enabled = true;
                    NumericA3.Enabled = true;
                    NumericA4.Enabled = true;
                    NumericP.Enabled = true;
                    NumericV.Enabled = true;
                    BotonGuardar.Enabled = true;
                    BotonCargar.Enabled = true;
                }
            }
        }

        void processData()
        {
            if (recievedData.Count > 50)
            {
                var packet = Enumerable.Range(0, 50).Select(i => recievedData.Dequeue());
            }
        }

        private void caja06_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            activa = 0;
            espera = 0;
            BotonActiva.BackColor = Color.Silver;
            datossalida[0] = 253;
            datossalida[1] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[1].Value);
            datossalida[2] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[2].Value);
            datossalida[3] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[3].Value);
            datossalida[4] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[4].Value);
            datossalida[5] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[5].Value);
            datossalida[6] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[6].Value);
            datossalida[7] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[7].Value);
            datossalida[8] = 254;
            PuertoSerial.DiscardOutBuffer();
            PuertoSerial.Write(datossalida, 0, 9);
            auto = 1;
            TablaProgramacion.Rows[controlcolor].DefaultCellStyle.BackColor = SystemColors.Window;
            TablaProgramacion.Rows[0].DefaultCellStyle.BackColor = Color.Lime;
            controlcolor = 0;
            BotonAuto.Enabled = false;
            BotonDetener.Enabled = true;
            BotonInsertar.Enabled = false;
            BotonInsertarNormal.Enabled = false;
            BotonBorrar.Enabled = false;
            BotonArriba.Enabled = false;
            BotonAbajo.Enabled = false;
            BotonSolo.Enabled = false;
            BotonActiva.Enabled = false;
            BotonLoop.Enabled = false;
            NumericBase.Enabled = false;
            NumericA1.Enabled = false;
            NumericA2.Enabled = false;
            NumericA3.Enabled = false;
            NumericA4.Enabled = false;
            NumericP.Enabled = false;
            NumericV.Enabled = false;
            BotonGuardar.Enabled = false;
            BotonCargar.Enabled = false;
        }

        private void caja00_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            n = TablaProgramacion.Rows.Add();
            TablaProgramacion.Rows[n].Cells[0].Value = n;
            TablaProgramacion.Rows[n].Cells[1].Value = NumericBase.Value;
            TablaProgramacion.Rows[n].Cells[2].Value = NumericA1.Value;
            TablaProgramacion.Rows[n].Cells[3].Value = NumericA2.Value;
            TablaProgramacion.Rows[n].Cells[4].Value = NumericA3.Value;
            TablaProgramacion.Rows[n].Cells[5].Value = NumericA4.Value;
            TablaProgramacion.Rows[n].Cells[6].Value = NumericP.Value;
            TablaProgramacion.Rows[n].Cells[7].Value = NumericV.Value;
            if (n > 0)
            {
                TablaProgramacion.Rows[n - 1].DefaultCellStyle.BackColor = SystemColors.Window;
                TablaProgramacion.Rows[controlcolor].DefaultCellStyle.BackColor = SystemColors.Window;
            }
            TablaProgramacion.Rows[n].DefaultCellStyle.BackColor = Color.Lime;
            controlcolor = n;
            y++;
            vacio = 0;
            BotonBorrar.Enabled = true;
            BotonAuto.Enabled = true;
            BotonSolo.Enabled = true;
            BotonArriba.Enabled = true;
            BotonAbajo.Enabled = true;
            BotonLoop.Enabled = true;
        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (n >= 0)
            {
                TablaProgramacion.Rows.RemoveAt(controlcolor);
                n--;
                if(n < 0)
                {
                    n = 0;
                    y = 0;
                    BotonBorrar.Enabled = false;
                    BotonAuto.Enabled = false;
                    BotonSolo.Enabled = false;
                    BotonArriba.Enabled = false;
                    BotonAbajo.Enabled = false;
                    BotonLoop.Enabled = false;
                    vacio = 1;
                }
                if (controlcolor > n)
                {
                    controlcolor = n;
                }
                if(vacio == 0)
                {
                    for (x = 0; x <= n; x++)
                    {
                        TablaProgramacion.Rows[x].Cells[0].Value = x;
                    }
                    TablaProgramacion.Rows[controlcolor].DefaultCellStyle.BackColor = Color.Lime;
                }else if (vacio == 1)
                {
                    TablaProgramacion.Rows[controlcolor].DefaultCellStyle.BackColor = SystemColors.Window;
                }
                
            }
            

        }

        private void button4_Click(object sender, EventArgs e)
        {
            activa = 0;
            BotonActiva.BackColor = Color.Silver;
            datossalida[0] = 253;
            datossalida[1] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[1].Value);
            datossalida[2] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[2].Value);
            datossalida[3] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[3].Value);
            datossalida[4] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[4].Value);
            datossalida[5] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[5].Value);
            datossalida[6] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[6].Value);
            datossalida[7] = Convert.ToByte(TablaProgramacion.Rows[controlcolor].Cells[7].Value);
            datossalida[8] = 254;
            PuertoSerial.DiscardOutBuffer();
            PuertoSerial.Write(datossalida, 0, 9);
            sola = 1;
            TablaProgramacion.Rows[controlcolor].DefaultCellStyle.BackColor = Color.Crimson;
            BotonAuto.Enabled = false;
            BotonInsertar.Enabled = false;
            BotonInsertarNormal.Enabled = false;
            BotonBorrar.Enabled = false;
            BotonArriba.Enabled = false;
            BotonAbajo.Enabled = false;
            BotonSolo.Enabled = false;
            BotonActiva.Enabled = false;
            BotonLoop.Enabled = false;
            NumericBase.Enabled = false;
            NumericA1.Enabled = false;
            NumericA2.Enabled = false;
            NumericA3.Enabled = false;
            NumericA4.Enabled = false;
            NumericP.Enabled = false;
            NumericV.Enabled = false;
            BotonGuardar.Enabled = false;
            BotonCargar.Enabled = false;
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            controlcolor--;
            if (controlcolor < 0)
            {
                controlcolor = n;
            }
            if (controlcolor < n)
            {
                TablaProgramacion.Rows[controlcolor + 1].DefaultCellStyle.BackColor = SystemColors.Window;
            }
            else if (controlcolor == n && controlcolor != 0)
            {
                TablaProgramacion.Rows[0].DefaultCellStyle.BackColor = SystemColors.Window;
            }
            else if (controlcolor == 0)
            {
                TablaProgramacion.Rows[n].DefaultCellStyle.BackColor = SystemColors.Window;
            }
            TablaProgramacion.Rows[controlcolor].DefaultCellStyle.BackColor = Color.Lime;
        }

        private void BotonAbajo_Click(object sender, EventArgs e)
        {
            controlcolor++;
            if (controlcolor > n)
            {
                controlcolor = 0;
            }
            if (controlcolor < n && controlcolor != 0)
            {
                TablaProgramacion.Rows[controlcolor - 1].DefaultCellStyle.BackColor = SystemColors.Window;
            }
            else if(controlcolor == n && controlcolor != 0)
            {
                TablaProgramacion.Rows[n - 1].DefaultCellStyle.BackColor = SystemColors.Window;
            }
            else if (controlcolor == 0)
            {
                TablaProgramacion.Rows[n].DefaultCellStyle.BackColor = SystemColors.Window;
            }
            TablaProgramacion.Rows[controlcolor].DefaultCellStyle.BackColor = Color.Lime;
        }

        private void button1_Click_3(object sender, EventArgs e)
        {
            if (n == 0)
            {
                y++;
            }
            if(n >= 0)
            {
                TablaProgramacion.Rows.Insert(controlcolor, 1);
                TablaProgramacion.Rows[controlcolor].Cells[0].Value = controlcolor;
                TablaProgramacion.Rows[controlcolor].Cells[1].Value = NumericBase.Value;
                TablaProgramacion.Rows[controlcolor].Cells[2].Value = NumericA1.Value;
                TablaProgramacion.Rows[controlcolor].Cells[3].Value = NumericA2.Value;
                TablaProgramacion.Rows[controlcolor].Cells[4].Value = NumericA3.Value;
                TablaProgramacion.Rows[controlcolor].Cells[5].Value = NumericA4.Value;
                TablaProgramacion.Rows[controlcolor].Cells[6].Value = NumericP.Value;
                TablaProgramacion.Rows[controlcolor].Cells[7].Value = NumericV.Value;
                if (controlcolor >= 0 && y >= 2)
                {
                    TablaProgramacion.Rows[controlcolor + 1].DefaultCellStyle.BackColor = SystemColors.Window;
                    n++;
                    for (x = 0; x <= n; x++)
                    {
                        TablaProgramacion.Rows[x].Cells[0].Value = x;
                    }
                }
                vacio = 0;
                TablaProgramacion.Rows[controlcolor].DefaultCellStyle.BackColor = Color.Lime;
                BotonBorrar.Enabled = true;
                BotonAuto.Enabled = true;
                BotonSolo.Enabled = true;
                BotonArriba.Enabled = true;
                BotonAbajo.Enabled = true;
                BotonLoop.Enabled = true;
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            cargar.InitialDirectory = @"C:\Users\SDMG\Downloads";
            cargar.Filter = "Todos los archivos (*.*)|*.*";
            cargar.FilterIndex = 1;
            cargar.RestoreDirectory = true;
            cargar.DefaultExt = "xlsx";
            if (cargar.ShowDialog() == DialogResult.OK)
            {
                string path = cargar.FileName;
                try
                {
                    SLDocument sl = new SLDocument(path);

                    int iRow = 2;
                    int g = 0;

                    if (n > 0)
                    {
                        while (n >= 0)
                        {
                            TablaProgramacion.Rows.RemoveAt(n);
                            n--;
                        }
                        n = 0;
                    }
                    
                    while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
                    {
                        n = TablaProgramacion.Rows.Add();
                        TablaProgramacion.Rows[g].Cells[0].Value = sl.GetCellValueAsInt32(iRow, 1);
                        TablaProgramacion.Rows[g].Cells[1].Value = sl.GetCellValueAsInt32(iRow, 2);
                        TablaProgramacion.Rows[g].Cells[2].Value = sl.GetCellValueAsInt32(iRow, 3);
                        TablaProgramacion.Rows[g].Cells[3].Value = sl.GetCellValueAsInt32(iRow, 4);
                        TablaProgramacion.Rows[g].Cells[4].Value = sl.GetCellValueAsInt32(iRow, 5);
                        TablaProgramacion.Rows[g].Cells[5].Value = sl.GetCellValueAsInt32(iRow, 6);
                        TablaProgramacion.Rows[g].Cells[6].Value = sl.GetCellValueAsInt32(iRow, 7);
                        TablaProgramacion.Rows[g].Cells[7].Value = sl.GetCellValueAsInt32(iRow, 8);
                        g++;
                        y++;
                        iRow++;
                        activo = 1;
                    }

                    if (activo == 1)
                    {
                        TablaProgramacion.Rows[n].DefaultCellStyle.BackColor = Color.Lime;
                        controlcolor = n;
                        vacio = 0;
                        BotonBorrar.Enabled = true;
                        BotonAuto.Enabled = true;
                        BotonSolo.Enabled = true;
                        BotonArriba.Enabled = true;
                        BotonAbajo.Enabled = true;
                        BotonLoop.Enabled = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void BotonGuardar_Click(object sender, EventArgs e)
        {
            SLDocument sl = new SLDocument();

            int iC = 1;
            for (int f = 0; f <= 7; f++)
            {
                sl.SetCellValue(1, iC, TablaProgramacion.Columns[f].HeaderText);
                iC++;
            }

            int iR = 2;
            for (int d = 0; d <= n; d++)
            {
                sl.SetCellValue(iR, 1, Convert.ToInt32(TablaProgramacion.Rows[d].Cells[0].Value));
                sl.SetCellValue(iR, 2, Convert.ToInt32(TablaProgramacion.Rows[d].Cells[1].Value));
                sl.SetCellValue(iR, 3, Convert.ToInt32(TablaProgramacion.Rows[d].Cells[2].Value));
                sl.SetCellValue(iR, 4, Convert.ToInt32(TablaProgramacion.Rows[d].Cells[3].Value));
                sl.SetCellValue(iR, 5, Convert.ToInt32(TablaProgramacion.Rows[d].Cells[4].Value));
                sl.SetCellValue(iR, 6, Convert.ToInt32(TablaProgramacion.Rows[d].Cells[5].Value));
                sl.SetCellValue(iR, 7, Convert.ToInt32(TablaProgramacion.Rows[d].Cells[6].Value));
                sl.SetCellValue(iR, 8, Convert.ToInt32(TablaProgramacion.Rows[d].Cells[7].Value));

                iR++;
            }

            guardado.InitialDirectory = @"C:\Users\SDMG\Downloads";
            guardado.Filter = "Todos los archivos (*.*)|*.*";
            guardado.FilterIndex = 1;
            guardado.RestoreDirectory = true;
            guardado.Title = "Guardar archivo";
            guardado.CheckPathExists = true;
            guardado.DefaultExt = "xlsx";
            if (guardado.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    sl.SaveAs(guardado.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void guardado_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click_4(object sender, EventArgs e)
        {
            if (BotonActiva.BackColor == Color.Lime)
            {
                activa = 1;
                BotonActiva.BackColor = Color.Crimson;
                BotonActiva.Text = "Apagar Programación Guiada";
            }
            else if (BotonActiva.BackColor == Color.Crimson)
            {
                activa = 0;
                BotonActiva.BackColor = Color.Lime;
                BotonActiva.Text = "Encender Programación Guiada";
            }
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void HoraLabel_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_5(object sender, EventArgs e)
        {
            ciclado = 1;
            espera = 0;
            tiempoinicial = DateTime.Now;
            tiempohoras = NumericHoras.Value;
            tiempominutos = NumericMinutos.Value;
            tiemposegundos = NumericSegundos.Value;
            tiempofinal = (tiempohoras*3600)+(tiempominutos*60)+tiemposegundos;
            auto = 1;
            BotonActiva.BackColor = Color.Silver;
            datossalida[0] = 253;
            datossalida[1] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[1].Value);
            datossalida[2] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[2].Value);
            datossalida[3] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[3].Value);
            datossalida[4] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[4].Value);
            datossalida[5] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[5].Value);
            datossalida[6] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[6].Value);
            datossalida[7] = Convert.ToByte(TablaProgramacion.Rows[0].Cells[7].Value);
            datossalida[8] = 254;
            PuertoSerial.DiscardOutBuffer();
            PuertoSerial.Write(datossalida, 0, 9);
            auto = 1;
            TablaProgramacion.Rows[controlcolor].DefaultCellStyle.BackColor = SystemColors.Window;
            TablaProgramacion.Rows[0].DefaultCellStyle.BackColor = Color.Lime;
            controlcolor = 0;
            BotonAuto.Enabled = false;
            BotonDetener.Enabled = true;
            BotonInsertar.Enabled = false;
            BotonInsertarNormal.Enabled = false;
            BotonBorrar.Enabled = false;
            BotonArriba.Enabled = false;
            BotonAbajo.Enabled = false;
            BotonSolo.Enabled = false;
            BotonActiva.Enabled = false;
            BotonLoop.Enabled = false;
            NumericHoras.Enabled = false;
            NumericMinutos.Enabled = false;
            NumericSegundos.Enabled = false;
            NumericBase.Enabled = false;
            NumericA1.Enabled = false;
            NumericA2.Enabled = false;
            NumericA3.Enabled = false;
            NumericA4.Enabled = false;
            NumericP.Enabled = false;
            NumericV.Enabled = false;
            BotonGuardar.Enabled = false;
            BotonCargar.Enabled = false;

        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            NumericBase.Value = trackBar1.Value;
        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            NumericA1.Value = trackBar2.Value;
        }

        private void trackBar3_Scroll(object sender, EventArgs e)
        {
            NumericA2.Value = trackBar3.Value;
        }

        private void trackBar4_Scroll(object sender, EventArgs e)
        {
            NumericA3.Value = trackBar4.Value;
        }

        private void trackBar5_Scroll(object sender, EventArgs e)
        {
            NumericA4.Value = trackBar5.Value;
        }

        private void trackBar6_Scroll(object sender, EventArgs e)
        {
            NumericP.Value = trackBar6.Value;
        }

        private void trackBar7_Scroll(object sender, EventArgs e)
        {
            NumericV.Value = trackBar7.Value;
        }

        private void NumericBase_ValueChanged(object sender, EventArgs e)
        {
            trackBar1.Value = Convert.ToInt32(NumericBase.Value);
        }

        private void NumericA1_ValueChanged(object sender, EventArgs e)
        {
            trackBar2.Value = Convert.ToInt32(NumericA1.Value);
        }

        private void NumericA2_ValueChanged(object sender, EventArgs e)
        {
            trackBar3.Value = Convert.ToInt32(NumericA2.Value);
        }

        private void NumericA3_ValueChanged(object sender, EventArgs e)
        {
            trackBar4.Value = Convert.ToInt32(NumericA3.Value);
        }

        private void NumericA4_ValueChanged(object sender, EventArgs e)
        {
            trackBar5.Value = Convert.ToInt32(NumericA4.Value);
        }

        private void NumericP_ValueChanged(object sender, EventArgs e)
        {
            trackBar6.Value = Convert.ToInt32(NumericP.Value);
        }

        private void NumericV_ValueChanged(object sender, EventArgs e)
        {
            trackBar7.Value = Convert.ToInt32(NumericV.Value);
        }

        private void button1_Click_6(object sender, EventArgs e)
        {
            detener = 1;
        }
    }
}
