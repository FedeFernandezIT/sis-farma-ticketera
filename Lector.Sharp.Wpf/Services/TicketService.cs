using Lector.Sharp.Wpf.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Lector.Sharp.Wpf.Helpers;

namespace Lector.Sharp.Wpf.Services
{    

    public class TicketService
    {        
        public string TicketTerminal { get; set; }
        public string TicketServer { get; set; }
        public string TicketDatabase { get; set; }
        
        public void InitializeConfiguration()
        {
            try
            {
                //var dir = ConfigurationManager.AppSettings["Directory.Setup"];

                //var pathTicketTerminal = Path.Combine(dir, ConfigurationManager.AppSettings["Ticket.Terminal"]);
                //TicketTerminal = new StreamReader(pathTicketTerminal).ReadLine();

                //var pathTicketServer = Path.Combine(dir, ConfigurationManager.AppSettings["Ticket.Server"]);
                //TicketServer = new StreamReader(pathTicketServer).ReadLine();

                //var pathTicketDatabase = Path.Combine(dir, ConfigurationManager.AppSettings["Ticket.Database"]);
                //TicketDatabase = new StreamReader(pathTicketDatabase).ReadLine();
            }
            catch (IOException ex)
            {
                throw new IOException("Error al leer archivos de configuración");
            }
        }

        public void SetTicketsPrinted()
        {
            try
            {
                using (var db = new TicketContext(TicketServer, TicketDatabase))
                {
                    var sql = @"update turnos set ticketImpreso ='1' WHERE 1";
                    db.Database.ExecuteSqlCommand(sql);
                }
            }
            catch (Exception)
            {
                Task.Delay(500).Wait();
            }            
        }
        
        public void PrintInside()
        {
            const string asterics = "*****************************************************";
            const string line = "_____________________________________________________";
            const string please = "Por favor espere su turno";
            const string lblNumber = "SU NÚMERO ES";
            const string lblDate = "Fecha:";

            try
            {                                
                using (var db = new TicketContext(TicketServer, TicketDatabase))
                {
                    using (var trans = db.Database.BeginTransaction(System.Data.IsolationLevel.ReadCommitted))
                    {
                        try
                        {
                            var sql = @"SELECT t.idturno AS IdTurno, t.numero AS Numero, t.fechaTurno AS Fecha, c.letra AS Letra, c.nombre AS Tipo From turnos t INNER JOIN colas c ON c.idcola = t.idcola WHERE  t.idterminal = @terminal AND t.ticketImpreso = '0' Order by t.idturno DESC";
                            var result = db.Database.SqlQuery<Turno>(sql,
                                new MySqlParameter("terminal", TicketTerminal))
                                .ToList();
                            if (result.Count > 0)
                            {
                                var turnoId = result[0].IdTurno;
                                sql = @"update turnos set ticketImpreso ='1' WHERE idturno = @turno";
                                db.Database.ExecuteSqlCommand(sql,
                                    new MySqlParameter("turno", turnoId));
                                sql = @"SELECT texto From ticket WHERE visible = '1'  Order by orden ASC";
                                result[0].Textos = db.Database.SqlQuery<string>(sql).ToList();
                                trans.Commit();

                                var turno = result[0];
                                var printer = new RawPrinter();
                                printer.SetAlign(RawPrinter.Align.Center);
                                printer.DrawLine(RawPrinter.SizeH1 + asterics);
                                printer.DrawLine(RawPrinter.SizeH2 + turno.IdTurno);
                                printer.DrawLine(RawPrinter.SizeH2 + lblNumber);
                                printer.DrawLine(RawPrinter.SizeH4 + turno.Letra + turno.Numero);
                                printer.DrawLine(RawPrinter.SizeH4 + turno.Letra + turno.Numero);
                                printer.DrawLine(asterics);
                                printer.DrawLine(RawPrinter.SizeH1 + asterics);
                                printer.DrawLine(RawPrinter.SizeH1 + please);
                                printer.DrawLine(RawPrinter.SizeH1 + lblDate + " " + turno.Fecha);
                                printer.DrawLine(RawPrinter.SizeH1 + line);
                                foreach (var txt in turno.Textos)
                                {
                                    printer.DrawLine(RawPrinter.SizeH1 + txt);
                                }

                                printer.BarCodeModelA("14159265");
                                printer.BarCodeModelB("14159265");
                                printer.Print();                                
                            }
                            trans.Rollback();
                        }
                        catch (Exception)
                        {
                            trans.Rollback();
                        }
                    }
                }                
            }
            catch (Exception)
            {
                Task.Delay(500).Wait();
            }            
        }
        
        private void PrintPage(object sender, PrintPageEventArgs e, Turno turno)
        {                        
            var currentY = 0f;
            var text = "*****************************************************";
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);            

            text = turno.IdTurno;
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 20f, FontStyle.Bold), StringAlignment.Center, ref currentY);            

            text = "SU NÚMERO ES";
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 15f, FontStyle.Bold), StringAlignment.Center, ref currentY);            

            text = turno.Letra + turno.Numero;
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 50f, FontStyle.Bold), StringAlignment.Center, ref currentY);

            text = "*****************************************************";
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);

            text = "Por favor espere su turno";
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);
            
            text = "Fecha: " + turno.Fecha;
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);
            
            text = "_____________________________________________________";
            PrintLine(text, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);            

            foreach (var item in turno.Textos)
            {
                PrintLine(item, e, new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold), StringAlignment.Center, ref currentY);                
            }

            e.HasMorePages = false;
        }

        private void PrintLine(string text, PrintPageEventArgs e, Font font, StringAlignment alignment, ref float currentY)
        {            
            using (var sf = new StringFormat())
            {
                sf.Alignment = alignment;                
                var size = e.Graphics.MeasureString(text, font, e.MarginBounds.Size);
                e.Graphics.DrawString(text, font, Brushes.Black, new PointF(e.MarginBounds.Left + e.MarginBounds.Width / 2, currentY), sf);                
                currentY += font.GetHeight(e.Graphics);
            }
        }        
    }    
}
//Testing above code

