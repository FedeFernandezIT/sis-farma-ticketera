using Lector.Sharp.Wpf.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Deployment.Application;
using Microsoft.Win32;
using System.IO;
using MySql.Data.MySqlClient;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace Lector.Sharp.Wpf
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        const UInt32 SWP_NOSIZE = 0x0001;
        const UInt32 SWP_NOMOVE = 0x0002;
        const UInt32 SWP_SHOWWINDOW = 0x0040;


        [DllImport("user32.dll")]
        static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);


        /// <summary>
        /// Listener que escucha cada vez que se presiona una tecla.       
        /// </summary>
        private LowLevelKeyboardListener _listener;
        private LowLevelWindowsListener _window;

        /// <summary>
        /// Gestiona todos los servicios de SisFarma, como acceso a la
        /// base de datos, lectura de archivos de configuración.
        /// </summary>
        private FarmaService _service;
        private TicketService _ticketService;

        /// <summary>
        /// Almacena el valor de las teclas presionadas, específcamente números.
        /// </summary>
        private string _keyData = string.Empty;

        /// <summary>
        /// Window para mostrar información de la base de datos.
        /// Está ventana emerge después de presionar ENTER y las teclas presionadas
        /// previamente forman un código existente en la base de datos.
        /// </summary>
        private BrowserWindow _infoBrowser;

        /// <summary>
        /// Window para mostrar una dirección web. Esta ventana emerge cuando se
        /// presiona SHIFT+F1 y se cierrra al presionar SHIFT+F2
        /// </summary>
        private BrowserWindow _customBrowser;

        private BrowserWindow _presentationBrowser;
                
        /// <summary>
        /// Icono de barra de tareas, gestiona la salida del programa
        /// </summary>
        private System.Windows.Forms.NotifyIcon _iconNotification;

        private System.Timers.Timer _ticketPrinterTimer;
        private System.Timers.Timer _shutdownTimer;

        /// <summary>
        /// Devuelve una ventana para mostrar info de la base de datos, si esta ya ce cerró devuelve una nueva
        /// </summary>
        public BrowserWindow InfoBrowser
        {
            get
            {
                if (_infoBrowser.IsClosed)                
                    _infoBrowser = new BrowserWindow();                
                return _infoBrowser;
            }
        }

        /// <summary>
        /// Devuelve una ventana para mostrar una web específica, si esta ya ce cerró devuelve una nueva
        /// </summary>
        public BrowserWindow CustomBrowser
        {
            get
            {
                if (_customBrowser.IsClosed)                
                    _customBrowser = new BrowserWindow();                
                return _customBrowser;
            }
        }

        private Process _processChrome;
        
        public MainWindow()
        {
            try
            {                        
                InitializeComponent();
                
                _service = new FarmaService();
                _ticketService = new TicketService();

                _listener = new LowLevelKeyboardListener();
                _window = new LowLevelWindowsListener();

                _infoBrowser = new BrowserWindow();
                _customBrowser = new BrowserWindow();
                _presentationBrowser = new BrowserWindow();                

                // Leemos los archivos de configuración
                _service.LeerFicherosConfiguracion();
                _ticketService.InitializeConfiguration();

                // Setamos el comportamiento de la aplicación al presionar una tecla
                _listener.OnKeyPressed += _listener_OnKeyPressed;

                // Activamos el listener de teclado
                _listener.HookKeyboard();

                _iconNotification = new System.Windows.Forms.NotifyIcon();
                _iconNotification.BalloonTipText = "La Aplicación SisFarma se encuentra ejecutando";
                _iconNotification.BalloonTipTitle = "SisFarma Notificación";
                _iconNotification.Text = "Presione Click para Mostrar";
                _iconNotification.Icon = Lector.Sharp.Wpf.Properties.Resources.Logo;
                _iconNotification.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;

                System.Windows.Forms.ContextMenu menu = new System.Windows.Forms.ContextMenu();
                System.Windows.Forms.MenuItem notificactionInfoMenu = new System.Windows.Forms.MenuItem("Info");
                notificactionInfoMenu.Click += notificactionInfoMenu_Click;
                System.Windows.Forms.MenuItem notificationQuitMenu = new System.Windows.Forms.MenuItem("Salir");
                notificationQuitMenu.Click += notificationQuitMenu_Click;

                menu.MenuItems.Add(notificactionInfoMenu);
                menu.MenuItems.Add(notificationQuitMenu);
                _iconNotification.ContextMenu = menu;
                _iconNotification.Visible = true;

                InitializeTicketTimer();
                InitializeShutdownTimer();

                //OpenWindowPresentation(_presentationBrowser, "http://www.google.com.ar");                

                _processChrome = new Process();                
                _processChrome.StartInfo.FileName = @"chrome.exe";
                _processChrome.StartInfo.Arguments = $@"--kiosk --fullscreen --app={_service.Presentation}";                
                _processChrome.Start();
                _processChrome.WaitForInputIdle(1000);
                _processChrome.Refresh();
                SetParent(_processChrome.MainWindowHandle, new WindowInteropHelper(this).Handle);                
                SetWindowPos(_processChrome.MainWindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
                SetForegroundWindow(_processChrome.Handle);
                _processChrome.Refresh();
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
            
        }

        private void InitializeTicketTimer()
        {
            _ticketService.SetTicketsPrinted();
            _ticketPrinterTimer = new System.Timers.Timer(1000);
            //_ticketPrinterTimer.AutoReset = false;
            _ticketPrinterTimer.Elapsed += (o, e) =>
            {                
                Task.Run(() => _ticketService.PrintInside());                
            };
            _ticketPrinterTimer.Start();
        }

        private void InitializeShutdownTimer()
        {
            _shutdownTimer = new System.Timers.Timer(60000);
            _shutdownTimer.Elapsed += (o, e) => _service.CheckShutdownTime();
            _shutdownTimer.Start();
        }

                        
        /// <summary>
        /// Acción para salir del programa.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void notificationQuitMenu_Click(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void notificactionInfoMenu_Click(object sender, EventArgs e)
        {
            MessageBox.Show("SisFarma Applicación\nsisfarma.es");
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Desactivamos el listener del teclado
            _listener.UnHookKeyboard();
        }

        /// <summary>
        /// Precesa el comporatiento de la aplicación al presionarse una tecla.
        /// </summary>
        /// <param name="sender">Listener del teclado</param>
        /// <param name="e">Información de la tecla presionada</param>
        private void _listener_OnKeyPressed(object sender, KeyPressedArgs e)
        {
            try
            {
                if (e.KeyPressed != Key.Enter)
                {
                    // Si presionamos SHIFT + F1
                    if (_listener.IsHardwareKeyDown(LowLevelKeyboardListener.VirtualKeyStates.VK_SHIFT) && e.KeyPressed == Key.F1)
                    {   // Si La ventana de información detallada está abierta la cerramos
                        if (InfoBrowser.IsVisible)
                            CloseWindowBrowser(InfoBrowser);
                        // Abrimos una ventana con la web personalizada.    
                        OpenWindowBrowser(CustomBrowser, _service.UrlNavegarCustom, InfoBrowser);
                    }
                    // Si presionamos SHIFT + F2
                    else if (_listener.IsHardwareKeyDown(LowLevelKeyboardListener.VirtualKeyStates.VK_SHIFT) && e.KeyPressed == Key.F2)
                    {
                        if (_processChrome != null && !_processChrome.HasExited)
                        {
                            _processChrome.CloseMainWindow();
                            _processChrome.Close();
                            _processChrome = null;
                        }

                        // Cerramos la ventana con la web personalizada
                        CloseWindowBrowser(CustomBrowser);
                    }
                    // Si la tecla presioanada es numérica
                    else if (!_listener.IsHardwareKeyDown(LowLevelKeyboardListener.VirtualKeyStates.VK_SHIFT) &&
                        (e.KeyPressed >= Key.D0 && e.KeyPressed <= Key.D9 || e.KeyPressed >= Key.NumPad0 && e.KeyPressed <= Key.NumPad9))
                    {
                        // Almacenamos el valor de la tecla.
                        StoreKey(e.KeyPressed);
                    }
                }
                else if (!CustomBrowser.IsVisible && !string.IsNullOrEmpty(_keyData))
                {
                    var enteredNumbers = _keyData;                    
                    Task.Run(() => ProccessEnterKey(enteredNumbers)).ContinueWith(t =>
                    {
                        if (t.Result)
                        {
                            OpenWindowBrowser(InfoBrowser, _service.UrlNavegar, CustomBrowser);
                        }
                    }, TaskScheduler.FromCurrentSynchronizationContext());

                    // Limpiamos _keyData para otro proceso.
                    _keyData = string.Empty;
                    SendKeyEnter();
                }                                                        
                else
                {
                    // Siempre que se presiona ENTER se limpia _keyData
                    _keyData = string.Empty;
                }
            }
            catch (MySqlException mysqle)
            {
                //if (!mysqle.Message.Contains("Timeout"))
                //{
                //    MessageBox.Show("Ha ocurrido un error. Comuníquese con el Administrador.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                //}
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ha ocurrido un error. Comuníquese con el Administrador.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }

        /// <summary>
        /// Procesa los _keyData para buscar datos en la base de datos
        /// </summary>
        private bool ProccessEnterKey(string enteredNumbers)
        {
            long number = 0;
            // Si el valor almacenado en _keyData en numérico y con longitud superior a 4

            if (long.TryParse(enteredNumbers, out number) && enteredNumbers.Length >= 4)
            {

                var lanzarBrowserWindow = false;
                var noEntrar = string.Empty;
                var continuar = false;
                var continuarCodNacional = false;
                var enteredNumbersAux = enteredNumbers;

                if (enteredNumbers.Length >= 13)
                {
                    enteredNumbers = enteredNumbers.Substring(enteredNumbers.Length - 13);
                    noEntrar = enteredNumbers.Substring(0, 4);

                    if (Array.Exists(new[] { "1010", "1111", "0000", "9902", "9900", "9901", "9903", "9904", "9905", "9906", "9907", "9908", "9909", "9910", "9911", "9912", "9913", "9915", "9916", "9917", "9918", "9919", "9920", "1001", "2014", "2015", "2016", "2017", "2018" }, x => x.Equals(noEntrar)))
                        continuar = true;
                    else
                    {
                        noEntrar = enteredNumbers.Substring(0, 7);
                        if (noEntrar.Equals("8470000"))
                            continuar = true;
                        else
                        {
                            noEntrar = enteredNumbers.Substring(0, 3);
                            if (Array.Exists(_service.GetCodigoBarraMedicamentos(), x => x.Equals(noEntrar)))
                                continuarCodNacional = true;
                            else if (Array.Exists(_service.GetCodigoBarraSinonimos(), x => x.Equals(noEntrar)))
                                continuarCodNacional = true;
                        }
                    }
                }

                if (enteredNumbers.Length >= 12 && !continuar && !continuarCodNacional)
                {
                    enteredNumbers = enteredNumbersAux.Substring(enteredNumbers.Length - 12);
                    noEntrar = enteredNumbers.Substring(0, 4);
                    if ("1111".Equals(noEntrar) || "0000".Equals(noEntrar))
                        continuar = true;
                }

                if (enteredNumbers.Length >= 10 && !continuar && !continuarCodNacional)
                {
                    enteredNumbers = enteredNumbersAux.Substring(enteredNumbers.Length - 10);
                    noEntrar = enteredNumbers.Substring(0, 4);
                    if ("1111".Equals(noEntrar) || "1930".Equals(noEntrar))
                        continuar = true;
                }

                if (enteredNumbers.Length >= 7 && !continuar && !continuarCodNacional)
                {
                    enteredNumbers = enteredNumbersAux.Substring(enteredNumbers.Length - 7);
                    noEntrar = enteredNumbers.Substring(0, 4);
                    if (Array.Exists(new[] { "1000", "1001", "1002", "1003" }, x => x.Equals(noEntrar)))
                        continuar = true;
                }

                if (enteredNumbers.Length >= 6 && !continuar && !continuarCodNacional)
                {
                    enteredNumbers = enteredNumbersAux.Substring(enteredNumbers.Length - 6);
                    noEntrar = enteredNumbers.Substring(0, 3);
                    if (Array.Exists(new[] { "000", "001", "002", "003", "004", "005", "006", "007", "008", "009", "010", "011", "100", "101", "102", "103" }, x => x.Equals(noEntrar)))
                        continuar = true;
                }

                if (enteredNumbers.Length >= 5 && !continuar && !continuarCodNacional)
                {
                    enteredNumbers = enteredNumbersAux.Substring(enteredNumbers.Length - 5);
                    noEntrar = enteredNumbers.Substring(0, 2);
                    if (Array.Exists(new[] { "00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12" }, x => x.Equals(noEntrar)))
                        continuar = true;
                }

                if (enteredNumbers.Length >= 4 && !continuar && !continuarCodNacional)
                {
                    enteredNumbers = enteredNumbersAux.Substring(enteredNumbers.Length - 4);
                    noEntrar = enteredNumbers.Substring(0, 2);
                    if (Array.Exists(new[] { "00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12" }, x => x.Equals(noEntrar)))
                        continuar = true;
                }

                if (continuar)
                {
                    var cliente = _service.GetCliente(enteredNumbers);
                    if (cliente != null)
                    {                     
                        lanzarBrowserWindow = true;
                        _service.UrlNavegar = _service.Url.Replace("codigo", cliente.ToString()) + "/" + _service.Mostrador;
                    }
                    else
                    {
                        var trabajador = _service.GetTrabajador(enteredNumbers);
                        if (trabajador != null)
                        {                     
                            lanzarBrowserWindow = true;
                            _service.UrlNavegar = _service.UrlMensajes.Replace("codigo", trabajador.ToString());
                        }
                    }
                }
                else
                {
                    if (continuarCodNacional)
                    {
                        string foundAsociado = null;
                        var codNacional = _service.GetCodigoNacionalSinonimo(enteredNumbers.Substring(0, enteredNumbers.Length > 12 ? 12 : enteredNumbers.Length));
                        if (codNacional == null)
                        {
                            codNacional = _service.GetCodigoNacionalMedicamento(enteredNumbers.Substring(0, enteredNumbers.Length > 12 ? 12 : enteredNumbers.Length));
                            if (codNacional == null)
                                codNacional = Convert.ToInt64(enteredNumbers.Substring(3, enteredNumbers.Length - 4));
                        }

                        var asociado = _service.GetAsociado(Convert.ToInt64(codNacional));
                        var mostrarVentana = false;
                        if (asociado != null)
                        {
                            mostrarVentana = true;
                            foundAsociado = asociado;
                        }
                        else
                        {
                            var articulo = _service.GetArticulo(Convert.ToInt64(codNacional));
                            if (articulo != null)
                            {
                                mostrarVentana = true;
                                foundAsociado = articulo;
                            }                                
                            else
                            {
                                var categ = _service.GetCategorizacion();
                                if (categ != null)
                                {
                                    var asociadoCodNacional = _service.GetAnyAsociadoMedicamento(Convert.ToInt64(codNacional));
                                    if (asociadoCodNacional != null)
                                    {
                                        mostrarVentana = true;
                                        foundAsociado = asociadoCodNacional;
                                    }
                                }                                                                    
                                else
                                {
                                    var asociadoCodNacional = _service.GetAsociadoCategorizacion(Convert.ToInt64(codNacional));                                    
                                    if (asociadoCodNacional != null)
                                    {
                                        mostrarVentana = true;
                                        foundAsociado = asociadoCodNacional;
                                    }
                                }
                            }
                        }

                        if (mostrarVentana && foundAsociado != null)
                        {                                                        
                            lanzarBrowserWindow = true;
                            _service.UrlNavegar = _service.Url.Replace("codigo", "cn" + foundAsociado + "/" + _service.Mostrador);                            
                        }
                    }
                    
                }

                if (lanzarBrowserWindow)
                {
                    // Mostramos el browser con información de la base de datos                                          
                    // Si es proceso de búsqueda en la base de datos es exitoso
                    // mostramos los resultados                                        
                    return true;

                }
            }
            return false;
        }

        /// <summary>
        /// Almacena el valor de la tacla en _keyData
        /// </summary>
        /// <param name="key">Tecla presionada</param>
        private void StoreKey(Key key)
        {
            if (_keyData.Length > 50)
                _keyData = _keyData.Substring(_keyData.Length - 20 - 1);
            var kc = new KeyConverter();
            // Key.NumPad# se convierte en 'NumPad#' por lo cual lo eliminamos
            _keyData += kc.ConvertToString(key)?.Replace("NumPad", string.Empty);
        }

        /// <summary>
        /// Cierra una ventana que contiene un browser
        /// </summary>
        /// <param name="browser"></param>
        private void CloseWindowBrowser(BrowserWindow browser)
        {
            browser.Hide();
        }

        /// <summary>
        /// Abre una ventana que contiene un browser
        /// </summary>
        /// <param name="browser">Ventana con un browser</param>
        private void OpenWindowBrowser(BrowserWindow browser, string url, BrowserWindow hidden)
        {
            
            if (_processChrome != null && !_processChrome.HasExited)
            {
                _processChrome.CloseMainWindow();
                _processChrome.Close();
                _processChrome = null;
            }              
            
            hidden.Topmost = false;
            browser.Topmost = true;
            hidden.Topmost = true;
            browser.Browser.Navigate(url);
            browser.Visibility = Visibility.Visible;
            browser.WindowState = WindowState.Maximized;            
            browser.Show();
            browser.Activate();
        }

        private void OpenWindowPresentation(BrowserWindow browser, string url)
        {
            browser.Topmost = true;
            browser.WindowStyle = WindowStyle.None;
            browser.Browser.Navigate(url);
            browser.Visibility = Visibility.Visible;
            browser.WindowState = WindowState.Maximized;
            browser.Show();
            browser.Activate();
        }

        /// <summary>
        /// Simula presionar la tecla ENTER
        /// </summary>        
        public static void SendKeyEnter()
        {
            // Utilizar SendWait para compatibilidad con WPF
            System.Windows.Forms.SendKeys.SendWait("{ENTER}");            
        }
        
    }
}
