using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.XtraWaitForm;
using DueStudioToolbar.Services;
using DueStudioToolbar.Utils;
using DueStudioToolbar.Vision;
using Emgu.CV;
using Emgu.CV.Structure;
using Gma.System.MouseKeyHook;
using ManagedWinapi.Windows;
using Nancy.Hosting.Self;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using static DueStudioToolbar.Dtos.Connection;

namespace DueStudioToolbar
{
    public partial class frmToolbar : DevExpress.XtraEditors.XtraForm
    {
        private int WindowButtons = 10;
        private IEnumerable<SystemWindow> all_windows;

        private frmAvisoTampa avisoTampa = new frmAvisoTampa();
        private frmAvisoSensores avisoSensores = new frmAvisoSensores();
        private frmRotaryAxisConnected _fRotaryAxisConnected = new frmRotaryAxisConnected();
        private frmRotaryAxisDisconnected _fRotaryAxisDisconnected = new frmRotaryAxisDisconnected();
        private BarCheckItem bufferbutton = null;
        private DateTime TimeToWait = new DateTime();

        //private IKeyboardMouseEvents m_GlobalHook;
        private bool KeyD = false;
        
        private int NumTriesState = 0;

        public bool _restartSearchThreads = false;

        public frmConnections RefToFrmConnections { get; set; }

        private ProgressPanel _p = new ProgressPanel();
        private object _buffer;

        private bool _MultiLines = false;
        private string _MultiLineEnd = "";
        private string LastReceivedMessageBuffer = "";

        private bool _rotaryAxisConnectedShown = false;
        private bool _sensoresError = false;
        private bool _tampaError = false;

        public frmToolbar()
        {
            Visible = false;
            DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle("DevExpress Dark Style");
            InitializeComponent();
            Height = 23;
            Global.fToolbar = this;

            HostConfiguration hostConfigs = new HostConfiguration()
            {
                UrlReservations = new UrlReservations() { CreateAutomatically = true }
            };

            NancyHost nancyHost = new NancyHost(new Uri("http://127.0.0.1:3777/"), new Nancy.DefaultNancyBootstrapper(), hostConfigs);
            nancyHost.Start();

            //SetupKeyboardHooks();

            //readUsrSettings();

            LoadUsrSettings();

            /// Flow Files
            Program.UsbFlowFiles.ValueChanged += () => { FillFilesComboBox(Globals.DueStudio.UsbGcodeFiles); };
            _p.Top = cmbFlowFiles.Top - 1;
            _p.Left = cmbFlowFiles.Left - 1;
            _p.Bounds = cmbFlowFiles.Bounds;
            _p.Caption = "Carregando...";
            _p.AppearanceCaption.ForeColor = Color.White;
            _p.LookAndFeel.UseDefaultLookAndFeel = false;
            this.Controls.Add(_p);
            _p.Visible = false;

            panWorkProgress.Top = cmbFlowFiles.Top;
            panWorkProgress.Left = cmbFlowFiles.Left;

            panStatus.Size = cmbFlowFiles.Size;
            panStatus.Top = cmbFlowFiles.Top;
            panStatus.Left = cmbFlowFiles.Left;
            panStatus.Bounds = cmbFlowFiles.Bounds;
            panStatus.Visible = false;
            lblStatus.Text = "Inicializando...";

            picDueLogo.ToolTip = string.Format("Versão {0}", Globals.DueStudio.Version);
        }

        private void LoadUsrSettings()
        {
            try
            {
                Globals.DueStudio.DueConfiguration.Load();
            }
            catch (Exception Ex)
            {
                Debug.WriteLine(Ex.Message);
            }
        }

        public void OnChangeConnectionState(bool connected)
        {
            if(connected)
            {
                MethodInvoker m = new MethodInvoker(() => btnConnection.ToolTip = DueFlowApiEsp32.Ip);
                btnConnection.Invoke(m);
                m = new MethodInvoker(() => btnConnection.ForeColor = Color.DarkGreen);
                btnConnection.Invoke(m);
                m = new MethodInvoker(() => cmbFlowFiles.Enabled = true);
                cmbFlowFiles.Invoke(m);
                // HidePanSearching();
            } else
            {
                btnConnection.ToolTip = "Desconectado";
                btnConnection.ForeColor = Color.OrangeRed;
                cmbFlowFiles.Enabled = false;
            }
        }

        private void timDueStudioPosition_Tick(object sender, EventArgs e)
        {
            FindDueflowInkWindow();
            if (Program.sys_window != null)
            {
                if (Program.sys_window.WindowState == FormWindowState.Maximized)
                {
                    this.Top = 24;
                    this.Left = Screen.PrimaryScreen.WorkingArea.Width - (this.Width + 2);
                }
                else if (Program.sys_window.WindowState == FormWindowState.Normal)
                {
                    this.Top = Program.sys_window.Position.Top + 31;
                    this.Left = Program.sys_window.Position.Left + Program.sys_window.Position.Width - this.Width - WindowButtons;
                }
                var MouseWin1 = SystemWindow.FromPoint(200, 200);
                var MouseWin2 = SystemWindow.FromPoint(Screen.PrimaryScreen.WorkingArea.Width - 200, 200);
                var MouseWin3 = SystemWindow.FromPoint(Screen.PrimaryScreen.WorkingArea.Width - 200, Screen.PrimaryScreen.WorkingArea.Height - 200);
                var MouseWin4 = SystemWindow.FromPoint(200, Screen.PrimaryScreen.WorkingArea.Height - 200);
                var MouseWin5 = SystemWindow.FromPoint(Screen.PrimaryScreen.WorkingArea.Width / 2, Screen.PrimaryScreen.WorkingArea.Height / 2);

                this.Visible = (MouseWin1.Title == Program.sys_window.Title) ||
                    (MouseWin2.Title == Program.sys_window.Title) || (MouseWin3.Title == Program.sys_window.Title) ||
                    (MouseWin4.Title == Program.sys_window.Title) || (MouseWin5.Title == Program.sys_window.Title);
                this.Height = 28;
            }
            else this.Visible = false;
        }

        private void FindDueflowInkWindow()
        {
            all_windows = SystemWindow.AllToplevelWindows.Where<SystemWindow>(g => g.Visible == true);

            /*foreach (SystemWindow m in all_windows)
            {
                if (m.Title.Contains("- Due Studio 4"))
                {
                    Program.sys_window = m;
                }
            }*/

            if (SystemWindow.ForegroundWindow.Title.Contains("- Due Studio 4"))
                Program.sys_window = SystemWindow.ForegroundWindow;
        }

        public void OpenRadialMenu()
        {
            //if (Visible && (DueFlowApiEsp32.Connected || DueFlowUsb.UsbFound))
            if (DueFlowApiEsp32.Connected || FlowUsbConnection.UsbFound)
                radialMenu1.ShowPopup(new Point(this.Location.X, this.Location.Y + 200));
        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            Program.fConfig.Show();
            Program.fConfig.BringToFront();
        }

        private void timStateConnection_Tick(object sender, EventArgs e)
        {
            GetProgressValue();
            checkFlowState();
            UpdateButtons();
        }

        private void GetProgressValue()
        {
            if (FlowUsbConnection.UsbFound)
            {
                // Streaming
                if (FlowUsbConnection.Core.InProgram)
                {
                    if (FlowUsbConnection.Core.ProgramTarget != 0)
                        Program.PrintProgress = (int)((FlowUsbConnection.Core.ProgramExecuted / (float)FlowUsbConnection.Core.ProgramTarget) * 100.0f);
                    else
                        Program.PrintProgress = 0;
                }
                else
                {
                    // SD File Job
                    Program.PrintProgress = (int)FlowUsbConnection.Core.ProgressSD;
                }
                    
            }
        }

        private void UpdateButtons()
        {
            btnCameraGetInk.Enabled = (CameraApi.CameraFound);
        }

        private void timCheckUploadedImageApp_Tick(object sender, EventArgs e)
        {
            // TODO: remover
        }

        private void DisplayMachineConnectionState(bool connected, bool usbOrWifi = false, string connectedMachine = "", string ip = "", string portName = "", string ssid = "", string version = "")
        {
            if (connected)
            {
                if (usbOrWifi)
                {
                    string toolTip;
                    if (ip == "")
                        toolTip = string.Format("{0}Versão: {1}\r\nConectado via USB na porta {2}", connectedMachine, version, portName);
                    else
                        toolTip = string.Format("{0}Versão: {1}\r\nConectado via USB na porta {2}\r\nRede: {3}\r\nIP: {4}", connectedMachine, version, portName, ssid, ip);
                    btnConnection.ToolTip = toolTip;
                    btnConnection.ImageOptions.ImageIndex = 0;
                    btnConnection.ForeColor = Color.DarkGreen;
                }
                else
                {
                    btnConnection.ToolTip = string.Format("{0}Versão: {1}\r\nConectado via Wi-Fi\r\nRede: {2}\r\nIP: {3}", connectedMachine, version, ssid, ip);
                    btnConnection.ImageOptions.ImageIndex = 1;
                    btnConnection.ForeColor = Color.DarkGreen;
                }
            }
            else
            {
                btnConnection.ToolTip = "Desconectado";
                btnConnection.ImageOptions.ImageIndex = 0;
                btnConnection.ForeColor = Color.OrangeRed;
            }
        }
        
        private void HandleMachineState(string state)
        {
            try
            {
                string status = state;

                if (state == "Fechado")
                {
                    status = "Conectando...";
                    EnableToolbarButtonGroup(false);
                    cmbFlowFiles.Visible = false;
                    panWorkProgress.Visible = false;
                    panStatus.Visible = true;
                }

                if (state == "Alerta" || state == "Alarme")
                {
                    string driverErrorMsg = CheckForDriverXYError();
                    status = driverErrorMsg != "" ? driverErrorMsg : "Desbloqueando...";
                    EnableToolbarButtonGroup(false);
                    cmbFlowFiles.Visible = false;
                    panWorkProgress.Visible = false;
                    panStatus.Visible = true;

                    DueFlowApiEsp32.Unlock();
                }

                if (state == "Operacional")
                {
                    EnableToolbarButtonGroup(true);
                    cmbFlowFiles.Visible = true;
                    panWorkProgress.Visible = false;
                    panStatus.Visible = false;
                }

                if (state == "Movimentando")
                {
                    status = "Movimentando";
                    EnableToolbarButtonGroup(false);
                    cmbFlowFiles.Visible = false;
                    panWorkProgress.Visible = false;
                    panStatus.Visible = true;
                }

                if (state == "Trabalhando" || state == "Executando")
                {
                    EnableToolbarButtonGroup(false);
                    cmbFlowFiles.Visible = false;
                    panWorkProgress.Visible = true;
                    panStatus.Visible = false;

                    btnPausePlay.ImageOptions.ImageIndex = 3;
                    btnPausePlay.ForeColor = Color.SandyBrown;
                    btnPausePlay.Tag = "pause";

                    MethodInvoker mWork = new MethodInvoker(() => pbWork.EditValue = Program.PrintProgress);
                    pbWork.Invoke(mWork);
                }

                if (state == "Pausado" || state == "Parado")
                {
                    EnableToolbarButtonGroup(false);
                    cmbFlowFiles.Visible = false;
                    panWorkProgress.Visible = true;
                    panStatus.Visible = false;

                    btnPausePlay.ImageOptions.ImageIndex = 4;
                    btnPausePlay.ForeColor = Color.LightSteelBlue;
                    btnPausePlay.Tag = "play";

                    if (!Program.UserPaused)
                    {
                        DueFlowApiEsp32.Resume();
                    }
                }

                if (state == "Tampa de segurança aberta")
                {
                    EnableToolbarButtonGroup(false);
                    cmbFlowFiles.Visible = false;
                    panWorkProgress.Visible = true;
                    panStatus.Visible = true;

                    if (!_tampaError)
                        ShowAvisoTampaErrorMsg();
                    _tampaError = true;
                }
                else
                {
                    if (state != "Alerta")
                    {
                        _tampaError = false;
                        avisoTampa.Hide();
                    }
                }

                if (state == "Tampa de segurança fechada")
                {
                    status = "Retomando...";
                    EnableToolbarButtonGroup(true);
                    cmbFlowFiles.Visible = false;
                    panWorkProgress.Visible = false;
                    panStatus.Visible = true;

                    DueFlowApiEsp32.Resume();
                }

                if (state == "Porta")
                {
                    status = "Tampa de segurança aberta";
                    EnableToolbarButtonGroup(true);
                    cmbFlowFiles.Visible = false;
                    panWorkProgress.Visible = false;
                    panStatus.Visible = true;

                    DueFlowApiEsp32.Resume();
                }

                if (state.StartsWith("Sensores"))
                {
                    status = "Erro nos sensores";
                    EnableToolbarButtonGroup(false);
                    cmbFlowFiles.Visible = false;
                    panWorkProgress.Visible = true;
                    panStatus.Visible = true;

                    if (!_sensoresError)
                        ShowAvisoSensoresErrorMsg();
                    _sensoresError = true;
                }
                else
                {
                    _sensoresError = false;
                    avisoSensores.Hide();
                }
                
                lblStatus.Text = status;

                SetBtnAlertVisibility();
            }
            catch
            {
            }
        }
        
        private void checkFlowState()
        {
            // ESP32
            string connectedMachine;
            if (DueFlowApiEsp32.CurrentMachine() == DueFlowApiEsp32.DueMachine.flow)
            {
                connectedMachine = "Due Flow\r\n";
            }
            else if (DueFlowApiEsp32.CurrentMachine() == DueFlowApiEsp32.DueMachine.nxt)
            {
                connectedMachine = "Due NXT\r\n";
            }

            else if (DueFlowApiEsp32.CurrentMachine() == DueFlowApiEsp32.DueMachine.max)
            {
                connectedMachine = "Due Max\r\n";
            }
            else
            {
                connectedMachine = "";
            }

            if (FlowUsbConnection.UsbFound || DueFlowApiEsp32.Connected)
            {
                if (FlowUsbConnection.UsbFound)
                {
                    //DisplayMachineConnectionState(true, true, connectedMachine, FlowUsbConnection.Ip, FlowUsbConnection.PortName);
                    DisplayMachineConnectionState(true, true, connectedMachine, Globals.DueStudio.MachineInfo.ip, FlowUsbConnection.PortName, Globals.DueStudio.MachineInfo.ssid, Globals.DueStudio.MachineInfo.version);
                }
                else
                {
                    DisplayMachineConnectionState(true, false, connectedMachine, DueFlowApiEsp32.Ip, "", Globals.DueStudio.MachineInfo.ssid, Globals.DueStudio.MachineInfo.version);
                }
                cmbFlowFiles.Enabled = true;
                /*panWorkProgress.Visible = true;
                panStatus.Visible = true;*/
                HandleMachineState(Program.FlowState);
            }
            else
            {
                DisplayMachineConnectionState(false); // commandButtons -> desativado panWorkProgress/panStatus -> invisível
                //cmbFlowFiles.Enabled = false;
                cmbFlowFiles.Enabled = false;
                EnableToolbarButtonGroup(false);
                panWorkProgress.Visible = false;
                panStatus.Visible = false;
            }
        }

        private void EnableToolbarButtonGroup(bool Enabled)
        {
            foreach (var c in Controls)
            {
                if (c.GetType() == typeof(SimpleButton))
                {
                    var ct = c as SimpleButton;
                    if (ct.Tag != null)
                    {
                        if (ct.Tag.ToString() == "GroupToolbar")
                            ct.Enabled = Enabled;
                    }
                }
            }
        }

        private void frmToolbar_FormClosed(object sender, FormClosedEventArgs e)
        {
            //m_GlobalHook.Dispose();
            Program.FinishThreads();
        }

        private void btnRadialMenu_Click(object sender, EventArgs e)
        {
            OpenRadialMenu();
        }

        private void pictureEdit1_DoubleClick(object sender, EventArgs e)
        {
            if (DueFlowApiEsp32.Connected)
            {
                if (!Program.IsAdmin)
                {
                    bool fAberta = General.FocusOnFormIfOpen("frmConfigPass");
                    if (!fAberta)
                    {
                        frmConfigPass fConfigPass = new frmConfigPass();
                        fConfigPass.Show();
                    }
                }
                else
                {
                    bool fAberta = General.FocusOnFormIfOpen("frmConfigIntern");
                    if (!fAberta)
                    {
                        frmConfigIntern fConfigIntern = new frmConfigIntern();
                        fConfigIntern.Show();
                    }
                }
            }
            else
            {
                XtraMessageBox.Show("Sem conexão!");
            }
        }

        private void btnSensors_Click(object sender, EventArgs e)
        {
            bool fAberta = General.FocusOnFormIfOpen("frmFlowSensorsActuators");
            if (!fAberta && (DueFlowApiEsp32.Connected || FlowUsbConnection.UsbFound))
            {
                frmFlowSensorsActuators fFlowCtrl = new frmFlowSensorsActuators();
                fFlowCtrl.Show();
            }
        }

        private void barCheckItem1_CheckedChanged(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // Step
            bufferbutton = sender as BarCheckItem;
            string buf = bufferbutton.Tag.ToString();
            Program.StepRun = float.Parse(buf);
        }

        private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (Program.FlowState != "Operacional")
                return;
            // Up
            if (Program.EixoRotativo)
                DueFlowApiEsp32.SendJogCommandZaxis(0, Program.StepRun);
            else
                DueFlowApiEsp32.SendJogCommand(0, Program.StepRun);
        }

        private void barButtonItem3_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (Program.FlowState != "Operacional")
                return;
            // Right
            DueFlowApiEsp32.SendJogCommand(Program.StepRun, 0);
        }

        private void barButtonItem7_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (Program.FlowState != "Operacional")
                return;
            // Left
            DueFlowApiEsp32.SendJogCommand(-Program.StepRun, 0);
        }

        private void barButtonItem5_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (Program.FlowState != "Operacional")
                return;
            // Down
            if (Program.EixoRotativo)
                DueFlowApiEsp32.SendJogCommandZaxis(0, -Program.StepRun);
            else
                DueFlowApiEsp32.SendJogCommand(0, -Program.StepRun);
        }

        private void barButtonItem10_ItemClick(object sender, ItemClickEventArgs e)
        {
            // Volta completa - eixo rotativo
            if (Globals.DueStudio.DueConfiguration.Values.RotaryAxisDiameter == "")
            {
                XtraMessageBox.Show("Defina o valor do diâmetro do eixo antes de movimentá-lo.", "Erro");
                return;
            }
            float volta = (float)(Math.PI * float.Parse(Globals.DueStudio.DueConfiguration.Values.RotaryAxisDiameter));
            DueFlowApiEsp32.SendFullTurnZaxis(volta);
        }

        private void barButtonItem11_ItemClick(object sender, ItemClickEventArgs e)
        {
        }

        private void barButtonItem12_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (Program.FlowState != "Operacional")
                return;
            // Retorna ao zero
            DueFlowApiEsp32.RetornaInicio();
        }

        private void barButtonItem14_ItemClick(object sender, ItemClickEventArgs e)
        {
            
        }

        private void barButtonItem17_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (Program.FlowState != "Operacional")
                return;
            DueFlowApiEsp32.Homing();
        }

        private void btnUploadFileName_Click(object sender, EventArgs e)
        {
            frmFileName fFileName = new frmFileName();
            fFileName.Show();
        }

        private void btnWifi_Click(object sender, EventArgs e)
        {
            RefToFrmConnections.Show();
            RefToFrmConnections.BringToFront();
        }

        private void btnPlay_Click(object sender, EventArgs e)
        {
            if (cmbFlowFiles.SelectedItem == null)
                return;

            try
            {
                string FlowFileName = "";
                if (DueFlowApiEsp32.Connected || FlowUsbConnection.UsbFound)
                {
                    if (FlowUsbConnection.UsbFound)
                    {
                        FlowFileName = cmbFlowFiles.SelectedItem.ToString();
                    }
                    else
                    {
                        FlowFileName = (cmbFlowFiles.SelectedItem as Dtos.GcodeFilesEsp32.FlowFileEsp32).name;
                    }
                    DueFlowApiEsp32.RunSDFile(FlowFileName);
                }
            }
            catch (Exception Ex) { Debug.WriteLine(Ex.Message); }
        }

        private void btnCancelWork_Click(object sender, EventArgs e)
        {
            // Cancelar
            DueFlowApiEsp32.CancelPrint();
        }

        private void barButtonItem6_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (Program.FlowState != "Operacional")
                return;
            // Disparo Rápido / Mira NXT
            DueFlowApiEsp32.DisparoRapidoOuMira();
        }

        private void barButtonItem8_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (Program.FlowState != "Operacional")
                return;
            // Seta início
            DueFlowApiEsp32.SetaInicio();
        }

        private void FillFilesComboBox(Dtos.GcodeFilesEsp32.FlowFilesEsp32 flowFiles)
        {
            Task.Factory.StartNew(() =>
            {
                try
                {
                    var m1 = new MethodInvoker(() =>
                    {
                        var buffer = cmbFlowFiles.SelectedItem;
                        cmbFlowFiles.Properties.Items.Clear();
                        if (flowFiles != null)
                        {
                            if (flowFiles.files != null)
                            {
                                try
                                {
                                    for (int n = 0; n < flowFiles.files.Count; n++)
                                    {
                                        if (flowFiles.files[n].size != "-1")
                                        {
                                            cmbFlowFiles.Properties.Items.Add(flowFiles.files[n]);
                                        }
                                    }
                                    cmbFlowFiles.SelectedItem = buffer;
                                    cmbFlowFiles.ShowPopup();
                                }
                                catch { }
                            }
                        }
                        var pIdx = this.Controls.IndexOf(_p);
                        this.Controls[pIdx].SendToBack();
                        _p.Visible = false;
                        var pIdxCmb = this.Controls.IndexOf(cmbFlowFiles);
                        this.Controls[pIdxCmb].BringToFront();
                        cmbFlowFiles.Visible = true;
                        this.ForceRefresh();

                    });
                    cmbFlowFiles.Invoke(m1);
                }
                catch { }
            });
        }

        private void btnPortal_Click(object sender, EventArgs e)
        {
            //frmPortalLogin fPortalLogin = new frmPortalLogin();
            //fPortalLogin.Show();

            if (Program.fPortalDueIt == null)
            {
                Program.fPortalDueIt = new frmPortalDueItWeb();
                Program.fPortalDueIt.Show();
            }
            else
            {
                Program.fPortalDueIt.Show();
            }
        }

        private void barButtonItem4_ItemClick(object sender, ItemClickEventArgs e)
        {
            // TODO: remover
            // Frame
            
        }

        private void bReset_Click(object sender, EventArgs e)
        {
            if (DueFlowApiEsp32.Connected || FlowUsbConnection.UsbFound)
            {
                Program.FlowState = "Fechado";
                DueFlowApiEsp32.Reset();
            }
        }

        private void barButtonItem9_ItemClick(object sender, ItemClickEventArgs e)
        {
            Process[] updaterProcesses = Process.GetProcessesByName("Studio Update");
            if (updaterProcesses.Length == 0)
            {
                string updaterPath = @".\Studio Update.exe";
                Process.Start(updaterPath, "show");
            }
        }

        private void itemSair_ItemClick(object sender, ItemClickEventArgs e)
        {
            foreach (SystemWindow m in all_windows)
            {
                if (m.Title.Contains("- Due Studio 4"))
                {
                    m.SendClose();
                }
            }
            Process.GetCurrentProcess().Kill();
        }

        private void menAdmin_Popup(object sender, EventArgs e)
        {
            // TODO: remover
        }

        private void frmToolbar_Load(object sender, EventArgs e)
        {
            FlowUsbConnection.Core = new LaserGRBL.GrblCore(this);

            FlowUsbConnection.Core.MachineStatusChanged += Core_MachineStatusChanged;
            FlowUsbConnection.Core.OnFileLoaded += Core_OnFileLoaded;
            FlowUsbConnection.Core.OnOverrideChange += Core_OnOverrideChange;
            FlowUsbConnection.Core.IssueDetected += Core_IssueDetected;

            FlowUsbConnection.Core.OnReceivedLine += Core_OnReceivedLine;

            RefToFrmConnections.StartSearchThreads(); // removido temporariamente
        }

        public bool InitializeCore()
        {
            try
            {
                FlowUsbConnection.Core = new LaserGRBL.GrblCore(this);

                FlowUsbConnection.Core.MachineStatusChanged += Core_MachineStatusChanged;
                FlowUsbConnection.Core.OnFileLoaded += Core_OnFileLoaded;
                FlowUsbConnection.Core.OnOverrideChange += Core_OnOverrideChange;
                FlowUsbConnection.Core.IssueDetected += Core_IssueDetected;

                FlowUsbConnection.Core.OnReceivedLine += Core_OnReceivedLine;
            }
            catch (Exception Ex)
            {
                Debug.WriteLine(Ex.Message);
                return false;
            }

            return true;
        }

        private void Core_OnReceivedLine(string Line)
        {
            //Parse all dueflow message here by Ederson
            if (!_MultiLines)
            {
                LastReceivedMessageBuffer = Line;
                Utils.SerialMsgInterpreter.InterpretMsg(Line);
                _MultiLines = (Line.Contains("AP_LIST"));

                if (_MultiLines)
                {
                    if (Line.Contains("AP_LIST"))
                        _MultiLineEnd = "]}";
                }
            }
            else
            {
                if (Line.Contains("]}<"))
                {
                    // Fix do bug de mensagem de status ao final da mensagem da lista de wifis
                    try
                    {
                        Line = Line.Split('<')[0];
                    }
                    catch (Exception Ex)
                    {
                        Debug.WriteLine(Ex.Message);
                    }
                }
                if (Line.Contains(_MultiLineEnd))
                {
                    LastReceivedMessageBuffer += Line;
                    _MultiLines = false;
                    Utils.SerialMsgInterpreter.InterpretMsg(LastReceivedMessageBuffer);
                    LastReceivedMessageBuffer = "";
                }
                else
                {
                    LastReceivedMessageBuffer += Line;
                }
            }
        }

        private void Core_IssueDetected(LaserGRBL.GrblCore.DetectedIssue issue)
        {
            Debug.WriteLine("Core_IssueDetected");
        }

        private void Core_OnOverrideChange()
        {
            Debug.WriteLine("Core_OnOverrideChange");
        }

        private void Core_OnFileLoaded(long elapsed, string filename)
        {
            Debug.WriteLine("Core_OnFileLoaded");
        }

        private void Core_MachineStatusChanged()
        {
            /*if (DueFlowApiEsp32.Connected)
                return;
            if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Idle)
            {
                Program.FlowState = "Operacional";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Alarm)
            {
                Program.FlowState = "Alarme";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Hold)
            {
                Program.FlowState = "Parado";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Check)
            {
                Program.FlowState = "Verificação";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Door)
            {
                Program.FlowState = "Tampa de segurança aberta";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Connecting)
            {
                Program.FlowState = "Conectando";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Run)
            {
                Program.FlowState = "Trabalhando";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Cooling)
            {
                Program.FlowState = "Pausado"; // "Cooling"
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Sen)
            {
                Program.FlowState = "Sensores";
            }
            else
            {
                Program.FlowState = "Desconhecido";
            }*/
        }

        private void barSubItem2_ItemClick(object sender, ItemClickEventArgs e)
        {
            // Desabilita botão de volta completa se eixo não estiver ativado
            if (Program.EixoRotativo)
                barButtonItem10.Enabled = true;
            else
                barButtonItem10.Enabled = false;
        }

        private void cmbFlowFiles_MouseDown_1(object sender, MouseEventArgs e)
        {
            if (cmbFlowFiles.IsPopupOpen)
            {
                cmbFlowFiles.ClosePopup();
            }
            else
            {
                var pIdx = this.Controls.IndexOf(_p);
                this.Controls[pIdx].BringToFront();
                _p.Visible = true;
                this.ForceRefresh();

                _buffer = cmbFlowFiles.SelectedItem;
                cmbFlowFiles.Properties.Items.Clear();

                Task.Factory.StartNew(() =>
                {
                    if (FlowUsbConnection.UsbFound)
                    {
                        //Globals.DueStudio.UsbGcodeFiles = new Dtos.GcodeFilesEsp32.FlowFilesEsp32();
                        Globals.DueStudio.UsbGcodeFiles.files = new List<Dtos.GcodeFilesEsp32.FlowFileEsp32>();
                        FlowUsbConnection.SendCommand("[ESP210]");
                        // FillFilesComboBox chamado pelo evento quando receber os dados
                    }
                    else
                    {
                        FillFilesComboBox(DueFlowApiEsp32.ListSDFiles());
                    }
                    var now = DateTime.Now;
                    while (DateTime.Now < now.AddMilliseconds(5000))
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    if (cmbFlowFiles.Properties.Items.Count == 0)
                    {
                        var m1 = new MethodInvoker(() =>
                        {
                            var pIdx2 = this.Controls.IndexOf(_p);
                            this.Controls[pIdx2].SendToBack();
                            _p.Visible = false;
                            var pIdxCmb = this.Controls.IndexOf(cmbFlowFiles);
                            this.Controls[pIdxCmb].BringToFront();
                            cmbFlowFiles.Visible = true;
                            this.ForceRefresh();
                        });
                        this.Invoke(m1);
                    }
                });
            }
        }

        private void btnPausePlay_Click(object sender, EventArgs e)
        {
            if (btnPausePlay.Tag.ToString() == "pause")
            {
                Program.UserPaused = true;
                DueFlowApiEsp32.PausePrint();
            }
            else if (btnPausePlay.Tag.ToString() == "play")
            {
                Program.UserPaused = false;
                DueFlowApiEsp32.Resume();
            }
        }

        private void bCamera_Click(object sender, EventArgs e)
        {
            frmSmartCam fSmartCam = new frmSmartCam();
            fSmartCam.Show();
        }

        private void btnCameraGetInk_Click(object sender, EventArgs e)
        {
            frmCameraScan fCameraScan = new frmCameraScan();
            fCameraScan.Show();

            Task.Factory.StartNew(() =>
            {    
                Image<Bgr, byte> image = DueCamV2.GetImageDone();
                if (image == null)
                    MessageBox.Show("Erro ao obter imagem, Verifique se a câmera está conectada e se o dispositivo está ligado.");

                var mSmartCamClose = new MethodInvoker(() =>
                {
                    fCameraScan.Close();
                });
                this.Invoke(mSmartCamClose);
            });
        }

        private void btnAlert_Click(object sender, EventArgs e)
        {
            if (_tampaError)
                ShowAvisoTampaErrorMsg();
            else if (_sensoresError)
                ShowAvisoSensoresErrorMsg();
        }

        private void timUpdateMachineState_Tick(object sender, EventArgs e)
        {
            if (DueFlowApiEsp32.Connected)
                return;
            if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Idle)
            {
                Program.FlowState = "Operacional";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Alarm)
            {
                Program.FlowState = "Alarme";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Hold)
            {
                Program.FlowState = "Parado";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Check)
            {
                Program.FlowState = "Verificação";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Door)
            {
                //Program.FlowState = "Porta";
                if (FlowUsbConnection.Core.DoorState == LaserGRBL.GrblCore.DoorStatus.Closed)
                    Program.FlowState = "Tampa de segurança fechada";
                else if (FlowUsbConnection.Core.DoorState == LaserGRBL.GrblCore.DoorStatus.Open)
                    Program.FlowState = "Tampa de segurança aberta";
                else
                    Program.FlowState = "Porta";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Connecting)
            {
                Program.FlowState = "Conectando";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Run)
            {
                Program.FlowState = "Trabalhando";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Jog)
            {
                Program.FlowState = "Movimentando";
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Cooling)
            {
                Program.FlowState = "Pausado"; // "Cooling"
            }
            else if (FlowUsbConnection.Core.MachineStatus == LaserGRBL.GrblCore.MacStatus.Sen)
            {
                if (FlowUsbConnection.Core.SensorErr == LaserGRBL.GrblCore.SensorError.Temperature)
                    Program.FlowState = "Sensores - temperatura muito alta";
                else if (FlowUsbConnection.Core.SensorErr == LaserGRBL.GrblCore.SensorError.WaterLevel)
                    Program.FlowState = "Sensores - sem água no reservatório";
                else if (FlowUsbConnection.Core.SensorErr == LaserGRBL.GrblCore.SensorError.WaterFlowRate)
                    Program.FlowState = "Sensores - fluxo de água insuficiente";
                else
                    Program.FlowState = "Sensores";
            }
            else
            {
                Program.FlowState = "Fechado";
            }
        }

        private void timCheckInputs_Tick(object sender, EventArgs e)
        {
            if (FlowUsbConnection.UsbFound)
            {
                ShowRotaryAxisMsg(FlowUsbConnection.Core.RotaryAxisState == 1);
                //ShowDriverXYErrorMsg(FlowUsbConnection.Core.DriverXErrorState == 1, FlowUsbConnection.Core.DriverYErrorState == 1);
            }
            else if (DueFlowApiEsp32.Connected)
            {
                ShowRotaryAxisMsg(Globals.DueStudio.RotaryAxisState == 1);
                //ShowDriverXYErrorMsg(Globals.DueStudio.DriverXErrorState == 1, Globals.DueStudio.DriverYErrorState == 1);
            }
        }

        private void ShowRotaryAxisMsg(bool rotaryConnected)
        {
            if (rotaryConnected)
            {
                if (!_rotaryAxisConnectedShown)
                {
                    _rotaryAxisConnectedShown = true;
                    Utils.General.CloseFormIfOpen("frmRotaryAxisConnected");
                    Utils.General.CloseFormIfOpen("frmRotaryAxisDisconnected");
                    _fRotaryAxisConnected.Show();
                }
            }
            else
            {
                if (_rotaryAxisConnectedShown)
                {
                    _rotaryAxisConnectedShown = false;
                    Utils.General.CloseFormIfOpen("frmRotaryAxisConnected");
                    Utils.General.CloseFormIfOpen("frmRotaryAxisDisconnected");
                    _fRotaryAxisDisconnected.Show();
                }
            }
        }

        private string CheckForDriverXYError()
        {
            if (FlowUsbConnection.UsbFound)
            {
                return DriverXYErrorMsg(FlowUsbConnection.Core.DriverXErrorState == 1, FlowUsbConnection.Core.DriverYErrorState == 1);
            }
            else if (DueFlowApiEsp32.Connected)
            {
                return DriverXYErrorMsg(Globals.DueStudio.DriverXErrorState == 1, Globals.DueStudio.DriverYErrorState == 1);
            }
            return "";
        }

        private string DriverXYErrorMsg(bool driverXError, bool driverYError)
        {
            if (driverXError || driverYError)
            {
                if (driverXError && driverYError)
                    return "Erro nos drivers XY, reinicie";
                else
                {
                    string driver = driverXError ? "X" : "Y";
                    return string.Format("Erro no driver {0}, reinicie", driver);
                }
            }
            return "";
        }

        private void btnImportImage_Click(object sender, EventArgs e)
        {
            var p = new Process();
            p.StartInfo.FileName = Directory.GetCurrentDirectory()+"\\dit\\dithering.exe";
            p.Start();
        }

        private void SetBtnAlertVisibility()
        {
            if (_sensoresError || _tampaError)
            {
                btnAlert.Visible = true;
            }
            else
            {
                btnAlert.Visible = false;
            }
        }

        private void ShowAvisoTampaErrorMsg()
        {
            avisoTampa.Show();
            General.FocusOnFormIfOpen("avisoTampa");
        }

        private void ShowAvisoSensoresErrorMsg()
        {
            avisoSensores.ShowWithMsg(Program.FlowState);
            General.FocusOnFormIfOpen("avisoSensores");
        }
    }
}
