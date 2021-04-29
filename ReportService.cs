using Cipher;
using EmailSender;
using ReportService.Core;
using ReportService.Core.Repository;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace ReportService
{
    public partial class ReportService : ServiceBase
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        //private const int sendHour = 8;  //Godzina wysylania raportu
        private int sendHour = Convert.ToInt32(ConfigurationManager.AppSettings["SendHours"]);

        //private const int IntervalInMinutes = 1; //Tutaj liczba minut
        private static int IntervalInMinutes = Convert.ToInt32(ConfigurationManager.AppSettings["IntervalInMinutes"]);
        private Timer _timer = new Timer(IntervalInMinutes * 60000);
        private ErrorRepository _errorRepository = new ErrorRepository();
        private ReportRepository _reportRepository = new ReportRepository();
        private Email _email;
        private GenerateHtmlEmail _htmlEmail = new GenerateHtmlEmail();
        private string _emailReceiver;
        private StringCipher _stringCipher = new StringCipher("845B5418-31C9-4B76-9A8D-1456E704B960"); 
        private const string NotEncryptedPasswordPrefix = "encrypt:";
        private bool _enableSendReport = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSendReport"]);

 

        //Konstruktor w ktorym takze inicjalizujemy objekt klasy Email
        public ReportService()
        {
            InitializeComponent();

            try
            {
                _emailReceiver = ConfigurationManager.AppSettings["ReceiverEmail"];

                var encryptedPassword = ConfigurationManager.AppSettings["SenderEmailPassword"];

                if (encryptedPassword.StartsWith("encrypt:"))
                {
                    encryptedPassword = _stringCipher.Encrypt(encryptedPassword.Replace("encrypt:",""));

                    //Podmieniamy wartsoc w naszym App.config
                    var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    configFile.AppSettings.Settings["SenderEmailPassword"].Value = encryptedPassword;
                    configFile.Save();
                }
                _email = new Email(new EmailParams
                {
                    HostSmtp = ConfigurationManager.AppSettings["HostSmtp"],
                    Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]),
                    EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSsl"]),
                    SenderName = ConfigurationManager.AppSettings["SenderName "],
                    SenderEmail = ConfigurationManager.AppSettings["SenderEmail"],
                    //SenderEmailPassword = _stringCipher.Decrypt(ConfigurationManager.AppSettings["SenderEmailPassword"])
                    SenderEmailPassword = DecryptSenderEmailPassword()
                }) ;
            }
            catch(Exception ex)
            {
                Logger.Error(ex, ex.Message); //Zapisywanie bledu do pliku jesli wystapi blad w Try.
                throw new Exception(ex.Message); //A nastepnie rzucenie tego bledu ponownie
            }

        }

        private string DecryptSenderEmailPassword()
        {
            var encryptedPassword = ConfigurationManager.AppSettings["SenderEmailPassword"];

            if (encryptedPassword.StartsWith(NotEncryptedPasswordPrefix))
            {
                encryptedPassword = _stringCipher.Encrypt(encryptedPassword.Replace(NotEncryptedPasswordPrefix, ""));

                //Podmieniamy wartsoc w naszym App.config
                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                configFile.AppSettings.Settings["SenderEmailPassword"].Value = encryptedPassword;
                configFile.Save();
            }
            return _stringCipher.Decrypt(encryptedPassword);
        }

        protected override void OnStart(string[] args)
        {
            _timer.Elapsed += DoWork;
            _timer.Start();
            Logger.Info("Service started...");
        }

        //W tym przypadku bedzie void zamiast Task poniewaz tak to nie mozna by uzyc  _timer.Elapsed += DoWork w OnStart()
        private async void DoWork(object sender, ElapsedEventArgs e)
        {
            //Tutaj bedzie cala logika dotyczaca wysylania raportu. Tutaj dodajemy takze informacje o wszystkich bledach w naszym serwisie
            try 
            {
               await SendError();

                if (_enableSendReport == true)
                    await SendReport();
            }
            catch (Exception ex)
            {
                Logger.Error(ex, ex.Message); //Zapisywanie bledu do pliku jesli wystapi blad w Try.
                throw new Exception(ex.Message); //A nastepnie rzucenie tego bledu ponownie
            }
        }

        private async Task SendError()
        {
            //Tutaj beda fejkowe dane wpisane na sztywno ale mozna w przyszlosci to napisac zeby wysylal bledy z bazy danych jesli bedzie taka potrzeba.
            var errors = _errorRepository.GetLastErrors(IntervalInMinutes);
            if (errors == null || !errors.Any())
            {
                return; //Jesli nic nie ma zadnych bledow to wychodzimy z tej metody
            }

            //Send email. 
            await _email.Send("Bledy w aplikacji", _htmlEmail.GenerateErrors(errors, IntervalInMinutes), _emailReceiver);

            Logger.Info("Error sent");
        }

        
        private async Task SendReport()
        {
            var actualHour = DateTime.Now.Hour;

            if (actualHour < sendHour)
            {
                return; 
            }
            //Ostatni nie wyslany raport
            var report = _reportRepository.GetLastNotSentReport();

            if (report == null)
            {
                return;
            }
            //Jesli raport istnieje to bedzie wyslany email.
            await _email.Send("Raport dobowy", _htmlEmail.GenerateReport(report), _emailReceiver);

            _reportRepository.ReportSent(report);

            Logger.Info("Report sent");
        }

        protected override void OnStop()
        {
            Logger.Info("Service stopped...");
        }
    }
}
