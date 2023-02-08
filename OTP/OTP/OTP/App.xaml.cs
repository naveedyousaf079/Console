using OTP.Services;
using OTP.Views;
using System;
using Xamarin.Forms;
using Xamarin.Forms.Xaml;
using OtpNet;
using System.Text;

namespace OTP
{
    public partial class App : Application
    {

        public App()
        {
            InitializeComponent();
            //GenerateOTP("JBSWY3DPEHPK3PXP");
            DependencyService.Register<MockDataStore>();
            //MainPage = new AppShell();
            MainPage = new GenerateOTP();
        }

        private void GenerateOTPFromKey(string secretKey)
        {
            //if (!string.IsNullOrWhiteSpace(secretKey))
            //{
            //    var totp = new Totp(Encoding.ASCII.GetBytes(secretKey),30,totpSize:6);
            //    var totpCode = totp.ComputeTotp(DateTime.UtcNow);
            //    // or use the overload that uses UtcNow
            //    var totpCodeValue = totp.ComputeTotp();
            //    var remainingTime = totp.RemainingSeconds();
            //    // there is also an overload that lets you specify the time
            //    var remainingSeconds = totp.RemainingSeconds(DateTime.UtcNow);
            //    Console.ReadLine();
            //}
            //else
            //{
            //    //Alert
            //}
        }

        protected override void OnStart()
        {
        }

        protected override void OnSleep()
        {
        }

        protected override void OnResume()
        {
        }
    }
}
