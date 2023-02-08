using Android.Graphics.Fonts;
using OtpNet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Xamarin.Forms;
using Xamarin.Forms.PlatformConfiguration.TizenSpecific;
using Xamarin.Forms.Xaml;

namespace OTP
{
    [XamlCompilation(XamlCompilationOptions.Compile)]
    public partial class GenerateOTP : ContentPage
    {
        //Secret key for the OTP should be encrypted
        private const string SecretKeyForOTP= "JBSWY3DPEHPK3PXP";

        //Setting Number of Digits in Token
        private const int OTPDigits = 6;

        //Settingthe Token Validity Period in Seconds
        private const int OTPTokePeriod = 30;

        //Content Stack Layout will be setting in CTOR
        private StackLayout stackLayout = null;


        public GenerateOTP()
        {
            //Setting the Background Color of Page
            this.BackgroundColor = Color.LightBlue;

            //Creating a Label to display OTP
            var tokenDisplayLabel = new Xamarin.Forms.Label { TextColor = Color.Wheat, IsVisible = false, Text = "" };

            //Creating the button for generation of OTP
            var OtpButton = new Button { TextColor = Color.White,Text= "Generate OTP" };

            //Registring the click event for this button
            OtpButton.Clicked += Button_Clicked;

            //Creating a stack Layout
            stackLayout = new StackLayout
            {
                Padding = new  Thickness(50),
                Margin = new Thickness(0,100,0,0),
                Children =
                {
                    OtpButton,tokenDisplayLabel
                }
            };

            //Adding th to Content
            this.Content = stackLayout;
        }

        private void Button_Clicked(object sender, EventArgs e)
        {

            //Check if the OTP Scret key is available
            if (!string.IsNullOrWhiteSpace(SecretKeyForOTP))
            {
                //Creating an instance of OTP Class with required Params
                var totp = new Totp(Encoding.ASCII.GetBytes(SecretKeyForOTP), OTPTokePeriod, totpSize: OTPDigits);

                //Setting the date as UTC to avoid date and time conflict
                var totpCode = totp.ComputeTotp(DateTime.UtcNow);

                // Calling the OTP Compute method
                var totpCodeValue = totp.ComputeTotp();

                //Calling the remaining Second not required but still
                var remainingTime = totp.RemainingSeconds();

                // there is also an overload that lets you specify the time
                var remainingSeconds = totp.RemainingSeconds(DateTime.UtcNow);

                //Creating a new Label and setting it's Properties
                Xamarin.Forms.Label secondLabel = new Xamarin.Forms.Label
                {
                    Text = totpCodeValue,
                    HorizontalOptions = LayoutOptions.StartAndExpand,
                    BackgroundColor = Color.Blue,
                    WidthRequest = 300,
                    HeightRequest= 50,
                    HorizontalTextAlignment= TextAlignment.Center,
                    TextColor = Color.White,
                    Margin = new Thickness(10,10),
                    Padding = new Thickness(0,10,0,0),
                    FontSize = 12
                };

                //removing athe already created Label
                stackLayout.Children.RemoveAt(1);

                // Adding this newly created Label
                stackLayout.Children.Add(secondLabel);
            }
            else
            {
                // To be replaced by Logger
                Console.WriteLine("OTP Secret key Missing");
            }
        }
    }
}