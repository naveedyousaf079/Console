﻿using OTP.ViewModels;
using System.ComponentModel;
using Xamarin.Forms;

namespace OTP.Views
{
    public partial class ItemDetailPage : ContentPage
    {
        public ItemDetailPage()
        {
            InitializeComponent();
            BindingContext = new ItemDetailViewModel();
        }
    }
}