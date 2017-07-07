using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SHDocVw;
using System.Runtime.InteropServices;
using mshtml;
using System.Collections;

namespace Minha_demanda_Jira
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("D40C654D-7C51-4EB3-95B2-1E23905C2A2D")]
    
    public partial class Form1 : Form
    {
        static InternetExplorer ie;
        private IWebBrowser2 _browser;
        private const string UrlLogin = "https://restrito.cotemig.com.br/login/";
        string Metodo = "Login";

        public Form1()
        {
            InitializeComponent();
        }

        private void btn_entrar_Click(object sender, EventArgs e)
        {
            CarregaDados();
        }

        private void CarregaDados()
        {
            Metodo = "Login";

            ie = new InternetExplorer() { Visible = false, };

            ie.Navigate2(UrlLogin);

            ie.DocumentComplete += Ie_DocumentComplete;
        }

        private void Ie_DocumentComplete(object pDisp, ref object URL)
        {
            switch (Metodo)
            {
                case "Login":
                    login();
                    break;
                case "ValidarTelaInicial":
                    ValidarTelaInicial();
                    break;
                case "PegarValor":
                    PegarValor();
                    break;
                default:
                    Console.WriteLine("Default case");
                    break;
            }
        }

        public void login()
        {
            Metodo = "ValidarTelaInicial";

            var document = ie.Document as IHTMLDocument3;

            var User = document.getElementById("user_login");
            User.innerText = txt_usuario.Text;

            var password = document.getElementById("user_password");
            password.innerText = txt_senha.Text;

            var SignIn = document.getElementsByTagName("button");
            var login = (IHTMLElement)SignIn.item(0);
            login.click();
        }

        public void ValidarTelaInicial()
        {
            Metodo = "PegarValor";
            ie.Navigate2("https://restrito.cotemig.com.br/diario/boletim.php");
        }

        public void PegarValor()
        {
            Metodo = "PegarValor";

            var document = ie.Document as IHTMLDocument3;

            var Dashboard = document.getElementById("ui-id-2").innerHTML;

            webBrowser1.DocumentText = Dashboard;
        }
    }
}
