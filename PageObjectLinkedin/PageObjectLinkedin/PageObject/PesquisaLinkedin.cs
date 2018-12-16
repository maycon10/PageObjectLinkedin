using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support;
using ClosedXML.Excel;
using System.Collections;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;

namespace PageObjectLinkedin
{



    [TestClass]
    public class PesquisaLinkedin
    {
       

        public By loginComboBox { get { return By.ClassName("login-email"); } }
        public By loginComboBoxSengundaOpcao { get { return By.Id("username"); } }
        public By passWordComboBox { get { return By.ClassName("login-password"); } }
        public By passWordComboBoxSegundaOpcao { get { return By.Id("password"); } }
        public By clickSignIN { get { return By.Id("login-submit"); } }
        public By clickSignINsegundaOpcao { get { return By.XPath("//button[@aria-label='Sign in']"); } }
        public By pesquisaPessoa { get { return By.CssSelector("input[role = 'combobox']"); } }
        public By clickConexoes { get { return By.ClassName("search-s-facet__name"); } }
        public By clickCheckBox { get { return By.CssSelector("li:nth-child(2) > label > p > span"); } }
        public By clickAplicar { get { return By.XPath("//span[text()='Aplicar']"); } }
        public By listaElementosConectar { get { return By.CssSelector("button[data-control-name='srp_profile_actions']"); } }
        public By clickConectar { get { return By.CssSelector("button[data-control-name='srp_profile_actions']"); } }
        public By mapearModal { get { return By.XPath("//button[text()='Enviar agora']"); } }
        public By clickEnviarAgora { get { return By.XPath("//button[text()='Enviar agora']"); } }

    }


}




 




