using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;
using System.Text;
using System.Threading;

namespace RpaChallenge
{
  /// <summary>
  /// TPA que extrai dados de uma spreadsheet e preenche formulário com tais dados
  /// </summary>
  class Program
  {
    public static ChromeDriver _driver;
    static void Main(string[] args)
    {
      Start();
    }
    /// <summary>
    /// Função que liga o robô e chama cada processo
    /// </summary>
    private static void Start()
    {
      try
      { // Continuar a partir de amanhã
        Logger("--- O Robô Iniciou ------------------------------------------------------");
        ExecuteAndLog(OpenBrowser(), "OpenBrowser()");
        ExecuteAndLog(NavigateToChallengePage(), "NavigateToChallengePage()");
        ExecuteAndLog(DownloadSpreadSheet(), "DownloadSpreadSheet()");
        // ExecuteAndLog(MoveAndReadSheet(), "MoveAndReadSheet()");
        ExecuteAndLog(FillFormAndSubmit(), "FillFormAndSubmit()");
        Logger("--- O Robô Finalizou ------------------------------------------------------");
      }
      catch (Exception e)
      {
        // Retornar erro em alguma lugar
      }
    }

    #region Utility Tools
    /// <summary>
    /// Registra feedbacks sobre o andamento do robô
    /// </summary>
    /// <param name="feedback">Mensagem de feedback</param>
    private static void Logger(string feedback)
    {
      string path = @"C:\Users\luis2\source\repos\RpaChallenge\Log.txt";
      if (!File.Exists(path))
      {
        // Cria um arquivo e escreve nele
        using (FileStream fileStream = new FileStream(path, FileMode.Create))
        {
          String message = String.Concat(DateTime.Now, ": ", feedback);
          byte[] byteMessage = new UTF8Encoding(true).GetBytes(message);
          fileStream.Write(byteMessage, 0, byteMessage.Length);
        }
      }
      else
      {
        // Acrescenta registro no final do arquivo já existente
        using (StreamWriter streamWriter = File.AppendText(path))
        {
          streamWriter.WriteLine(String.Concat(DateTime.Now, ": ", feedback));
        }
      }
    }
    /// <summary>
    /// Método que verifica o resultado de uma ação foi positivo, se sim registra isto no Log.txt
    /// </summary>
    /// <param name="actionResult">Resultado da ação</param>
    /// <param name="actionName">Nome da ação</param>
    private static void ExecuteAndLog(bool actionResult, string actionName)
    {
      if (actionResult)
      {
        Logger($"A ação {actionName} foi executada com sucesso");
      }
    }
    /// <summary>
    /// Mantém-se tentando interagir com um elemento HTML até que um certo tempo limite seja atingido
    /// </summary>
    /// <param name="actionName">Identifica a ação</param>
    /// <param name="xpath">Identifica o elemento</param>
    /// <param name="timeOut">Tempo máximo de tentativa</param>
    /// <param name="inputKeys">Dados de entrada (Opcional)</param>
    /// <returns></returns>
    private static bool KeepTryingUntil(string actionName, string xpath, int timeOut, string inputKeys = "")
    {
      if (actionName == "Click")
      {
        int duration = -1;
        do
        {
          Thread.Sleep(100);
          _driver.FindElement(By.XPath(xpath)).Click();
          duration++;
        } while (_driver.FindElements(By.XPath(xpath)).Count == 0 && duration < timeOut);
        return true;
      }
      else if (actionName == "InputKeys")
      {
        int duration = -1;
        do
        {
          Thread.Sleep(100);
          _driver.FindElement(By.XPath(xpath)).SendKeys(inputKeys);
          duration++;
        } while (_driver.FindElements(By.XPath(xpath)).Count == 0 && duration < timeOut);
        return true;
      }
      return false;
    }
    #endregion
    /// <summary>
    /// Processo simples que apenas tenta abrir o navegador
    /// </summary>
    private static bool OpenBrowser()
    {
      try
      {
        _driver = new ChromeDriver();
        return true;
      }
      catch (Exception e)
      {
        Logger(e.Message);
        return false;
      }
    }
    /// <summary>
    /// Navega para o endereço deste desafio
    /// </summary>
    private static bool NavigateToChallengePage()
    {
      try
      {
        _driver.Navigate().GoToUrl("http://www.rpachallenge.com/");
        return true;
      }
      catch (Exception e)
      {
        Logger(e.Message);
        return false;
      }
    }
    /// <summary>
    /// Fica tentando realizar o download
    /// </summary>
    /// <returns></returns>
    private static bool DownloadSpreadSheet()
    {
      return KeepTryingUntil("Click", "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a", 3000);
    }
    /// <summary>
    /// Gerenciar o arquivo baixado (de Excel)
    /// </summary>
    /// <returns></returns>
    private static bool MoveAndReadSheet()
    {
      try
      {
        // Move o arquivo, se necessário
        if (!File.Exists(@"C:\Users\luis2\source\repos\RpaChallenge\challenge.xlsx"))
        {
          File.Move(@"C:\Users\luis2\Downloads\challenge.xlsx", @"C:\Users\luis2\source\repos\RpaChallenge\challenge.xlsx");
        }
        else
        {
          Logger(@"O arquivo challenge.xlsx já existia na pasta de destino C:\Users\luis2\source\repos\RpaChallenge\ portanto o que foi baixado não foi movido");
        }
        // Ler arquivo do excel   --------------------------------------------------aqui
        // Extrair informações do Excel File
        // - ------------------------------------------------------------------------------
        return true;
      }
      catch (Exception e)
      {
        Logger(e.Message);
        return false;
      }
    }
    /// <summary>
    /// Preenche o formulário e o submete
    /// </summary>
    /// <returns></returns>
    private static bool FillFormAndSubmit()
    {
      var CorrectPath = _driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div/div[1]/rpa1-field/div/label")).Text;
      bool address = KeepTryingUntil("InputKeys", "/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div/div[1]/rpa1-field/div/input", 3000, "Joao Pinheiro 2020");
      bool lastName = KeepTryingUntil("InputKeys", "/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div/div[2]/rpa1-field/div/input", 3000, "da Silva Pinheiro");
      bool email = KeepTryingUntil("InputKeys", "/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div/div[3]/rpa1-field/div/input", 3000, "luis.2.pinheiro@gmail.com");
      bool role = KeepTryingUntil("InputKeys", "/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div/div[4]/rpa1-field/div/input", 3000, "Centers Junior");
      bool firstName = KeepTryingUntil("InputKeys", "/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div/div[5]/rpa1-field/div/input", 3000, "Luís Henrique");
      bool phoneNumber = KeepTryingUntil("InputKeys", "/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div/div[6]/rpa1-field/div/input", 3000, "34920009266");
      bool companyName = KeepTryingUntil("InputKeys", "/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div/div[7]/rpa1-field/div/input", 3000, "everis");
      return address && lastName && email && role && firstName && phoneNumber && companyName;
    }
  }
}
