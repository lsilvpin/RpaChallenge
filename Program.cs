using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using Syncfusion.XlsIO;
using Newtonsoft.Json;

namespace RpaChallenge
{
  /// <summary>
  /// RPA que extrai dados de uma spreadsheet e preenche formulário com tais dados
  /// </summary>
  class Program
  {
    public static ChromeDriver _driver; // Objeto de gerenciamento da página da web
    public static List<Employee> _employees = new List<Employee>(); // Receberá os funcionários do Excell
    /// <summary>
    /// Ponto de entrada do programa
    /// </summary>
    /// <param name="args">Podem ser passadas entradas ao programa quando chamá-lo para execussão (desnecessário)</param>
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
      {
        Logger("--- O Robô Iniciou ------------------------------------------------------");
        ExecuteAndLog(OpenBrowser(), "OpenBrowser()"); // Abre o navegador
        ExecuteAndLog(NavigateToChallengePage(), "NavigateToChallengePage()"); // Navega até a página do desafio
        ExecuteAndLog(ClickOnStart(), "ClickOnStart()"); // Clica no botão Start
        ExecuteAndLog(DownloadSpreadSheet(), "DownloadSpreadSheet()"); // Baixa o arquivo do excel
        ExecuteAndLog(MoveSheetIfNecessary(), "MoveSheetIfNecessary()"); // Move o excel se ainda não existir no destino
        ExecuteAndLog(XlToJson(), "XlToJson()"); // Converte Excel para Json
        ExecuteAndLog(JsonToEntityAndCollect(), "JsonToEntityAndCollect()"); // Converte Json para Entidade e Coleta lista
        foreach (Employee employee in _employees)
        { // Este loop deve ser repetido, uma vez para cada empregado presente na spreadsheet
          ExecuteAndLog(FillForm(employee.roleInCompany, employee.lastName, employee.email,
          employee.firstName, employee.companyName, employee.phoneNumber, employee.address), "FillForm()"); // Preenche formulário
          ExecuteAndLog(SubmitForm(), "SubmitForm()"); // Submete o formulário
        }
        Logger("--- O Robô Finalizou ------------------------------------------------------");
      }
      catch (Exception e)
      {
        Logger(e.Message);
        Logger("--- O Robô Finalizou com ERRO ------------------------------------------------------");
      }
    }
    // Abaixo uma caixa de ferramentas úteis que o robô precisa de forma recorrente
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
    /// Método que verifica se o resultado de uma ação foi positivo, se sim registra isto no Log.txt
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
    // A seguir estão as ações (processos/passo a passo) que o robô executará em Start()
    #region Actions
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
    private static bool MoveSheetIfNecessary()
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
    /// Converte arquivo Excel para Json
    /// </summary>
    /// <returns></returns>
    private static bool XlToJson()
    {
      try
      {
        //Instantiate the spreadsheet creation engine.
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
          IApplication application = excelEngine.Excel;

          //The workbook is opened.
          FileStream fileStream = new FileStream(@"C:\Users\luis2\source\repos\RpaChallenge\challenge.xlsx", FileMode.Open);

          IWorkbook workbook = application.Workbooks.Open(fileStream, ExcelOpenType.Automatic);
          IWorksheet worksheet = workbook.Worksheets[0];

          //Export worksheet data into CLR Objects
          IList<Employee> employees = worksheet.ExportData<Employee>(1, 1, 11, 7);

          //open file stream
          using (StreamWriter file = File.CreateText(@"C:\Users\luis2\source\repos\RpaChallenge\challenge.json"))
          {
            JsonSerializer serializer = new JsonSerializer();

            //serialize object directly into file stream
            serializer.Serialize(file, employees);
          }
        }
        return true;
      }
      catch (Exception e)
      {
        Logger(e.Message);
        return false;
      }
    }
    /// <summary>
    /// Converte o arquivo Json para uma Entidade do CSharp e salva globalmente para uso do robô
    /// </summary>
    /// <returns></returns>
    private static bool JsonToEntityAndCollect()
    {
      try
      {
        string json = File.ReadAllText(@"C:\Users\luis2\source\repos\RpaChallenge\challenge.json"); // Abre o arquivo Json como string
        List<Employee> employees = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Employee>>(json); // Coleta a lista de empregados
        _employees = employees; // Salva a lista dos empregados globalmente
        return true;
      }
      catch (Exception e)
      {
        Logger(e.Message);
        return false;
      }
    }
    /// <summary>
    /// Preenche o formulário dinâmico
    /// </summary>
    /// <returns></returns>
    private static bool FillForm(string roleInCompany, string lastName, string email, string firstName,
      string companyName, string phoneNumber, string address)
    {
      // Primeiro temos que mapear os campos do formulário
      Dictionary<string, string> formMap = new Dictionary<string, string>();
      for (int i = 1; i <= 7; i++)
      {
        formMap.Add(_driver.FindElement(By.XPath($"/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div/div[{i}]/rpa1-field/div/label")).Text, $"/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/div/div[{i}]/rpa1-field/div/input");
      }
      // Agora é só usar o mapa para preencher os campos corretamente
      return KeepTryingUntil("InputKeys", formMap["Role in Company"], 3000, roleInCompany) &&
        KeepTryingUntil("InputKeys", formMap["Last Name"], 3000, lastName) &&
        KeepTryingUntil("InputKeys", formMap["Email"], 3000, email) &&
        KeepTryingUntil("InputKeys", formMap["First Name"], 3000, firstName) &&
        KeepTryingUntil("InputKeys", formMap["Company Name"], 3000, companyName) &&
        KeepTryingUntil("InputKeys", formMap["Phone Number"], 3000, phoneNumber) &&
        KeepTryingUntil("InputKeys", formMap["Address"], 3000, address);
    }
    /// <summary>
    /// Submete o formulário dinâmico (já preenchido)
    /// </summary>
    /// <returns></returns>
    private static bool SubmitForm()
    {
      return KeepTryingUntil("Click", "/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input", 3000);
    }
    /// <summary>
    /// Clica no botão "Start" da tela inicial do challenge
    /// </summary>
    /// <returns></returns>
    private static bool ClickOnStart()
    {
      return KeepTryingUntil("Click", "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button", 3000);
    }
    #endregion
  }
}
