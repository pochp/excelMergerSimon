/// Made by : Pedro Rodriguez
/// Date : 14 mars 2019
/// Github : https://github.com/pochp/excelMergerSimon


using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excelMerger
{
    class Program
    {
        const string SHEET_NAME = "Ventes par vendeur";
        const int FIRST_ROW = 7;
        const string CREATED_FILE_PATH = @"C:\Users\Pedro\Documents\devoir cegep simon\data\Liste des ventes par client Combine.xls";

        static void Main(string[] args)
        {
            if(File.Exists(CREATED_FILE_PATH))
            {
                MessageBox.Show("Error : File already exists with designated path and name");
                return;
            }

            List<List<ClientDataYear>> client_data = new List<List<ClientDataYear>>();

            client_data.Add(GetClientData(@"C:\Users\Pedro\Documents\devoir cegep simon\data\Liste des ventes par client 2016.xls", "2016"));
            client_data.Add(GetClientData(@"C:\Users\Pedro\Documents\devoir cegep simon\data\Liste des ventes par client 2017.xls", "2017"));
            client_data.Add(GetClientData(@"C:\Users\Pedro\Documents\devoir cegep simon\data\Liste des ventes par client 2018.xls", "2018"));

            CreateExportData(CombineData(client_data), CREATED_FILE_PATH);
        }

        static List<ClientDataYear> GetClientData(string filePath, string year)
        {
            //npoi code taken from https://stackoverflow.com/questions/5855813/how-to-read-file-using-npoi

            List<ClientDataYear> data = new List<ClientDataYear>();

            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }
            string vendeurCourant = String.Empty;
            ISheet sheet = hssfwb.GetSheet(SHEET_NAME);
            for (int row = FIRST_ROW; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    //La ligne contenant l'information du Vendeur ne contient pas d'information de client
                    String ligneVendeur = sheet.GetRow(row).GetCell(0)?.StringCellValue;
                    if(!string.IsNullOrEmpty(ligneVendeur))
                    {
                        vendeurCourant = ligneVendeur;
                    }
                    else if(!string.IsNullOrEmpty(sheet.GetRow(row).GetCell(1)?.StringCellValue) &&
                            !string.IsNullOrEmpty(sheet.GetRow(row).GetCell(1)?.StringCellValue))
                    {
                        ClientDataYear toInsert = new ClientDataYear()
                        {
                            Name = sheet.GetRow(row).GetCell(1).StringCellValue,
                            Montant = sheet.GetRow(row).GetCell(4)?.NumericCellValue.ToString() ?? String.Empty,
                            Year = year,
                            Vendeur = vendeurCourant
                        };
                        data.Add(toInsert);
                    }
                }
            }

            data = data.OrderBy(o => o.Name).ToList();

            return data;
        }

        static List<VendeurData> CombineData(List<List<ClientDataYear>> dataYear)
        {
            //trier les clients des diverses annees ensemble
            List<ClientData> dataAllClientsAllYears = new List<ClientData>();
            foreach(var annualData in dataYear)
            {
                foreach(var clientData in annualData)
                {
                    if(dataAllClientsAllYears.Any(o=>o.Name == clientData.Name))
                    {
                        var clientAllYearsData = dataAllClientsAllYears.First(o => o.Name == clientData.Name);
                        clientAllYearsData.Montants.Add(clientData);
                        clientAllYearsData.VendeurCourant = clientData.Vendeur;
                    }
                    else
                    {
                        var clientAllYearsData = new ClientData();
                        clientAllYearsData.Name = clientData.Name;
                        clientAllYearsData.Montants.Add(clientData);
                        clientAllYearsData.VendeurCourant = clientData.Vendeur;
                        dataAllClientsAllYears.Add(clientAllYearsData);
                    }
                }
            }
            dataAllClientsAllYears = dataAllClientsAllYears.OrderBy(o => o.VendeurCourant).ToList();

            //trier les clients par vendeurs
            List<VendeurData> vendeurDataList = new List<VendeurData>();
            foreach(var clientDataAllYears in dataAllClientsAllYears)
            {
                if(vendeurDataList.Any(o=>o.VendeurName == clientDataAllYears.VendeurCourant))
                {
                    var vendeurData = vendeurDataList.First(o => o.VendeurName == clientDataAllYears.VendeurCourant);
                    vendeurData.Clients.Add(clientDataAllYears);
                }
                else
                {
                    var toAdd = new VendeurData();
                    toAdd.VendeurName = clientDataAllYears.VendeurCourant;
                    toAdd.Clients.Add(clientDataAllYears);
                    vendeurDataList.Add(toAdd);
                }
            }

            //mettre en ordre de vendeur, et dans chaque vendeur, en ordre de client
            vendeurDataList.ForEach(o => o.Clients = o.Clients.OrderBy(p => p.Name).ToList());
            vendeurDataList = vendeurDataList.OrderBy(o => o.VendeurName).ToList();
            return vendeurDataList;
        }

        static void CreateExportData(List<VendeurData> vendeursData, string fileFullName)
        {
            // https://stackoverflow.com/questions/19838743/trying-to-create-a-new-xlsx-file-using-npoi-and-write-to-it

            
            
            using (FileStream stream = new FileStream(fileFullName, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet(SHEET_NAME);
                ICreationHelper cH = wb.GetCreationHelper();
                int rowIndex = FIRST_ROW;
                foreach (var vendeur in vendeursData)
                {
                    var vendeurCourant = vendeur.VendeurName;
                    IRow rowVendeur = sheet.CreateRow(rowIndex);
                    ICell cellVendeur = rowVendeur.CreateCell(0);
                    cellVendeur.SetCellValue(cH.CreateRichTextString(vendeurCourant));
                    //merged region https://scottstoecker.wordpress.com/2011/05/18/merging-excel-cells-with-npoi/
                    NPOI.SS.Util.CellRangeAddress craVendeur = new NPOI.SS.Util.CellRangeAddress(rowIndex, rowIndex, 0, 4);
                    sheet.AddMergedRegion(craVendeur);
                    rowIndex++;
                    foreach (var clientDataAllYears in vendeur.Clients)
                    {
                        IRow rowClient = sheet.CreateRow(rowIndex);
                        ICell cellClientName = rowClient.CreateCell(1);
                        cellClientName.SetCellValue(cH.CreateRichTextString(clientDataAllYears.Name));
                        NPOI.SS.Util.CellRangeAddress craClient = new NPOI.SS.Util.CellRangeAddress(rowIndex, rowIndex, 1, 3);
                        sheet.AddMergedRegion(craClient);
                        InsertCellForYear(cH, clientDataAllYears, rowClient, "2016", 4);
                        InsertCellForYear(cH, clientDataAllYears, rowClient, "2017", 5);
                        InsertCellForYear(cH, clientDataAllYears, rowClient, "2018", 6);
                        rowIndex++;
                    }

                    rowIndex += 2;
                }
                wb.Write(stream);
            }
        }

        private static void InsertCellForYear(ICreationHelper cH, ClientData clientDataAllYears, IRow rowClient, string year, int columnIndex)
        {
            if (clientDataAllYears.Montants.Any(o => o.Year == year))
            {
                ICell cellClientyear = rowClient.CreateCell(columnIndex);
                var clientYearData = clientDataAllYears.Montants.First(o => o.Year == year);
                cellClientyear.SetCellValue(cH.CreateRichTextString(clientYearData.Montant));
            }
        }
    }

    public class VendeurData
    {
        public string VendeurName;
        public List<ClientData> Clients;

        public VendeurData()
        {
            VendeurName = String.Empty;
            Clients = new List<ClientData>();
        }
    }

    public class ClientData
    {
        public string Name;
        public string VendeurCourant;
        public List<ClientDataYear> Montants;

        public ClientData()
        {
            Name = String.Empty;
            VendeurCourant = String.Empty;
            Montants = new List<ClientDataYear>();
        }
    }

    public class ClientDataYear
    {
        public string Name;
        public string Montant;
        public string Year;
        public string Vendeur;
    }
}
