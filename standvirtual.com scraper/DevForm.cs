using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using scrapingTemplateV51.Models;
using standvirtual.com_scraper.Models;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace standvirtual.com_scraper
{
    public partial class DevForm : DevExpress.XtraEditors.XtraForm
    {
        private List<Car> cars = new List<Car>();
        private Config _config = new Config();
        public HttpCaller _httpCaller = new HttpCaller();
        public List<InputModel> inputModels = new List<InputModel>();
        public DevForm()
        {
            InitializeComponent();
        }
        public delegate void DisplayD(string s);
        public void Display(string s)
        {
            if (InvokeRequired)
            {
                Invoke(new DisplayD(Display), s);
                return;
            }
            DispalyArea.Text = s;
        }
        public delegate void SetProgressD(int x);
        public void SetProgress(int x)
        {
            if (InvokeRequired)
            {
                Invoke(new SetProgressD(SetProgress), x);
                return;
            }
            if ((x <= 100))
            {
                progressB.EditValue = x;
            }
        }
        private async void simpleButton1_Click(object sender, EventArgs e)
        {
            //var car = await ScrapeCarDetails("https://www.standvirtual.com/anuncio/bmw-i3-94ah-comfort-package-advance-ID8OXwcr.html#80a0ec84c1");
            //return;
            if (!Daily.Checked && !ThreeDays.Checked)
            {
                MessageBox.Show(@"Please select time base scraping ""Daily"" or ""3 days"" option ");
                return;
            }

            do
            {
                inputModels = new List<InputModel>();
                cars = new List<Car>();
                for (int i = 0; i < FiltersDGV.RowCount; i++)
                {
                    var input = FiltersDGV.GetRow(i) as InputModel;
                    inputModels.Add(input);
                }
                var d2 = new DateTime();
                var days = 1;
                if (Daily.Checked)
                {
                    d2 = DateTime.Now.AddDays(1);
                }
                if (ThreeDays.Checked)
                {
                    days = 3;
                    d2 = DateTime.Now.AddDays(3);
                }
                await MainWork();
                var d1 = DateTime.Now;
                Display($@"work done for today next run will be {DateTime.Now.AddDays(days):dd/MM/yyyy} ");
                await Task.Delay(d2 - d1);
            } while (true);
        }

        private async Task MainWork()
        {
            foreach (var inputModel in inputModels)
            {
                await StartScraping(inputModel);
            }
            await SaveData();
        }

        private async Task StartScraping(InputModel inputModel)
        {
            var make = "";
            if (inputModel.Make != null)
            {
                make = inputModel.Make.SearchKey;
            }
            var model = "";
            var modelName = "";
            if (inputModel.Model != null)
            {
                var brand = _config.Makes.Find(x => x.SearchKey == make);
                var mod = brand.Models.Find(xx => xx.SearchKey == inputModel.Model.SearchKey);
                if (mod != null)
                {
                    model = inputModel.Model.SearchKey;
                    modelName = inputModel.Model.Name;
                }
            }
            var vat = "";
            if (inputModel.Vat)
            {
                vat = "1";
            }

            var formData = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("search[order]: created_at","desc"),
                new KeyValuePair<string, string>("view","list"),
                new KeyValuePair<string, string>("min_id",""),
                new KeyValuePair<string, string>("search[category_id]","29"),
                new KeyValuePair<string, string>("search[filter_enum_make][]",make),
                new KeyValuePair<string, string>("search[filter_enum_model]",model),
                new KeyValuePair<string, string>("search[filter_float_price:from]",inputModel.MinPrice),
                new KeyValuePair<string, string>("search[filter_float_price:from]_dd",inputModel.MinPrice),
                new KeyValuePair<string, string>("search[filter_float_price:to]",inputModel.MaxPrice),
                new KeyValuePair<string, string>("search[filter_float_price:to]_dd",inputModel.MaxPrice),
                new KeyValuePair<string, string>("search[filter_float_first_registration_year:from]",inputModel.FromDate),
                new KeyValuePair<string, string>("search[filter_float_first_registration_year:from]_dd",inputModel.FromDate),
                new KeyValuePair<string, string>("search[filter_float_first_registration_year:to]",inputModel.ToDate),
                new KeyValuePair<string, string>("search[filter_float_first_registration_year:to]_dd",inputModel.ToDate),
                new KeyValuePair<string, string>("search[filter_enum_fuel_type]","electric"),
                new KeyValuePair<string, string>("search[filter_float_mileage:from]",inputModel.MinKilometers),
                new KeyValuePair<string, string>("search[filter_float_mileage:from]_dd",inputModel.MinKilometers),
                new KeyValuePair<string, string>("search[filter_float_mileage:to]",inputModel.MaxKilometers),
                new KeyValuePair<string, string>("search[filter_float_mileage:to]_dd",inputModel.MaxKilometers),
                new KeyValuePair<string, string>("search[filter_float_power:from]",inputModel.BatteryCapacityFrom),
                new KeyValuePair<string, string>("search[filter_float_power:from]_dd",inputModel.BatteryCapacityFrom),
                new KeyValuePair<string, string>("search[filter_float_power:to]",inputModel.BatteryCapacityTo),
                new KeyValuePair<string, string>("search[filter_float_power:to]_dd",inputModel.BatteryCapacityTo),
                new KeyValuePair<string, string>("search[brand_program_id][]",""),
                new KeyValuePair<string, string>("q",""),
                new KeyValuePair<string, string>("search[filter_float_engine_capacity:from]",""),
                new KeyValuePair<string, string>("search[filter_float_engine_capacity:from]_dd",""),
                new KeyValuePair<string, string>("search[filter_float_engine_capacity:to]",""),
                new KeyValuePair<string, string>("search[filter_float_engine_capacity:to]_dd",""),
                new KeyValuePair<string, string>("search[filter_enum_particle_filter]",""),
                new KeyValuePair<string, string>("search[filter_enum_tax_deductible]",vat),
                new KeyValuePair<string, string>("search[map_zoom]",""),
                new KeyValuePair<string, string>("search[map_lon]",""),
                new KeyValuePair<string, string>("search[map_lat]",""),
                new KeyValuePair<string, string>("search[district_id]",""),
                new KeyValuePair<string, string>("search[cityName]",""),
                new KeyValuePair<string, string>("search[city_id]",""),
                new KeyValuePair<string, string>("search[subregion_id]",""),
                new KeyValuePair<string, string>("search[country]",""),
                new KeyValuePair<string, string>("search[region_id]",""),
                new KeyValuePair<string, string>("initiator","search_field"),
            };
            var html = await _httpCaller.PostFormData("https://www.standvirtual.com/ajax/search/list/", formData);
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);
            //doc.Save("test1.html");
            //Process.Start("test.html");
            //Process.Start("test1.html");
            var carNbr = int.Parse(doc.DocumentNode.SelectSingleNode("//span[text()='Todos']/following-sibling::span").InnerText.Trim().Replace("(", "")
                .Replace(")", ""));
            if (carNbr == 0)
            {
                Display($@"no results from filter {inputModel.Make.Name + " : " + inputModel.Model.Name}");
                await Task.Delay(3000);
                return;
            }
            var nextPage = "";
            var counter = 1;
            var page = 3;
            var finish = false;
            do
            {
                var carNodes =
                    doc.DocumentNode.SelectNodes("//div[@class='offers list']//a[@class='offer-title__link']") ??
                    doc.DocumentNode.SelectNodes("//article[@data-variant]//h2/a");
               
                foreach (var carNode in carNodes)
                {
                    var urlCar = carNode.GetAttributeValue("href", "");
                    var car = await ScrapeCarDetails(urlCar);
                    cars.Add(car);
                    if (cars.Count == carNbr)
                    {
                        finish = true;
                        break;
                    }
                    Display($@"{counter} car scraped/{carNbr}   from filter => Make: { inputModel.Make.Name + "    Model: " + modelName}");
                    SetProgress(((counter * 100) / (carNbr)));
                    counter++;
                }

                if (finish)
                {
                    break;
                }
                if (nextPage == "")
                {
                    nextPage = doc.DocumentNode.SelectSingleNode("//a[@rel='next']")?.GetAttributeValue("href", "").Replace("/?search", "?search");
                    doc = await _httpCaller.GetDoc(nextPage);
                }
                else
                {
                    var lastIndex = nextPage[nextPage.Length-1] + "";
                    nextPage = nextPage.Replace("lastIndex", page+"");
                    doc = await _httpCaller.GetDoc(nextPage);
                    page++;
                }
               
                //doc.Save("test.html");
                //Process.Start("test.html");
            } while (true);
        }

        private async Task<Car> ScrapeCarDetails(string urlCar)
        {
            var car = new Car();
            car.Url = urlCar;
            var doc = await _httpCaller.GetDoc(urlCar);

            var phoneCode = doc.DocumentNode.SelectSingleNode(" //span[@data-id] ")?.GetAttributeValue("data-id", "").Trim();
            var phones = new StringBuilder();
            var counter = 0;
            do
            {
                var json = await _httpCaller.GetHtml($@"https://www.standvirtual.com/ajax/misc/contact/multi_phone/{phoneCode}/{counter}/");
                if (json.Contains("encontrada"))
                {
                    break;
                }
                var obj = JObject.Parse(json);
                var phone = (string)obj.SelectToken("value");
                phones.Append(phone + " | ");
                counter++;
            } while (true);

            if (phones.Length > 0)
            {
                car.Phone = phones.ToString().Substring(0, phones.ToString().LastIndexOf('|'));
            }
            else
            {
                car.Phone = "N/A";
            }
            car.Make = doc.DocumentNode.SelectSingleNode("//div[@id='parameters']//span[text()='Marca']/../div").InnerText.Trim();
            var model = doc.DocumentNode.SelectSingleNode("//div[@id='parameters']//span[text()='Modelo']/../div").InnerText.Trim();
            var subModel = doc.DocumentNode.SelectSingleNode("//div[@id='parameters']//span[text()='Sub-modelo']/../div")?.InnerText.Trim() ?? "";
            var version = doc.DocumentNode.SelectSingleNode("//div[@id='parameters']//span[text()='Versão']/../div")?.InnerText.Trim() ?? "";
            car.Model = (model + ", " + subModel + ", " + version).Replace(", ,", " , ");
            car.Kilometre = doc.DocumentNode.SelectSingleNode("//div[@id='parameters']//span[text()='Quilómetros']/../div")?.InnerText.Trim() ?? "N/A";
            car.Year = doc.DocumentNode.SelectSingleNode("//div[@id='parameters']//span[text()='Ano de Registo']/../div")?.InnerText.Trim() ?? "N/A";
            car.Price = "€" + doc.DocumentNode.SelectSingleNode("//div[@data-price]").GetAttributeValue("data-price", "").Trim().Replace(" ", ",");

            return car;
        }

        private async Task ScrapeConfigurations()
        {
            var doc = await _httpCaller.GetDoc("https://www.standvirtual.com/");
            var json = doc.DocumentNode.SelectSingleNode("//script[@id='__NEXT_DATA__']").InnerText.Trim();

            var obj = JObject.Parse(json);
            var brands = obj.SelectToken("..filtersValues.1[1].group_values");
            var categories = obj.SelectToken("..filtersValues")?.ToList();
            var makes = new List<Make>();
            #region Collect Makes
            foreach (var brand in brands)
            {
                var make = new Make();
                make.Id = (string)brand.SelectToken("value_key");
                var title = (string)brand.SelectToken("name");
                var x = title.LastIndexOf(" ", StringComparison.Ordinal);
                make.Name = title.Substring(0, x);
                make.SearchKey = (string)brand.SelectToken("search_key");
                for (var i = 6; i < categories.Count; i++)
                {
                    var catgName = ((JProperty)categories[i]).Name;
                    if (catgName.Contains(make.Id ?? string.Empty))
                    {
                        var pModels = ((JToken)categories[i]).SelectToken("..group_values");
                        foreach (var pModel in pModels)
                        {
                            var model = new Model();
                            title = (string)pModel.SelectToken("name");
                            x = title.LastIndexOf(" ", StringComparison.Ordinal);
                            model.Name = title.Substring(0, x);
                            model.Id = (string)pModel.SelectToken("value_key");
                            model.SearchKey = (string)pModel.SelectToken("search_key");
                            var subCtgIsFound = false;
                            for (var j = 6; j < categories.Count; j++)
                            {
                                var subCategory = (JProperty)categories[j];
                                if (subCategory.Name.Contains(model.Id))
                                {
                                    subCtgIsFound = true;
                                    foreach (var sub in ((JToken)subCategory)?.SelectToken("..group_values"))
                                    {
                                        var subSearchKey = (string)sub.SelectToken("search_key");
                                        title = (string)sub.SelectToken("name");
                                        x = title.LastIndexOf(" ", StringComparison.Ordinal);
                                        var name = title.Substring(0, x);
                                        make.Models.Add(new Model { Name = name, SearchKey = subSearchKey + "|" + model.SearchKey });
                                    }
                                }
                            }

                            make.Models.Add(model);
                        }
                    }
                }
                makes.Add(make);
            }
            #endregion
            var datesFilters = obj.SelectToken("..filtersValues.2006..group_values");
            var dates = new List<string>();
            foreach (var datesFilter in datesFilters)
            {
                var date = (string)datesFilter.SelectToken("name");
                dates.Add(date);
            }

            _config.Makes = makes;
            _config.Dates = dates;
            File.WriteAllText("conff", JsonConvert.SerializeObject(_config, Formatting.Indented));
        }

        private async void DevForm_Load(object sender, EventArgs e)
        {
            //var url =
            //    "https://www.standvirtual.com/carros/bmw/desde-1910?search%5Bfilter_enum_fuel_type%5D=electric&search%5Border%5D=desc&search%5Bfilter_float_first_registration_year%3Ato%5D=2021&page=4";
            //var doc = await _httpCaller.GetDoc(url);
            //doc.Save("test.html");
            //Process.Start("test.html");
            //if (!Directory.Exists("outcomes"))
            {
                Directory.CreateDirectory("outcomes");
            }
            ServicePointManager.DefaultConnectionLimit = 65000;
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //await ScrapeConfigurations();
            //return;
            _config = JsonConvert.DeserializeObject<Config>(File.ReadAllText("conff"));

            MakesRepositoryItemLookUpEdit.DataSource = _config.Makes;
            DateFromRepositoryItemLookUpEdit.DataSource = _config.Dates;
            DateToRepositoryItemLookUpEdit.DataSource = _config.Dates;

            var d = JsonConvert.DeserializeObject<List<InputModel>>(File.ReadAllText("vv"));
            GridControle.DataSource = new BindingList<InputModel>(d);
        }
        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString(), @"Unhandled Thread Exception");
        }
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show((e.ExceptionObject as Exception)?.ToString(), @"Unhandled UI Exception");
        }
        private void DevForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            inputModels = new List<InputModel>();
            for (int i = 0; i < FiltersDGV.RowCount; i++)
            {
                var input = FiltersDGV.GetRow(i) as InputModel;
                inputModels.Add(input);
            }
            File.WriteAllText("vv", JsonConvert.SerializeObject(inputModels, Formatting.Indented));
        }

        private void FiltersDGV_CustomRowCellEdit(object sender, DevExpress.XtraGrid.Views.Grid.CustomRowCellEditEventArgs e)
        {
            if (e.Column.FieldName.Equals("Model"))
            {

                var inputModel = FiltersDGV.GetRow(e.RowHandle) as InputModel;
                if (inputModel == null) return;
                if (inputModel.Make == null)
                {
                    var r2 = new RepositoryItemLookUpEdit();
                    e.RepositoryItem = r2;
                    e.Column.ColumnEdit = r2;
                    return;
                }
                var r = new RepositoryItemLookUpEdit();
                r.DataSource = inputModel.Make.Models;
                r.Columns.Add(new LookUpColumnInfo { FieldName = "Name" });
                r.DisplayMember = "Name";
                GridControle.RepositoryItems.Add(r);
                e.RepositoryItem = r;
                e.Column.ColumnEdit = r;
            }
        }
        private async Task SaveData()
        {
            var date = DateTime.Now.ToString("dd_MM_yyyy");
            var path = $@"outcomes\standvirtual.com{date}.xlsx";
            var excelPkg = new ExcelPackage(new FileInfo(path));

            var sheet = excelPkg.Workbook.Worksheets.Add("Cars");
            sheet.Protection.IsProtected = false;
            sheet.Protection.AllowSelectLockedCells = false;
            sheet.Row(1).Height = 20;
            sheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Row(1).Style.Font.Bold = true;
            sheet.Row(1).Style.Font.Size = 8;
            sheet.Cells[1, 1].Value = "Make";
            sheet.Cells[1, 2].Value = "Model";
            sheet.Cells[1, 3].Value = "Price";
            sheet.Cells[1, 4].Value = "Phone";
            sheet.Cells[1, 5].Value = "Year";
            sheet.Cells[1, 6].Value = "Kilometers";
            sheet.Cells[1, 7].Value = "Weblink";

            var range = sheet.Cells[$"A1:G{cars.Count + 1}"];
            var tab = sheet.Tables.Add(range, "");

            tab.TableStyle = TableStyles.Medium2;
            sheet.Cells.Style.Font.Size = 12;

            var row = 2;
            foreach (var car in cars)
            {

                sheet.Cells[row, 1].Value = car.Make;
                sheet.Cells[row, 2].Value = car.Model;
                sheet.Cells[row, 3].Value = car.Price;
                sheet.Cells[row, 4].Value = car.Phone;
                sheet.Cells[row, 5].Value = car.Year;
                sheet.Cells[row, 6].Value = car.Kilometre;
                sheet.Cells[row, 7].Value = car.Url;
                row++;
            }

            for (int i = 2; i <= sheet.Dimension.End.Column; i++)
                sheet.Column(i).AutoFit();

            sheet.View.FreezePanes(2, 1);
            await excelPkg.SaveAsync();

        }
    }
}