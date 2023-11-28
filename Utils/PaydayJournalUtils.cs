using OfficeOpenXml;
using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;


namespace JournalGenerator.Utils
{
    internal class PaydayJournalUtils
    {
        internal static (byte[] outputBytes, int journalEntriesCount) GenerateJournalFromIslandsbankiForeignPayments(byte[] fileBytes)
        {
            string json = @"{
	        'settings':[{
		        'senderTitle':'AIRBNB PAYMENTS LUXEMBOURG S.A.',
		        'value': 'Airbnb Experience',
		        'commission': '0.2',
		        'accountNumber': '2381',
                'bankCommission': '680'
	        },
	        {
		        'senderTitle':'GetYourGuide Deutsch',
		        'value': 'GetYourGuide',
		        'commission': '0.3',
		        'accountNumber': '2382',
                'bankCommission': '680'
	        },
	        {
		        'senderTitle':'VIATOR LIMITED 7',
		        'value': 'Viator',
		        'commission': '0.255',
		        'accountNumber': '2383',
                'bankCommission': '680'
	        },
	        {
		        'senderTitle':'TRUST MY TRAVEL LIMITED',
		        'value': 'Bókun',
		        'commission': '0.015',
		        'accountNumber': '2384',
                'bankCommission': '680'
	        },
            {
		        'senderTitle':'TRUST MY TRAVEL LIMITED THE C',
		        'value': 'Bókun',
		        'commission': '0.015',
		        'accountNumber': '2384',
                'bankCommission': '0'
	        },
            {
		        'senderTitle':'Booknordics As',
		        'value': 'BookNordics AS',
		        'commission': '0.2',
		        'accountNumber': '2385',
                'bankCommission': '680'
	        },
            {
		        'senderTitle':'IPS EHF',
		        'value': 'Iceland Pro Services',
		        'commission': '0.15',
		        'accountNumber': '2385',
                'bankCommission': '680'
	        }]
        }";

            PaydayJournalSettings settings = JsonConvert.DeserializeObject<PaydayJournalSettings>(json);

            List<PaydayJournalEntry> journalEntries = new();

            // Import
            using (MemoryStream memStream = new(fileBytes))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


                ExcelPackage excel = new(memStream);
                for (int worksheetNr = 0; worksheetNr < excel.Workbook.Worksheets.Count; worksheetNr++)
                {
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets[worksheetNr];
                    int? postfix = null;
                    bool isLast = false;

                    for (int rowNumber = 2; rowNumber <= worksheet.Dimension.Rows; rowNumber++)
                    {

                        // Find start of next invoice
                        string dateString = worksheet.GetValue<string>(rowNumber, 1);
                        string payer = worksheet.GetValue<string>(rowNumber, 2);
                        string reference = worksheet.GetValue<string>(rowNumber, 3);
                        string currencySymbol = worksheet.GetValue<string>(rowNumber, 4);
                        string sourceAmountString = worksheet.GetValue<string>(rowNumber, 5);
                        string destinationAmountString = worksheet.GetValue<string>(rowNumber, 6);
                        string status = worksheet.GetValue<string>(rowNumber, 7);


                        if (string.IsNullOrWhiteSpace(dateString) || string.IsNullOrWhiteSpace(payer) || string.IsNullOrWhiteSpace(destinationAmountString))
                        {
                            break;
                        }

                        DateTime date = DateTime.ParseExact(dateString, "dd.MM.yy", CultureInfo.InvariantCulture, DateTimeStyles.None);

                        int day = date.Day;
                        int month = date.Month;
                        int year = date.Year;

                        string ledgerEntryDate = $"{day}.{month}.{year}";

                        decimal destinationAmount = decimal.Parse(destinationAmountString.Replace(".", ""), CultureInfo.InvariantCulture);

                        PaydayJournalSetting paydayJournalSetting = settings.Settings.FirstOrDefault(x => x.SenderTitle.Equals(payer, StringComparison.InvariantCultureIgnoreCase));

                        if (paydayJournalSetting == null)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"Error: No settings found for {payer}");
                            Console.ResetColor();
                            continue;
                        }

                        // Check if we need to add a day apendix
                        string nextLineDateString = worksheet.GetValue<string>(rowNumber + 1, 1);
                        if (!string.IsNullOrWhiteSpace(nextLineDateString))
                        {
                            DateTime nextLineDate = DateTime.ParseExact(nextLineDateString, "dd.MM.yy", CultureInfo.InvariantCulture, DateTimeStyles.None);

                            int nextLineDay = nextLineDate.Day;
                            int nextLineMonth = nextLineDate.Month;
                            int nextLineYear = nextLineDate.Year;
                            string nextLinePayer = worksheet.GetValue<string>(rowNumber + 1, 2);
                            PaydayJournalSetting nextLinePaydayJournalSetting = settings.Settings.FirstOrDefault(x => x.SenderTitle.Equals(nextLinePayer, StringComparison.InvariantCultureIgnoreCase));


                            if (day == nextLineDay && month == nextLineMonth && year == nextLineYear && !string.IsNullOrWhiteSpace(paydayJournalSetting.Value) && paydayJournalSetting.Value == nextLinePaydayJournalSetting.Value)
                            {
                                postfix = postfix == null ? 1 : postfix + 1;
                            }
                            else if (postfix != null)
                            {
                                isLast = true;
                                postfix = postfix + 1;
                            }
                            else
                            {
                                postfix = null;
                            }
                        }
                        else
                        {
                            postfix = null;
                        }

                        string description = $"{paydayJournalSetting.Value} - {day}.{month}.{year}" + (postfix != null && postfix.HasValue ? $" - {postfix.Value.ToString()}" : "");

                        if (isLast)
                        {
                            isLast = false;
                            postfix = null;
                        }


                        decimal amount = Math.Round(destinationAmount, 0, MidpointRounding.AwayFromZero);
                        decimal amountInclVat = Math.Round(paydayJournalSetting.Commission != 1 ? (destinationAmount / (1 - paydayJournalSetting.Commission)) : 0, MidpointRounding.AwayFromZero);
                        decimal vat = amountInclVat - amount;

                        journalEntries.Add(new PaydayJournalEntry
                        {
                            EntryNr = rowNumber - 1,
                            Date = ledgerEntryDate,
                            DateTimeValue = new DateTime(year, month, day),
                            Description = description,
                            Type = 1,
                            Key = "1110",
                            Amount = amountInclVat * -1, // Credit
                            VAT = "11",
                            //Reference = reference,
                            ReceiverKey = string.Empty
                        });


                        journalEntries.Add(new PaydayJournalEntry
                        {
                            EntryNr = rowNumber - 1,
                            Date = ledgerEntryDate,
                            Description = description,
                            Type = 1,
                            Key = paydayJournalSetting.AccountNumber,
                            Amount = vat,
                            VAT = "",
                            Reference = reference,
                            ReceiverKey = string.Empty
                        });

                        journalEntries.Add(new PaydayJournalEntry
                        {
                            EntryNr = rowNumber - 1,
                            Date = ledgerEntryDate,
                            Description = description,
                            Type = 1,
                            Key = "3200",
                            Amount = amount,
                            VAT = "",
                            Reference = reference,
                            ReceiverKey = string.Empty
                        });

                        journalEntries.Add(new PaydayJournalEntry
                        {
                            EntryNr = rowNumber - 1,
                            Date = ledgerEntryDate,
                            Description = description,
                            Type = 1,
                            Key = "2990",
                            Amount = paydayJournalSetting.BankCommission > 0 ? paydayJournalSetting.BankCommission : 0,
                            VAT = "",
                            Reference = reference,
                            ReceiverKey = string.Empty
                        });

                        journalEntries.Add(new PaydayJournalEntry
                        {
                            EntryNr = rowNumber - 1,
                            Date = ledgerEntryDate,
                            Description = description,
                            Type = 1,
                            Key = "3200",
                            Amount = paydayJournalSetting.BankCommission > 0 ? paydayJournalSetting.BankCommission * -1 : 0, // Credit
                            VAT = "",
                            Reference = reference,
                            ReceiverKey = string.Empty
                        });
                    }
                }
            }

            // Export
            if (journalEntries.Count > 0)
            {
                journalEntries = (from je in journalEntries
                                  orderby je.EntryNr descending
                                  select je).ToList();
            }

            using MemoryStream stream = new();
            using (ExcelPackage package = new(stream))
            {
                int rowNumber = 1;
                int numberOfColumns = 9;

                int sheetNr = 1;
                int entryNr = 1;

                foreach (var batch in journalEntries.Chunk(9 * 5)) // 9 Entries, each 5 lines
                {
                    var w = package.Workbook.Worksheets.Add($"Sheet{sheetNr}");
                    sheetNr++;

                    w.Cells[rowNumber, 1].Value = "Færsla nr.";
                    w.Cells[rowNumber, 2].Value = "Dags.";
                    w.Cells[rowNumber, 3].Value = "Lýsing";
                    w.Cells[rowNumber, 4].Value = "Tegund (1 = Fjárhagur, 2 = Viðskiptavinur, 3 = Lánardrottinn)";
                    w.Cells[rowNumber, 5].Value = "Lykill (fjárhagslykill eða kennitala)";
                    w.Cells[rowNumber, 6].Value = "Fjárhæð m/vsk";
                    w.Cells[rowNumber, 7].Value = "VSK (0, 11, 24) - tómt engin vsk";
                    w.Cells[rowNumber, 8].Value = "Tilvísun";
                    w.Cells[rowNumber, 9].Value = "Mótlykill";

                    foreach (var subBatch in batch.Chunk(5))
                    {
                        foreach (var journalEntry in subBatch)
                        {
                            if (journalEntry.Amount != 0)
                            {
                                rowNumber++;
                                w.Cells[rowNumber, 1].Value = entryNr;
                                w.Cells[rowNumber, 2].Value = journalEntry.Date;
                                w.Cells[rowNumber, 3].Value = journalEntry.Description;
                                w.Cells[rowNumber, 4].Value = journalEntry.Type;
                                w.Cells[rowNumber, 5].Value = journalEntry.Key;
                                w.Cells[rowNumber, 6].Value = journalEntry.Amount;
                                w.Cells[rowNumber, 7].Value = journalEntry.VAT;
                                //w.Cells[rowNumber, 8].Value = journalEntry.Reference;
                                w.Cells[rowNumber, 9].Value = journalEntry.ReceiverKey;
                            }
                        }
                        entryNr++;
                    }
                    w.Cells.AutoFitColumns(0); // Autofit columns for all cells
                    rowNumber = 1; // reset
                }

                package.Workbook.Properties.Title = "Journal " + DateTime.Now.ToString("yyyyMMdd");
                package.Save();
            }
            return (stream.ToArray(), journalEntries.Count / 5);
        }
    }

    public partial class PaydayJournalEntry
    {
        public int EntryNr { get; set; }
        public string Date { get; set; }
        public string Description { get; set; }
        public int Type { get; set; }
        public string Key { get; set; }
        public Decimal Amount { get; set; }
        public string VAT { get; set; }
        public string Reference { get; set; }
        public string ReceiverKey { get; set; }
        public DateTime DateTimeValue { get; set; }
    }


    public partial class PaydayJournalSettings
    {
        [JsonProperty("settings")]
        public List<PaydayJournalSetting> Settings { get; set; }
    }

    public partial class PaydayJournalSetting
    {
        [JsonProperty("senderTitle")]
        public string SenderTitle { get; set; }

        [JsonProperty("value")]
        public string Value { get; set; }

        [JsonProperty("commission")]
        public decimal Commission { get; set; }

        [JsonProperty("accountNumber")]
        public string AccountNumber { get; set; }

        [JsonProperty("bankCommission")]
        public decimal BankCommission { get; set; }
    }

    public partial class Temperatures
    {
        public static Temperatures FromJson(string json) => JsonConvert.DeserializeObject<Temperatures>(json, Converter.Settings);
    }

    public static class Serialize
    {
        public static string ToJson(this Temperatures self) => JsonConvert.SerializeObject(self, Converter.Settings);
    }

    internal static class Converter
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters =
            {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }

    internal class ParseStringConverter : JsonConverter
    {
        public override bool CanConvert(Type t) => t == typeof(long) || t == typeof(long?);

        public override object ReadJson(JsonReader reader, Type t, object existingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.Null) return null;
            var value = serializer.Deserialize<string>(reader);
            long l;
            if (Int64.TryParse(value, out l))
            {
                return l;
            }
            throw new Exception("Cannot unmarshal type long");
        }

        public override void WriteJson(JsonWriter writer, object untypedValue, JsonSerializer serializer)
        {
            if (untypedValue == null)
            {
                serializer.Serialize(writer, null);
                return;
            }
            var value = (long)untypedValue;
            serializer.Serialize(writer, value.ToString());
            return;
        }
    }
}
