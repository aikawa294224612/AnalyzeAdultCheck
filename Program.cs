using Hl7.Fhir.Model;
using Hl7.Fhir.Rest;
using OfficeOpenXml;
using System;
using System.Net.Http.Headers;
using System.Xml.Linq;

namespace TestAdultCheck
{
    class Program
    {
        static DataService secret = new DataService(); 
        static string root = secret.root;

        static async System.Threading.Tasks.Task Main(string[] args)
        {
            string token = secret.token;
            string org_id = secret.org_id;
            string fhirserver = secret.fhirserver;
            string excelFilePath = root + "延平衛生所_50筆.xlsx";
            int index = 2;

            var handler = new AuthorizationMessageHandler();
            handler.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var client = new FhirClient(fhirserver, FhirClientSettings.CreateDefault(), handler);
            
            FileInfo excelFile = new FileInfo(excelFilePath);
            if (excelFile.Exists)
            {
                excelFile.Delete();
                excelFile = new FileInfo(excelFilePath);
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(excelFile);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
            InitialExcel(worksheet);  //初始化: 產生表頭

            Bundle results = client.Search<Composition>(new string[] { "author=Organization/" + org_id });
            Console.WriteLine("總比數:"+ results.Entry.Count);

            foreach (Bundle.EntryComponent entry in results.Entry)
            {
                Composition comp = (Composition)entry.Resource;
                FillExcel(comp, worksheet, index);
                index++; 
            }
            package.Save();
        }

        public class AuthorizationMessageHandler : HttpClientHandler
        {
            public System.Net.Http.Headers.AuthenticationHeaderValue Authorization { get; set; }
            protected async override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                if (Authorization != null)
                    request.Headers.Authorization = Authorization;
                return await base.SendAsync(request, cancellationToken);
            }
        }

        public static void InitialExcel(ExcelWorksheet worksheet)
        {
            addExcel(worksheet, 1, "Composition.Id", "就醫序號", "Composition section數", "Patient", "Encounter", "衛生所", "檢驗單位",
                 "高血壓", "糖尿病", "高血脂症", "心臟病", "腦中風", "腎臟病",
                 "身高", "體重", "BMI", "血壓", "腰圍",
                 "Protein 尿蛋白", "血糖", "膽固醇", "三酸甘油脂", "低密度脂蛋白膽固醇計算", "高密度膽固醇",
                 "ＧＯＴ", "ＧＰＴ", "肌酸酐", "腎絲球過濾率計算", "Ｂ型肝炎表面抗原", "Ｃ型肝炎病毒抗體",
                 "憂鬱檢測", "吸煙", "喝酒", "嚼檳榔", "運動",
                 "健康諮詢：戒煙", "健康諮詢：節酒", "健康諮詢：戒檳榔", "健康諮詢：規律運動",
                 "健康諮詢：維持正常體重", "健康諮詢：健康飲食", "健康諮詢：事故傷害預防",
                 "健康諮詢：口腔保健", "檢查過B、C型肝炎", "血壓檢查結果與建議", "血糖檢查結果判讀",
                 "血脂肪檢查結果判讀", "腎功能檢查結果判讀", "肝功能檢查結果判讀", "代謝症候群檢查結果與建議",
                 "B型肝炎檢查結果與建議", "C型肝炎檢查結果與建議", "憂鬱檢測結果與建議");
        }

        public static void FillExcel(Composition comp, ExcelWorksheet worksheet, int index)
        {
            string com_id = comp.Id;
            string num = comp.Section.Count.ToString();

            string org2 = "";
            string hypertension = "";
            string diabetes = "";
            string hyperlipidemia = "";
            string heartDisease = "";
            string stroke = "";
            string kidneyDisease = "";
            string height = "";
            string weight = "";
            string bmi = "";
            string bloodPressure = "";
            string waistCircumference = "";
            string urineProtein = "";
            string bloodSugar = "";
            string cholesterol = "";
            string triglycerides = "";
            string ldlCholesterol = "";
            string hdlCholesterol = "";
            string got = "";
            string gpt = "";
            string creatinine = "";
            string egfr = "";
            string hepatitisB = "";
            string hepatitisC = "";
            string depressionScreening = "";
            string smoking = "";
            string alcoholConsumption = "";
            string betelNutChewing = "";
            string exercise = "";
            string quitSmokingConsultation = "";
            string reduceAlcoholConsultation = "";
            string quitBetelNutConsultation = "";
            string regularExerciseConsultation = "";
            string maintainNormalWeightConsultation = "";
            string healthyDietConsultation = "";
            string accidentInjuryPreventionConsultation = "";
            string oralCareConsultation = "";
            string checkedHepatitisBC = "";
            string bloodPressureCheckResult = "";
            string bloodSugarCheckResult = "";
            string lipidCheckResult = "";
            string kidneyFunctionCheckResult = "";
            string liverFunctionCheckResult = "";
            string metabolicSyndromeCheckResult = "";
            string hepatitisBCheckResult = "";
            string hepatitisCCheckResult = "";
            string depressionScreeningResult = "";


            string patient = comp.Subject.Reference.ToString();
            string enc = comp.Encounter.Reference.ToString();
            string org1 = comp.Author[0].Reference.ToString();
            string id = comp.Identifier.Value.ToString();
            if (comp.Author.Count > 1)
            {
                org2 = comp.Author[1].Reference.ToString();
            }

            foreach (var sec in comp.Section)
            {
                switch (sec.Title)
                {
                    case "疾病史-高血壓":
                        hypertension = sec.Entry[0].Reference.ToString();
                        break;
                    case "疾病史-糖尿病":
                        diabetes = sec.Entry[0].Reference.ToString();
                        break;
                    case "疾病史-高血脂症":
                        hyperlipidemia = sec.Entry[0].Reference.ToString();
                        break;
                    case "疾病史-心臟病":
                        heartDisease = sec.Entry[0].Reference.ToString();
                        break;
                    case "疾病史-腦中風":
                        stroke = sec.Entry[0].Reference.ToString();
                        break;
                    case "疾病史-腎臟病":
                        kidneyDisease = sec.Entry[0].Reference.ToString();
                        break;
                    case "生理量測-身高":
                        height = sec.Entry[0].Reference.ToString();
                        break;
                    case "生理量測-體重":
                        weight = sec.Entry[0].Reference.ToString();
                        break;
                    case "生理量測-BMI":
                        bmi = sec.Entry[0].Reference.ToString();
                        break;
                    case "生理量測-血壓":
                        bloodPressure = sec.Entry[0].Reference.ToString();
                        break;
                    case "生理量測-腰圍":
                        waistCircumference = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-Protein 尿蛋白":
                        urineProtein = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-血糖":
                        bloodSugar = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-膽固醇":
                        cholesterol = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-三酸甘油脂":
                        triglycerides = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-低密度脂蛋白膽固醇計算":
                        ldlCholesterol = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-高密度膽固醇":
                        hdlCholesterol = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-ＧＯＴ":
                        got = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-ＧＰＴ":
                        gpt = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-肌酸酐":
                        creatinine = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-腎絲球過濾率計算":
                        egfr = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-Ｂ型肝炎表面抗原":
                        hepatitisB = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢驗檢查-Ｃ型肝炎病毒抗體":
                        hepatitisC = sec.Entry[0].Reference.ToString();
                        break;
                    case "憂鬱檢測":
                        depressionScreening = sec.Entry[0].Reference.ToString();
                        break;
                    case "生活史-吸煙":
                        smoking = sec.Entry[0].Reference.ToString();
                        break;
                    case "生活史-喝酒":
                        alcoholConsumption = sec.Entry[0].Reference.ToString();
                        break;
                    case "生活史-嚼檳榔":
                        betelNutChewing = sec.Entry[0].Reference.ToString();
                        break;
                    case "生活史-運動":
                        exercise = sec.Entry[0].Reference.ToString();
                        break;
                    case "健康諮詢：戒煙":
                        quitSmokingConsultation = sec.Entry[0].Reference.ToString();
                        break;
                    case "健康諮詢：節酒":
                        reduceAlcoholConsultation = sec.Entry[0].Reference.ToString();
                        break;
                    case "健康諮詢：戒檳榔":
                        quitBetelNutConsultation = sec.Entry[0].Reference.ToString();
                        break;
                    case "健康諮詢：規律運動":
                        regularExerciseConsultation = sec.Entry[0].Reference.ToString();
                        break;
                    case "健康諮詢：維持正常體重":
                        maintainNormalWeightConsultation = sec.Entry[0].Reference.ToString();
                        break;
                    case "健康諮詢：健康飲食":
                        healthyDietConsultation = sec.Entry[0].Reference.ToString();
                        break;
                    case "健康諮詢：事故傷害預":
                        accidentInjuryPreventionConsultation = sec.Entry[0].Reference.ToString();
                        break;
                    case "健康諮詢：口腔保健":
                        oralCareConsultation = sec.Entry[0].Reference.ToString();
                        break;
                    case "檢查過B、C型肝炎":
                        checkedHepatitisBC = sec.Entry[0].Reference.ToString();
                        break;
                    case "血壓檢查結果與建議":
                        bloodPressureCheckResult = sec.Entry[0].Reference.ToString();
                        break;
                    case "血糖檢查結果判讀":
                        bloodSugarCheckResult = sec.Entry[0].Reference.ToString();
                        break;
                    case "血脂肪檢查結果判讀":
                        lipidCheckResult = sec.Entry[0].Reference.ToString();
                        break;
                    case "腎功能檢查結果判讀":
                        kidneyFunctionCheckResult = sec.Entry[0].Reference.ToString();
                        break;
                    case "肝功能檢查結果判讀":
                        liverFunctionCheckResult = sec.Entry[0].Reference.ToString();
                        break;
                    case "代謝症候群檢查結果與建議":
                        metabolicSyndromeCheckResult = sec.Entry[0].Reference.ToString();
                        break;
                    case "B型肝炎檢查結果與建議":
                        hepatitisBCheckResult = sec.Entry[0].Reference.ToString();
                        break;
                    case "C型肝炎檢查結果與建議":
                        hepatitisCCheckResult = sec.Entry[0].Reference.ToString();
                        break;
                    case "憂鬱檢測結果與建議":
                        depressionScreeningResult = sec.Entry[0].Reference.ToString();
                        break;
                }
                Console.WriteLine(com_id + "Finish!");
            }

            addExcel(worksheet, index, com_id, id, num, patient, enc, org1, org2,
                        hypertension, diabetes, hyperlipidemia, heartDisease, stroke, kidneyDisease,
                        height, weight, bmi, bloodPressure, waistCircumference,
                        urineProtein, bloodSugar, cholesterol, triglycerides,
                        ldlCholesterol, hdlCholesterol, got, gpt,
                        creatinine, egfr, hepatitisB, hepatitisC,
                        depressionScreening, smoking, alcoholConsumption, betelNutChewing,
                        exercise, quitSmokingConsultation, reduceAlcoholConsultation,
                        quitBetelNutConsultation, regularExerciseConsultation,
                        maintainNormalWeightConsultation, healthyDietConsultation,
                        accidentInjuryPreventionConsultation, oralCareConsultation,
                        checkedHepatitisBC, bloodPressureCheckResult,
                        bloodSugarCheckResult, lipidCheckResult, kidneyFunctionCheckResult,
                        liverFunctionCheckResult, metabolicSyndromeCheckResult,
                        hepatitisBCheckResult, hepatitisCCheckResult, depressionScreeningResult);
        }

        public static ExcelWorksheet addExcel(ExcelWorksheet worksheet, int index, 
            string com_id, string id, string count, string patient, string encounter, string hos, string test,
           string hypertension, string diabetes, string hyperlipidemia, string heartDisease, string stroke, string kidneyDisease,
            string height, string weight, string bmi, string bloodPressure, string waistCircumference,
            string urineProtein, string bloodSugar, string cholesterol, string triglycerides,
            string ldlCholesterol, string hdlCholesterol, string got, string gpt,
            string creatinine, string egfr, string hepatitisB, string hepatitisC,
            string depressionScreening, string smoking, string alcoholConsumption, string betelNutChewing,
            string exercise, string quitSmokingConsultation, string reduceAlcoholConsultation,
            string quitBetelNutConsultation, string regularExerciseConsultation,
            string maintainNormalWeightConsultation, string healthyDietConsultation,
            string accidentInjuryPreventionConsultation, string oralCareConsultation,
            string checkedHepatitisBC, string bloodPressureCheckResult,
            string bloodSugarCheckResult, string lipidCheckResult, string kidneyFunctionCheckResult,
            string liverFunctionCheckResult, string metabolicSyndromeCheckResult,
            string hepatitisBCheckResult, string hepatitisCCheckResult, string depressionScreeningResult) 
        {
            worksheet.Cells[index, 1].Value = com_id;
            worksheet.Cells[index, 2].Value = id;
            worksheet.Cells[index, 3].Value = count;
            worksheet.Cells[index, 4].Value = patient;
            worksheet.Cells[index, 5].Value = encounter;
            worksheet.Cells[index, 6].Value = hos;
            worksheet.Cells[index, 7].Value = test;
            worksheet.Cells[index, 8].Value = hypertension;
            worksheet.Cells[index, 9].Value = diabetes;
            worksheet.Cells[index, 10].Value = hyperlipidemia;
            worksheet.Cells[index, 11].Value = heartDisease;
            worksheet.Cells[index, 12].Value = stroke;
            worksheet.Cells[index, 13].Value = kidneyDisease;
            worksheet.Cells[index, 14].Value = height;
            worksheet.Cells[index, 15].Value = weight;
            worksheet.Cells[index, 16].Value = bmi;
            worksheet.Cells[index, 17].Value = bloodPressure;
            worksheet.Cells[index, 18].Value = waistCircumference;
            worksheet.Cells[index, 19].Value = urineProtein;
            worksheet.Cells[index, 20].Value = bloodSugar;
            worksheet.Cells[index, 21].Value = cholesterol;
            worksheet.Cells[index, 22].Value = triglycerides;
            worksheet.Cells[index, 23].Value = ldlCholesterol;
            worksheet.Cells[index, 24].Value = hdlCholesterol;
            worksheet.Cells[index, 25].Value = got;
            worksheet.Cells[index, 26].Value = gpt;
            worksheet.Cells[index, 27].Value = creatinine;
            worksheet.Cells[index, 28].Value = egfr;
            worksheet.Cells[index, 29].Value = hepatitisB;
            worksheet.Cells[index, 30].Value = hepatitisC;
            worksheet.Cells[index, 31].Value = depressionScreening;
            worksheet.Cells[index, 32].Value = smoking;
            worksheet.Cells[index, 33].Value = alcoholConsumption;
            worksheet.Cells[index, 34].Value = betelNutChewing;
            worksheet.Cells[index, 35].Value = exercise;
            worksheet.Cells[index, 36].Value = quitSmokingConsultation;
            worksheet.Cells[index, 37].Value = reduceAlcoholConsultation;
            worksheet.Cells[index, 38].Value = quitBetelNutConsultation;
            worksheet.Cells[index, 39].Value = regularExerciseConsultation;
            worksheet.Cells[index, 40].Value = maintainNormalWeightConsultation;
            worksheet.Cells[index, 41].Value = healthyDietConsultation;
            worksheet.Cells[index, 42].Value = accidentInjuryPreventionConsultation;
            worksheet.Cells[index, 43].Value = oralCareConsultation;
            worksheet.Cells[index, 44].Value = checkedHepatitisBC;
            worksheet.Cells[index, 45].Value = bloodPressureCheckResult;
            worksheet.Cells[index, 46].Value = bloodSugarCheckResult;
            worksheet.Cells[index, 47].Value = lipidCheckResult;
            worksheet.Cells[index, 48].Value = kidneyFunctionCheckResult;
            worksheet.Cells[index, 49].Value = liverFunctionCheckResult;
            worksheet.Cells[index, 50].Value = metabolicSyndromeCheckResult;
            worksheet.Cells[index, 51].Value = hepatitisBCheckResult;
            worksheet.Cells[index, 52].Value = hepatitisCCheckResult;
            worksheet.Cells[index, 53].Value = depressionScreeningResult;



            return worksheet;
        }

    }
}