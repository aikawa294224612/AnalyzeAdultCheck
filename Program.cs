﻿using Hl7.Fhir.Model;
using Hl7.Fhir.Rest;
using OfficeOpenXml;
using System.Net.Http.Headers;

namespace TestAdultCheck
{
    class Program
    {
        static DataService secret = new DataService(); 
        static LogicModels logic = new LogicModels();
        static string root = secret.root;

        static async System.Threading.Tasks.Task Main(string[] args)
        {

            string token = secret.token;
            string[] orgIds = { "177246" };  //海端177246  //延平2
            string fhirserver = secret.fhirserver;
            string monthanddate = System.DateTime.Now.ToString("MMdd");
            string excelFilePath = root + "台東衛生所成健_" + monthanddate + ".xlsx";

            var handler = new AuthorizationMessageHandler();
            handler.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var client = new FhirClient(fhirserver, FhirClientSettings.CreateDefault(), handler);

            FileInfo excelFile = new FileInfo(excelFilePath);
            if (excelFile.Exists)
            {
                excelFile.Delete();
                excelFile = new FileInfo(excelFilePath);
            }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(excelFile);
            try
            {
                foreach (string id in orgIds)
                {
                    // ExpiredContinue(package, client, id, "94813174-bbec-4384-9688-9e7caddb9b41", "3170", 3172);
                    int index = 2;

                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(logic.GetOrgName(id));
                    InitialExcel(worksheet);  //初始化: 產生表頭

                    var searchParams = new SearchParams();
                    searchParams.Add("_total", "accurate");
                    searchParams.Add("author", "Organization/" + id);
                    searchParams.Count = 10;

                    Bundle results = client.Search<Composition>(searchParams);

                    Console.WriteLine(logic.GetOrgName(id) + "總比數: " + results.Total);

                    int? total = results.Total;

                    while (index - 1 <= total)
                    {
                        foreach (Bundle.EntryComponent entry in results.Entry)
                        {
                            Composition comp = (Composition)entry.Resource;
                            FillExcel(comp, worksheet, index, client, true);  //取得資料且寫入excel
                            index++;
                        }
                        results = logic.GetNextPages(results, client);
                    }

                }
                package.Save();  //儲存excel
            }
            catch (Exception e)
            {
                Console.WriteLine("Error:" + e.Message);
            }
            finally
            {
                package.Save();  //儲存excel
            }

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

        public static void ExpiredContinue(ExcelPackage package, FhirClient client, 
            string id, string pageId, string offset, int index)
        {

            ExcelWorksheet worksheet = package.Workbook.Worksheets[logic.GetOrgName(id)];

            var searchParams = new SearchParams();
            searchParams.Add("_total", "accurate");
            searchParams.Add("_getpages", pageId);
            searchParams.Add("_getpagesoffset", offset);  //Important
            searchParams.Add("_count", "10");
            searchParams.Add("_pretty", "true");
            searchParams.Add("_bundletype", "searchset");

            Bundle results = client.Search(searchParams);

            Console.WriteLine(logic.GetOrgName(id) + "總比數: " + results.Total);

            int? total = results.Total;

            while (index - 1 <= total)
            {
                foreach (Bundle.EntryComponent entry in results.Entry)
                {
                    Composition comp = (Composition)entry.Resource;
                    FillExcel(comp, worksheet, index, client, true);  //取得資料且寫入excel
                    index++;
                }
                results = logic.GetNextPages(results, client);
            }

        }

        public static void InitialExcel(ExcelWorksheet worksheet)
        {
            addExcel(worksheet, 1, "姓名", "身分證號", "出生日期", "性別", "電話", "第一階段檢查日期",
                "第二階段檢查日期", "檢查結果上傳日期", "檢查過B、C型肝炎", "疾病史：高血壓",
                "疾病史：糖尿病", "疾病史：高血脂症", "心臟病", "腦中風", "腎臟病", "血壓 -- 收縮壓",
                "血壓 -- 舒張壓", "高血壓", "三高", "腰圍", "BMI", "吸煙", "喝酒", "嚼檳榔", "運動",
                "健康諮詢：戒煙", "健康諮詢：節酒", "健康諮詢：戒檳榔", "健康諮詢：規律運動", "健康諮詢：維持正常體",
                "健康諮詢：健康飲食", "健康諮詢：事故傷害預", "健康諮詢：口腔保健", "B型肝炎表面抗原", "C型肝炎病毒抗體",
                "感覺情緒低落沮喪", "做事情失去興趣", "尿液--酸鹼值", "Protein 尿蛋白", "尿糖", "尿沉渣鏡檢", "潛血",
                "尿液紅血球", "尿液白血球", "尿液上皮細胞", "Cast", "Bacteria", "Appearance", "高血脂", "膽固醇",
                "三酸甘油脂", "空腹血糖", "肌酸肝", "ＧＯＴ", "ＧＰＴ", "高密度膽固醇", "低密度脂蛋白膽固醇計",
                "腎絲球過濾率計算", "肝功能檢查結果判讀", "血糖檢查結果判讀", "血脂肪檢查結果判讀", "腎功能檢查結果判讀",
                "血壓檢查結果與建議", "代謝症候群檢查結果與", "憂鬱檢測結果與建議", "Composition Id");

        }

        public static void FillExcel(Composition comp, ExcelWorksheet worksheet, int index, FhirClient client, bool decryp)
        {
            string com_id = comp.Id;

            string name = string.Empty, id = string.Empty, birthDate = string.Empty, gender = string.Empty,
                phone = string.Empty, firstCheckDate = string.Empty, secondCheckDate = string.Empty,
                resultUploadDate = string.Empty, checkBCTypeHepatitis = string.Empty,
                hypertensionHistory = "否", diabetesHistory = "否", hyperlipidemiaHistory = "否",
                heartDisease = "否", stroke = "否", kidneyDisease = "否", systolicPressure = string.Empty,
                diastolicPressure = string.Empty, highBloodPressure = string.Empty, threeHigh = string.Empty,
                waistCircumference = string.Empty, BMI = string.Empty, smoking = string.Empty,
                alcoholConsumption = string.Empty, betelNutChewing = string.Empty, exercise = string.Empty,
                smokingCessationConsultation = string.Empty, alcoholReductionConsultation = string.Empty,
                betelNutCessationConsultation = string.Empty, regularExerciseConsultation = string.Empty,
                maintainNormalWeightConsultation = string.Empty, healthyDietConsultation = string.Empty,
                accidentInjuryPreventionConsultation = string.Empty, oralHealthCareConsultation = string.Empty,
                hepatitisBSurfaceAntigen = "未執行", hepatitisCAntibody = "未執行", lowMood = string.Empty,
                lossOfInterest = string.Empty, urineAcidityValue = "未執行", urineProtein = string.Empty,
                urineSugar = "未執行", urineSedimentMicroscopy = "未執行", occultBlood = "未執行",
                urineRedBloodCells = "未執行", urineWhiteBloodCells = "未執行",
                urineEpithelialCells = "未執行", cast = "未執行", bacteria = "未執行",
                appearance = "未執行", hyperlipidemia = string.Empty, cholesterol = string.Empty,
                triglycerides = string.Empty, fastingBloodSugar = string.Empty, creatinine = string.Empty,
                GOT = string.Empty, GPT = string.Empty, highDensityLipoproteinCholesterol = string.Empty,
                lowDensityLipoproteinCholesterol = string.Empty, glomerularFiltrationRate = string.Empty,
                liverFunctionResultInterpretation = string.Empty, bloodSugarResultInterpretation = string.Empty,
                lipidProfileResultInterpretation = string.Empty, kidneyFunctionResultInterpretation = string.Empty,
                bloodPressureResultAndRecommendation = string.Empty, metabolicSyndromeResultAndRecommendation = string.Empty,
                depressionDetectionResultAndRecommendation = string.Empty;

            string hyperglycemiaTemp = string.Empty;


            string patientId = comp.Subject.Reference.ToString();
            Patient pat = client.Read<Patient>(patientId);
            if (pat != null)
            {
                name = decryp ? logic.DecryptStringFromBytes_Aes(pat.Name[0].Text) : pat.Name[0].Text; //姓名
                foreach (Identifier identifier in pat.Identifier)
                {
                    if (identifier.System == "http://www.moi.gov.tw/")
                    {
                        id = decryp? logic.DecryptStringFromBytes_Aes(identifier.Value) : identifier.Value; //身分證字號
                    }
                }
                birthDate = pat.BirthDate;
                gender = logic.ChangeGender(pat.Gender);
                if (pat.Telecom != null && pat.Telecom.Count > 0)
                {
                    phone = decryp ? logic.DecryptStringFromBytes_Aes(pat.Telecom[0].Value) : pat.Telecom[0].Value;  //電話
                }
                
            }

            string encId = comp.Encounter.Reference.ToString();
            Encounter enc = client.Read<Encounter>(encId);

            if (enc.StatusHistory != null && enc.StatusHistory.Count > 0)
            {
                Encounter.StatusHistoryComponent first = enc.StatusHistory[0];  //第一階段檢查日期
                if (enc.StatusHistory.Count > 1)
                {
                    Encounter.StatusHistoryComponent second = enc.StatusHistory[1]; //第二階段檢查日期
                    secondCheckDate = logic.AdToRocEra(second.Period.Start);
                }
                firstCheckDate = logic.AdToRocEra(first.Period.Start);
            }

            foreach (var sec in comp.Section)
            {
                switch (sec.Title)
                {
                    case "疾病史-高血壓":
                        hypertensionHistory = "是";
                        break;
                    case "疾病史-糖尿病":
                        diabetesHistory = "是";
                        break;
                    case "疾病史-高血脂症":
                        hyperlipidemiaHistory = "是";
                        break;
                    case "疾病史-心臟病":
                        heartDisease = "是";
                        break;
                    case "疾病史-腦中風":
                        stroke = "是";
                        break;
                    case "疾病史-腎臟病":
                        kidneyDisease = "是";
                        break;
                    //case "生理量測-身高":
                    //    height = sec.Entry[0].Reference.ToString();
                    //    break;
                    //case "生理量測-體重":
                    //    weight = sec.Entry[0].Reference.ToString();
                    //    break;
                    case "生理量測-BMI":
                        string bmiId = sec.Entry[0].Reference.ToString();
                        Observation bmi = client.Read<Observation>(bmiId);
                        if (bmi.Value is Hl7.Fhir.Model.Quantity bmiValue)
                        {
                            BMI = bmiValue.Value.ToString();
                        }
                        break;
                    case "生理量測-血壓":
                        string bloodPressureId = sec.Entry[0].Reference.ToString();
                        Observation pb = client.Read<Observation>(bloodPressureId);

                        foreach (var component in pb.Component)
                        {
                            if (component.Code.Coding[0].Code == "8480-6")
                            {
                                if (component.Value is Hl7.Fhir.Model.Quantity systolicValue)
                                {
                                    systolicPressure = systolicValue.Value.ToString();

                                    if (highBloodPressure != "1")
                                    {
                                        highBloodPressure = systolicValue.Value >= 140 ? "1" : "0";  //高血壓: 收縮壓持續處於140 毫米水銀柱(mmHg) 或以上，或舒張壓持續處於90 毫米水銀柱或以上
                                    }
                                }                               
                            }
                            if (component.Code.Coding[0].Code == "8462-4")
                            {
                                if (component.Value is Hl7.Fhir.Model.Quantity diastolicValue)
                                {
                                    diastolicPressure = diastolicValue.Value.ToString();
                                    if (highBloodPressure != "1")
                                    {
                                        highBloodPressure = diastolicValue.Value >= 90 ? "1" : "0";  //高血壓: 收縮壓持續處於140 毫米水銀柱(mmHg) 或以上，或舒張壓持續處於90 毫米水銀柱或以上
                                    }
                                }
                            }
                        }
                        break;
                    case "生理量測-腰圍":
                        string waistId = sec.Entry[0].Reference.ToString();
                        Observation waist = client.Read<Observation>(waistId);

                        if (waist.Value is Hl7.Fhir.Model.Quantity waisyValue)
                        {
                            waistCircumference = waisyValue.Value.ToString();
                        }
                        break;
                    case "檢驗檢查-Protein 尿蛋白":
                        string urineProteinId = sec.Entry[0].Reference.ToString();
                        Observation urine = client.Read<Observation>(urineProteinId);

                        if (urine.Value is Hl7.Fhir.Model.Quantity upValue)
                        {
                            urineProtein = upValue.Value.ToString();
                        }
                        break;
                    case "檢驗檢查-血糖":
                        string bloodSugarId = sec.Entry[0].Reference.ToString();
                        Observation sugar = client.Read<Observation>(bloodSugarId);
                        resultUploadDate = logic.AdToRocEra(sugar.Issued.ToString());

                        if (sugar.Value is Hl7.Fhir.Model.Quantity sugarValue)
                        {
                            fastingBloodSugar = sugarValue.Value.ToString(); 
                            if(hyperglycemiaTemp != "1")
                            {
                                hyperglycemiaTemp = sugarValue.Value >= 130 ? "1" : "0";  //空腹血糖超過130mg/dL
                            }
                        }
                        break;
                    case "檢驗檢查-膽固醇":
                        string cholesterolId = sec.Entry[0].Reference.ToString();
                        Observation chol = client.Read<Observation>(cholesterolId);
                        resultUploadDate = logic.AdToRocEra(chol.Issued.ToString());

                        if (chol.Value is Hl7.Fhir.Model.Quantity cholValue)
                        {
                            cholesterol = cholValue.Value.ToString();
                            if (hyperlipidemia != "1")
                            {
                                hyperlipidemia = cholValue.Value >= 200 ? "1" : "0";  //高血脂: 總膽固醇之理想濃度為 <200mg/dl，三酸甘油酯之理想濃度為<130mg/dl。 當血中之三酸甘油酯和總膽固醇其中之一或兩者皆超過正常值時，即稱為高血脂。
                            }
                        }
                        break;
                    case "檢驗檢查-三酸甘油脂":
                        string triglyceridesId = sec.Entry[0].Reference.ToString();
                        Observation tri = client.Read<Observation>(triglyceridesId);
                        resultUploadDate = logic.AdToRocEra(tri.Issued.ToString());

                        if (tri.Value is Hl7.Fhir.Model.Quantity triValue)
                        {
                            triglycerides = triValue.Value.ToString();
                            if (hyperlipidemia != "1")
                            {
                                hyperlipidemia = triValue.Value >= 130 ? "1" : "0";  //高血脂: 總膽固醇之理想濃度為 <200mg/dl，三酸甘油酯之理想濃度為<130mg/dl。 當血中之三酸甘油酯和總膽固醇其中之一或兩者皆超過正常值時，即稱為高血脂。
                            }
                        }
                        break;
                    case "檢驗檢查-低密度脂蛋白膽固醇計算":
                        string ldlCholesterolId = sec.Entry[0].Reference.ToString();
                        Observation ldlCholesterol = client.Read<Observation>(ldlCholesterolId);
                        resultUploadDate = logic.AdToRocEra(ldlCholesterol.Issued.ToString());

                        if (ldlCholesterol.Value is Hl7.Fhir.Model.Quantity ldlCholesterolValue)
                        {
                            lowDensityLipoproteinCholesterol = ldlCholesterolValue.Value.ToString();
                        }
                        break;
                    case "檢驗檢查-高密度膽固醇":
                        string hdlCholesterolId = sec.Entry[0].Reference.ToString();
                        Observation hdlCholesterol = client.Read<Observation>(hdlCholesterolId);
                        resultUploadDate = logic.AdToRocEra(hdlCholesterol.Issued.ToString());

                        if (hdlCholesterol.Value is Hl7.Fhir.Model.Quantity hdlCholesterolValue)
                        {
                            highDensityLipoproteinCholesterol = hdlCholesterolValue.Value.ToString();
                        }
                        break;
                    case "檢驗檢查-ＧＯＴ":
                        string gotId = sec.Entry[0].Reference.ToString();
                        Observation got = client.Read<Observation>(gotId);
                        resultUploadDate = logic.AdToRocEra(got.Issued.ToString());

                        if (got.Value is Hl7.Fhir.Model.Quantity gotValue)
                        {
                            GOT = gotValue.Value.ToString();
                        }
                        break;
                    case "檢驗檢查-ＧＰＴ":
                        string gptId = sec.Entry[0].Reference.ToString();
                        Observation gpt = client.Read<Observation>(gptId);
                        resultUploadDate = logic.AdToRocEra(gpt.Issued.ToString());

                        if (gpt.Value is Hl7.Fhir.Model.Quantity gptValue)
                        {
                            GPT = gptValue.Value.ToString();
                        }
                        break;
                    case "檢驗檢查-肌酸酐":
                        string creatinineId = sec.Entry[0].Reference.ToString();
                        Observation creat = client.Read<Observation>(creatinineId);
                        resultUploadDate = logic.AdToRocEra(creat.Issued.ToString());

                        if (creat.Value is Hl7.Fhir.Model.Quantity creatinineValue)
                        {
                            creatinine = creatinineValue.Value.ToString();
                        }
                        break;
                    case "檢驗檢查-腎絲球過濾率計算":
                        string egfrId = sec.Entry[0].Reference.ToString();
                        Observation egfr = client.Read<Observation>(egfrId);
                        resultUploadDate = logic.AdToRocEra(egfr.Issued.ToString());

                        if (egfr.Value is Hl7.Fhir.Model.Quantity egfrValue)
                        {
                            glomerularFiltrationRate = egfrValue.Value.ToString();
                        }
                        break;
                    //case "檢驗檢查-Ｂ型肝炎表面抗原":
                    //    hepatitisB = sec.Entry[0].Reference.ToString();
                    //    break;
                    //case "檢驗檢查-Ｃ型肝炎病毒抗體":
                    //    hepatitisC = sec.Entry[0].Reference.ToString();
                    //    break;
                    case "憂鬱檢測：感覺情緒低落沮喪與做事情失去興趣":
                        string depressionId = sec.Entry[0].Reference.ToString();
                        Observation depression = client.Read<Observation>(depressionId);
                        resultUploadDate = logic.AdToRocEra(depression.Issued.ToString());
                        foreach (var component in depression.Component)
                        {
                            if (component.Code.Coding[0].Code == "66446005")  //感覺情緒低落沮喪
                            {
                                if (component.Value is FhirString depressValue)
                                {
                                    lowMood = depressValue.ToString();
                                }
                            }
                            if (component.Code.Coding[0].Code == "713566001")
                            {
                                if (component.Value is FhirString lossInterestValue)
                                {
                                    lossOfInterest = lossInterestValue.ToString();
                                }
                            }
                        }
                        break;
                    case "生活史-吸煙":
                        string smokeId = sec.Entry[0].Reference.ToString();
                        Observation smoke = client.Read<Observation>(smokeId);
                        resultUploadDate = logic.AdToRocEra(smoke.Issued.ToString());
                        if (smoke.Value is FhirString smokeValue)
                        {
                            smoking = smokeValue.ToString();
                        }
                        break;
                    case "生活史-喝酒":
                        string alcoholId = sec.Entry[0].Reference.ToString();
                        Observation alcohol = client.Read<Observation>(alcoholId);
                        resultUploadDate = logic.AdToRocEra(alcohol.Issued.ToString());
                        if (alcohol.Value is FhirString alcoholValue)
                        {
                            alcoholConsumption = alcoholValue.ToString();
                        }
                        break;
                    case "生活史-嚼檳榔":
                        string betelNutId = sec.Entry[0].Reference.ToString();
                        Observation betelNut = client.Read<Observation>(betelNutId);
                        resultUploadDate = logic.AdToRocEra(betelNut.Issued.ToString());
                        if (betelNut.Value is FhirString betelNutValue)
                        {
                            betelNutChewing = betelNutValue.ToString();
                        }
                        break;
                    case "生活史-運動":
                        string exerId = sec.Entry[0].Reference.ToString();
                        Observation exer = client.Read<Observation>(exerId);
                        resultUploadDate = logic.AdToRocEra(exer.Issued.ToString());
                        if (exer.Value is FhirString exerValue)
                        {
                            exercise =exerValue.ToString();
                        }
                        break;
                    case "健康諮詢：戒煙":
                        string quitSmokingConsultationId = sec.Entry[0].Reference.ToString();
                        Observation quitSmokingConsultation = client.Read<Observation>(quitSmokingConsultationId);
                        resultUploadDate = logic.AdToRocEra(quitSmokingConsultation.Issued.ToString());
                        if (quitSmokingConsultation.Value is FhirString quitSmokingConsultationValue)
                        {
                            smokingCessationConsultation = quitSmokingConsultationValue.ToString();
                        }
                        break;
                    case "健康諮詢：節酒":
                        string alcoholReductionConsultationId = sec.Entry[0].Reference.ToString();
                        Observation alcoholReduction = client.Read<Observation>(alcoholReductionConsultationId);
                        resultUploadDate = logic.AdToRocEra(alcoholReduction.Issued.ToString());
                        if (alcoholReduction.Value is FhirString alcoholReductionConsultationValue)
                        {
                            alcoholReductionConsultation = alcoholReductionConsultationValue.ToString();
                        }
                        break;
                    case "健康諮詢：戒檳榔":
                        string betelNutCessationConsultationId = sec.Entry[0].Reference.ToString();
                        Observation betelNutCessation = client.Read<Observation>(betelNutCessationConsultationId);
                        resultUploadDate = logic.AdToRocEra(betelNutCessation.Issued.ToString());
                        if (betelNutCessation.Value is FhirString betelNutCessationConsultationValue)
                        {
                            betelNutCessationConsultation = betelNutCessationConsultationValue.ToString();
                        }
                        break;
                    case "健康諮詢：規律運動":
                        string regularExerciseConsultationId = sec.Entry[0].Reference.ToString();
                        Observation regularExercise = client.Read<Observation>(regularExerciseConsultationId);
                        resultUploadDate = logic.AdToRocEra(regularExercise.Issued.ToString());
                        if (regularExercise.Value is FhirString regularExerciseConsultationValue)
                        {
                            regularExerciseConsultation = regularExerciseConsultationValue.ToString();
                        }
                        break;
                    case "健康諮詢：維持正常體重":
                        string maintainNormalWeightConsultationId = sec.Entry[0].Reference.ToString();
                        Observation maintainNormalWeight = client.Read<Observation>(maintainNormalWeightConsultationId);
                        resultUploadDate = logic.AdToRocEra(maintainNormalWeight.Issued.ToString());
                        if (maintainNormalWeight.Value is FhirString maintainNormalWeightConsultationValue)
                        {
                            maintainNormalWeightConsultation = maintainNormalWeightConsultationValue.ToString();
                        }
                        break;
                    case "健康諮詢：健康飲食":
                        string healthyDietConsultationId = sec.Entry[0].Reference.ToString();
                        Observation healthyDiet = client.Read<Observation>(healthyDietConsultationId);
                        resultUploadDate = logic.AdToRocEra(healthyDiet.Issued.ToString());
                        if (healthyDiet.Value is FhirString healthyDietConsultationValue)
                        {
                            healthyDietConsultation = healthyDietConsultationValue.ToString();
                        }
                        break;
                    case "健康諮詢：事故傷害預":
                        string accidentInjuryPreventionConsultationId = sec.Entry[0].Reference.ToString();
                        Observation accidentInjuryPrevention = client.Read<Observation>(accidentInjuryPreventionConsultationId);
                        if (accidentInjuryPrevention.Value is FhirString accidentInjuryPreventionConsultationValue)
                        {
                            accidentInjuryPreventionConsultation = accidentInjuryPreventionConsultationValue.ToString();
                        }
                        break;
                    case "健康諮詢：口腔保健":
                        string oralHealthCareConsultationId = sec.Entry[0].Reference.ToString();
                        Observation oralHealthCare = client.Read<Observation>(oralHealthCareConsultationId);
                        if (oralHealthCare.Value is FhirString oralHealthCareConsultationValue)
                        {
                            oralHealthCareConsultation = oralHealthCareConsultationValue.ToString();
                        }
                        break;
                    case "檢查過B、C型肝炎":
                        string checkBCTypeHepatitisId= sec.Entry[0].Reference.ToString();
                        Observation checkBC = client.Read<Observation>(checkBCTypeHepatitisId);
                        if (checkBC.Value is FhirString checkBCTypeHepatitisValue)
                        {
                            checkBCTypeHepatitis = checkBCTypeHepatitisValue.ToString();
                        }
                        break;
                    case "血壓檢查結果與建議":
                        string bloodPressureResultAndRecommendationId = sec.Entry[0].Reference.ToString();
                        DiagnosticReport bloodPressureReport = client.Read<DiagnosticReport>(bloodPressureResultAndRecommendationId);
                        bloodPressureResultAndRecommendation = bloodPressureReport.Conclusion;
                        break;
                    case "血糖檢查結果判讀":
                        string bloodSugarResultInterpretationId = sec.Entry[0].Reference.ToString();
                        DiagnosticReport bloodSugarReport = client.Read<DiagnosticReport>(bloodSugarResultInterpretationId);
                        bloodSugarResultInterpretation = bloodSugarReport.Conclusion;
                        break;
                    case "血脂肪檢查結果判讀":
                        string lipidProfileResultInterpretationId = sec.Entry[0].Reference.ToString();
                        DiagnosticReport lipidProfileReport = client.Read<DiagnosticReport>(lipidProfileResultInterpretationId);
                        lipidProfileResultInterpretation = lipidProfileReport.Conclusion;
                        break;
                    case "腎功能檢查結果判讀":
                        string kidneyFunctionResultInterpretationId = sec.Entry[0].Reference.ToString();
                        DiagnosticReport kidneyFunctionReport = client.Read<DiagnosticReport>(kidneyFunctionResultInterpretationId);
                        kidneyFunctionResultInterpretation = kidneyFunctionReport.Conclusion;
                        break;
                    case "肝功能檢查結果判讀":
                        string liverFunctionResultInterpretationId = sec.Entry[0].Reference.ToString();
                        DiagnosticReport liverFunctionReport = client.Read<DiagnosticReport>(liverFunctionResultInterpretationId);
                        liverFunctionResultInterpretation = liverFunctionReport.Conclusion;
                        break;
                    case "代謝症候群檢查結果與建議":
                        string metabolicSyndromeResultAndRecommendationId = sec.Entry[0].Reference.ToString();
                        DiagnosticReport metabolicSyndromeReport = client.Read<DiagnosticReport>(metabolicSyndromeResultAndRecommendationId);
                        metabolicSyndromeResultAndRecommendation = metabolicSyndromeReport.Conclusion;
                        break;
                    //case "B型肝炎檢查結果與建議":
                    //    hepatitisBCheckResult = sec.Entry[0].Reference.ToString();
                    //    break;
                    //case "C型肝炎檢查結果與建議":
                    //    hepatitisCCheckResult = sec.Entry[0].Reference.ToString();
                    //    break;
                    case "憂鬱檢測結果與建議":
                        string depressionDetectionResultAndRecommendationId = sec.Entry[0].Reference.ToString();
                        DiagnosticReport depressionDetectionReport = client.Read<DiagnosticReport>(depressionDetectionResultAndRecommendationId);
                        depressionDetectionResultAndRecommendation = depressionDetectionReport.Conclusion;
                        break;
                }

                threeHigh = (highBloodPressure == "1" || hyperlipidemia == "1" || hyperglycemiaTemp == "1") ? "1" : "0";  //三高: 高血壓、高血糖或高血脂
            }

            addExcel(worksheet, index, name, id, birthDate, gender, phone,firstCheckDate, secondCheckDate, 
                resultUploadDate,checkBCTypeHepatitis, hypertensionHistory,diabetesHistory, hyperlipidemiaHistory, 
                heartDisease,stroke, kidneyDisease, systolicPressure,diastolicPressure, highBloodPressure, threeHigh,
                waistCircumference, BMI, smoking, alcoholConsumption,betelNutChewing, exercise, 
                smokingCessationConsultation,alcoholReductionConsultation, betelNutCessationConsultation,
                regularExerciseConsultation, maintainNormalWeightConsultation,healthyDietConsultation, 
                accidentInjuryPreventionConsultation,oralHealthCareConsultation, hepatitisBSurfaceAntigen,
                hepatitisCAntibody, lowMood, lossOfInterest, urineAcidityValue,urineProtein, urineSugar, 
                urineSedimentMicroscopy, occultBlood, urineRedBloodCells, urineWhiteBloodCells, urineEpithelialCells,
                cast, bacteria, appearance, hyperlipidemia, cholesterol, triglycerides, fastingBloodSugar, 
                creatinine, GOT, GPT,highDensityLipoproteinCholesterol, lowDensityLipoproteinCholesterol,
                glomerularFiltrationRate, liverFunctionResultInterpretation,bloodSugarResultInterpretation, 
                lipidProfileResultInterpretation,kidneyFunctionResultInterpretation, 
                bloodPressureResultAndRecommendation,metabolicSyndromeResultAndRecommendation, 
                depressionDetectionResultAndRecommendation, com_id);

            Console.WriteLine(index + "_" + com_id + "Finish!");

        }

        public static ExcelWorksheet addExcel(ExcelWorksheet worksheet, int index,
            string name, string id, string birthDate, string gender, string phone, 
            string firstCheckDate, string secondCheckDate, string resultUploadDate, 
            string checkBCTypeHepatitis, string hypertensionHistory, 
            string diabetesHistory, string hyperlipidemiaHistory, string heartDisease, 
            string stroke, string kidneyDisease, string systolicPressure, 
            string diastolicPressure, string highBloodPressure, string threeHigh, 
            string waistCircumference, string BMI, string smoking, string alcoholConsumption, 
            string betelNutChewing, string exercise, string smokingCessationConsultation, 
            string alcoholReductionConsultation, string betelNutCessationConsultation, 
            string regularExerciseConsultation, string maintainNormalWeightConsultation, 
            string healthyDietConsultation, string accidentInjuryPreventionConsultation, 
            string oralHealthCareConsultation, string hepatitisBSurfaceAntigen, 
            string hepatitisCAntibody, string lowMood, string lossOfInterest, string urineAcidityValue, 
            string urineProtein, string urineSugar, string urineSedimentMicroscopy, string occultBlood, 
            string urineRedBloodCells, string urineWhiteBloodCells, string urineEpithelialCells, 
            string cast, string bacteria, string appearance, string hyperlipidemia, string cholesterol, 
            string triglycerides, string fastingBloodSugar, string creatinine, string GOT, string GPT, 
            string highDensityLipoproteinCholesterol, string lowDensityLipoproteinCholesterol, 
            string glomerularFiltrationRate, string liverFunctionResultInterpretation, 
            string bloodSugarResultInterpretation, string lipidProfileResultInterpretation, 
            string kidneyFunctionResultInterpretation, string bloodPressureResultAndRecommendation, 
            string metabolicSyndromeResultAndRecommendation, string depressionDetectionResultAndRecommendation, 
            string com_id)
        {
            worksheet.Cells[index, 1].Value = name;
            worksheet.Cells[index, 2].Value = id;
            worksheet.Cells[index, 3].Value = birthDate;
            worksheet.Cells[index, 4].Value = gender;
            worksheet.Cells[index, 5].Value = phone;
            worksheet.Cells[index, 6].Value = firstCheckDate;
            worksheet.Cells[index, 7].Value = secondCheckDate;
            worksheet.Cells[index, 8].Value = resultUploadDate;
            worksheet.Cells[index, 9].Value = checkBCTypeHepatitis;
            worksheet.Cells[index, 10].Value = hypertensionHistory;
            worksheet.Cells[index, 11].Value = diabetesHistory;
            worksheet.Cells[index, 12].Value = hyperlipidemiaHistory;
            worksheet.Cells[index, 13].Value = heartDisease;
            worksheet.Cells[index, 14].Value = stroke;
            worksheet.Cells[index, 15].Value = kidneyDisease;
            worksheet.Cells[index, 16].Value = systolicPressure;
            worksheet.Cells[index, 17].Value = diastolicPressure;
            worksheet.Cells[index, 18].Value = highBloodPressure;
            worksheet.Cells[index, 19].Value = threeHigh;
            worksheet.Cells[index, 20].Value = waistCircumference;
            worksheet.Cells[index, 21].Value = BMI;
            worksheet.Cells[index, 22].Value = smoking;
            worksheet.Cells[index, 23].Value = alcoholConsumption;
            worksheet.Cells[index, 24].Value = betelNutChewing;
            worksheet.Cells[index, 25].Value = exercise;
            worksheet.Cells[index, 26].Value = smokingCessationConsultation;
            worksheet.Cells[index, 27].Value = alcoholReductionConsultation;
            worksheet.Cells[index, 28].Value = betelNutCessationConsultation;
            worksheet.Cells[index, 29].Value = regularExerciseConsultation;
            worksheet.Cells[index, 30].Value = maintainNormalWeightConsultation;
            worksheet.Cells[index, 31].Value = healthyDietConsultation;
            worksheet.Cells[index, 32].Value = accidentInjuryPreventionConsultation;
            worksheet.Cells[index, 33].Value = oralHealthCareConsultation;
            worksheet.Cells[index, 34].Value = hepatitisBSurfaceAntigen;
            worksheet.Cells[index, 35].Value = hepatitisCAntibody;
            worksheet.Cells[index, 36].Value = lowMood;
            worksheet.Cells[index, 37].Value = lossOfInterest;
            worksheet.Cells[index, 38].Value = urineAcidityValue;
            worksheet.Cells[index, 39].Value = urineProtein;
            worksheet.Cells[index, 40].Value = urineSugar;
            worksheet.Cells[index, 41].Value = urineSedimentMicroscopy;
            worksheet.Cells[index, 42].Value = occultBlood;
            worksheet.Cells[index, 43].Value = urineRedBloodCells;
            worksheet.Cells[index, 44].Value = urineWhiteBloodCells;
            worksheet.Cells[index, 45].Value = urineEpithelialCells;
            worksheet.Cells[index, 46].Value = cast;
            worksheet.Cells[index, 47].Value = bacteria;
            worksheet.Cells[index, 48].Value = appearance;
            worksheet.Cells[index, 49].Value = hyperlipidemia;
            worksheet.Cells[index, 50].Value = cholesterol;
            worksheet.Cells[index, 51].Value = triglycerides;
            worksheet.Cells[index, 52].Value = fastingBloodSugar;
            worksheet.Cells[index, 53].Value = creatinine;
            worksheet.Cells[index, 54].Value = GOT;
            worksheet.Cells[index, 55].Value = GPT;
            worksheet.Cells[index, 56].Value = highDensityLipoproteinCholesterol;
            worksheet.Cells[index, 57].Value = lowDensityLipoproteinCholesterol;
            worksheet.Cells[index, 58].Value = glomerularFiltrationRate;
            worksheet.Cells[index, 59].Value = liverFunctionResultInterpretation;
            worksheet.Cells[index, 60].Value = bloodSugarResultInterpretation;
            worksheet.Cells[index, 61].Value = lipidProfileResultInterpretation;
            worksheet.Cells[index, 62].Value = kidneyFunctionResultInterpretation;
            worksheet.Cells[index, 63].Value = bloodPressureResultAndRecommendation;
            worksheet.Cells[index, 64].Value = metabolicSyndromeResultAndRecommendation;
            worksheet.Cells[index, 65].Value = depressionDetectionResultAndRecommendation;
            worksheet.Cells[index, 66].Value = com_id;

            return worksheet;
        }

        
    

    }
}