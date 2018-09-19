using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Xml;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace RepentYourSins2
{
    [FormAttribute("RepentYourSins2.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button0.PressedAfter +=
                new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button2.PressedAfter +=
                new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button1.PressedAfter +=
                new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.Button3.PressedAfter +=
                new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button3_PressedAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button4.PressedAfter +=
                new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button4_PressedAfter);
            this.Button5 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.Button5.PressedAfter +=
                new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button5_PressedAfter);
            this.Button6 = ((SAPbouiCOM.Button)(this.GetItem("Item_6").Specific));
            this.Button6.PressedAfter +=
                new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button6_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;

        private void OnCustomInitialize()
        {

        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //    Program.XCompany.StartTransaction();

            Program.recSet.DoQuery(SAPApi.DIManager.QueryHanaTransalte(@"
                SELECT distinct OJDT.TransId from JDT1
                left join OJDT on OJDT.TransId = JDT1.TransId where (((JDT1.Ref2 is not Null and JDT1.Ref2 != '') or (JDT1.Ref3line != '' and JDT1.Ref3line is not Null)  or (OJDT.Ref2 is not null and OJDT.Ref2 != '')) 
							and  
				  (JDT1.Ref2 not like '%[0123456789]%' and  JDT1.Ref2 is not null and JDT1.Ref2 != '')
				or (JDT1.Ref3Line not like '%[0123456789]%' and  JDT1.Ref3Line is not null and JDT1.Ref3Line != ''))
				 
               "));


            JournalEntries journalEntry =
                (JournalEntries)Program.XCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
            int count = Program.recSet.RecordCount;
            int countmatch = 0;
            while (!Program.recSet.EoF)
            {
                int transid = int.Parse(Program.recSet.Fields.Item("TransId").Value.ToString());
                journalEntry.GetByKey(transid);
                countmatch++;

                string ref3RowValue = string.Empty;
                string ref2Rowvalue = string.Empty;


                bool ref2Row = false;
                bool refref3row = false;



                for (int i = 0; i < journalEntry.Lines.Count; i++)
                {
                    journalEntry.Lines.SetCurrentLine(i);
                    ref2Rowvalue = journalEntry.Lines.Reference2;
                    if (!string.IsNullOrWhiteSpace(ref2Rowvalue) && !ref2Rowvalue.All(char.IsDigit))
                    {
                        ref2Row = true;
                    }
                    ref3RowValue = journalEntry.Lines.AdditionalReference;
                    if (!string.IsNullOrWhiteSpace(ref3RowValue) && !ref3RowValue.All(char.IsDigit))
                    {
                        refref3row = true;
                    }
                    if (ref2Row)
                    {
                        try
                        {
                            if (journalEntry.Lines.ControlAccount != "1430" &&
                                journalEntry.Lines.ControlAccount != "3130")
                            {
                                continue;
                            }
                            journalEntry.Lines.CostingCode3 = "";
                        }
                        catch (Exception e)
                        {
                            Program.recSet.MoveNext();
                        }
                    }
                    else if (refref3row)
                    {
                        try
                        {
                            if (journalEntry.Lines.ContraAccount == "1430" ||
                                journalEntry.Lines.ContraAccount == "3130")
                            {
                                journalEntry.Lines.SetCurrentLine(1);
                                journalEntry.Lines.CostingCode3 = "";
                                refref3row = false;
                            }
                        }
                        catch (Exception e)
                        {
                            Program.recSet.MoveNext();
                        }
                    }
                }

                var x = journalEntry.Update();
                var y = Program.XCompany.GetLastErrorDescription();
                if (x != 0)
                {
                    //  Program.XCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                Program.recSet.MoveNext();
            }


            //   Program.XCompany.EndTransaction(BoWfTransOpt.wf_Commit);
        }

        private SAPbouiCOM.Button Button2;

        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            List<string> Emplyees = new List<string>();
            Program.recSet.DoQuery(SAPApi.DIManager.QueryHanaTransalte(
                @"select distinct CONCAT(OHEM.firstName, ' ' ,OHEM.lastName) as [FirstLast], CONCAT(OHEM.lastName, ' ' ,OHEM.firstName) as [LastFirst] from OHEM"));
            while (!Program.recSet.EoF)
            {
                Emplyees.Add(Program.recSet.Fields.Item("FirstLast").Value.ToString());
                Emplyees.Add(Program.recSet.Fields.Item("LastFirst").Value.ToString());
                Program.recSet.MoveNext();
            }
            Emplyees = Emplyees.Distinct().ToList();
            string x = "ბაკურიძე ინგა1";
            string x2 = "აბაზაძე მიხეილ";

            var x1 = Emplyees.IndexOf(x);
            var x3 = Emplyees.Equals(x2);
        }

        private SAPbouiCOM.Button Button1;

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Program.XCompany.StartTransaction();
            Recordset recSet =
                (Recordset)Program.XCompany.GetBusinessObject(BoObjectTypes
                    .BoRecordset);
            recSet.DoQuery(SAPApi.DIManager.QueryHanaTransalte("select empID, concat(lastName, ' ', firstName) as [Name] from OHEM"));

            while (!recSet.EoF)
            {


                SAPbobsCOM.CompanyService oCmpSrv;
                SAPbobsCOM.IProfitCentersService oProfitCentersService;
                SAPbobsCOM.IProfitCenter oProfitCenter;
                try
                {
                    oCmpSrv = Program.XCompany.GetCompanyService();
                    oProfitCentersService =
                        (IProfitCentersService)oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes
                            .ProfitCentersService);
                    oProfitCenter =
                        (IProfitCenter)oProfitCentersService.GetDataInterface(SAPbobsCOM
                            .ProfitCentersServiceDataInterfaces.pcsProfitCenter);

                    oProfitCenter.CenterCode = $"{recSet.Fields.Item("empID").Value.ToString()}";
                    oProfitCenter.CenterName = $"{recSet.Fields.Item("Name").Value.ToString()}";
                    oProfitCenter.EffectiveTo = new DateTime(2099, 1, 1);
                    oProfitCenter.Effectivefrom = new DateTime(2017, 1, 1);
                    oProfitCenter.InWhichDimension = 3;

                    oProfitCentersService.AddProfitCenter((SAPbobsCOM.ProfitCenter)oProfitCenter);

                }
                catch (Exception ex)
                {
                    Console.WriteLine("exepton is :" + ex.Message);
                    Program.XCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                recSet.MoveNext();
            }
            Program.XCompany.EndTransaction(BoWfTransOpt.wf_Commit);
        }

        private SAPbouiCOM.Button Button3;

        private void Button3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Program.XCompany.StartTransaction();
            Recordset recSet =
                (Recordset)Program.XCompany.GetBusinessObject(BoObjectTypes
                    .BoRecordset);
            recSet.DoQuery(SAPApi.DIManager.QueryHanaTransalte(
                "select OHEM.empID, OPRC.PrcCode from OHEM inner join OPRC on CONVERT(nvarchar, OPRC.PrcCode) = CONVERT(nvarchar, OHEM.empID)"));

            while (!recSet.EoF)
            {
                SAPbobsCOM.EmployeesInfo employee =
                    (SAPbobsCOM.EmployeesInfo)Program.XCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes
                        .oEmployeesInfo);
                employee.GetByKey(int.Parse(recSet.Fields.Item("empID").Value.ToString()));
                employee.CostCenterCode = recSet.Fields.Item("empID").Value.ToString();
                var x = employee.Update();
                if (x != 0)
                {
                    Program.XCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                recSet.MoveNext();
            }
            Program.XCompany.EndTransaction(BoWfTransOpt.wf_Commit);
        }

        private SAPbouiCOM.Button Button4;

        private void Button4_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            JournalEntries journalEntry =
                (JournalEntries)Program.XCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
            journalEntry.GetByKey(14323);

            string ref3RowValue = string.Empty;
            string ref2Rowvalue = string.Empty;


            bool ref2Row = false;
            bool refref3row = false;


            var countx = 0;
            for (int i = 0; i < journalEntry.Lines.Count; i++)
            {
                journalEntry.Lines.SetCurrentLine(i);
                ref2Rowvalue = journalEntry.Lines.Reference2;
                if (!string.IsNullOrWhiteSpace(ref2Rowvalue) && !ref2Rowvalue.All(char.IsDigit))
                {
                    ref2Row = true;
                    countx++;
                }
                ref3RowValue = journalEntry.Lines.AdditionalReference;
                if (!string.IsNullOrWhiteSpace(ref3RowValue) && !ref3RowValue.All(char.IsDigit))
                {
                    refref3row = true;
                }
                if (ref2Row)
                {
                    try
                    {
                        var x1 = journalEntry.Lines.ControlAccount != "1430";
                        if (journalEntry.Lines.ControlAccount != "1430" && journalEntry.Lines.ControlAccount != "3130")
                        {
                            continue;
                        }
                        journalEntry.Lines.CostingCode3 = Program.emps[ref2Rowvalue];


                    }
                    catch (Exception e)
                    {
                        Program.recSet.MoveNext();
                    }
                }
                else if (refref3row)
                {
                    try
                    {
                        if (journalEntry.Lines.ContraAccount == "1430" || journalEntry.Lines.ContraAccount == "3130")
                        {
                            journalEntry.Lines.SetCurrentLine(1);
                            journalEntry.Lines.CostingCode3 = Program.emps[ref3RowValue];
                            refref3row = false;
                        }
                    }
                    catch (Exception e)
                    {
                        Program.recSet.MoveNext();
                    }
                }
            }

            var x = journalEntry.Update();
            var y = Program.XCompany.GetLastErrorDescription();
        }

        private SAPbouiCOM.Button Button5;
        private SAPbouiCOM.Button Button6;

        private void Button5_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Program.recSet.DoQuery(SAPApi.DIManager.QueryHanaTransalte(@"  SELECT   distinct JDT1.TransId  from JDT1
                left join OJDT on OJDT.TransId = JDT1.TransId 
				where (JDT1.Ref2 != '' and JDT1.Ref2 is not Null) and JDT1.Ref2 not like '%[0123456789]%'		 
               "));
            JournalEntries journalEntry =
                (JournalEntries)Program.XCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);

            var count = 0;
            while (!Program.recSet.EoF)
            {
                int transid = int.Parse(Program.recSet.Fields.Item("TransId").Value.ToString());
                journalEntry.GetByKey(transid);
                string ref2Rowvalue = string.Empty;


                for (int i = 0; i < journalEntry.Lines.Count; i++)
                {
                    journalEntry.Lines.SetCurrentLine(i);
                    ref2Rowvalue = journalEntry.Lines.Reference2;


                    if (journalEntry.Lines.ControlAccount != "1430" && journalEntry.Lines.ControlAccount != "3130")
                    {
                        continue;
                    }

                    journalEntry.Lines.CostingCode3 = Program.emps[ref2Rowvalue];

                }
                var x = journalEntry.Update();
                var y = Program.XCompany.GetLastErrorDescription();
                if (x != 0)
                {
                    //  Program.XCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                Program.recSet.MoveNext();
                count++;
                Debug.WriteLine(count);
            }

        }

        private void Button6_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            var count = 0;

            Program.recSet.DoQuery(SAPApi.DIManager.QueryHanaTransalte(@"
                  SELECT distinct JDT1.TransId  from JDT1
                left join OJDT on OJDT.TransId = JDT1.TransId 
				where (JDT1.Ref3line != '' and JDT1.Ref3line is not Null) and JDT1.Ref3line not like '%[0123456789]%'
               "));

            JournalEntries journalEntry =
                (JournalEntries)Program.XCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);

            while (!Program.recSet.EoF)
            {
                int transid = int.Parse(Program.recSet.Fields.Item("TransId").Value.ToString());
                journalEntry.GetByKey(transid);

                string ref3RowValue = string.Empty;


                for (int i = 0; i < journalEntry.Lines.Count; i++)
                {
                    journalEntry.Lines.SetCurrentLine(i);
                    ref3RowValue = journalEntry.Lines.AdditionalReference;
                    if (journalEntry.Lines.ContraAccount == "1430" || journalEntry.Lines.ContraAccount == "3130")
                    {
                        journalEntry.Lines.SetCurrentLine(1);
                        journalEntry.Lines.CostingCode3 = Program.emps[ref3RowValue];
                    }
                    var x = journalEntry.Update();
                    var y = Program.XCompany.GetLastErrorDescription();
                    if (x != 0)
                    {
                        //  Program.XCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                    }
                    count++;
                    Debug.WriteLine(count);
                }
                Program.recSet.MoveNext();
            }
        }
    }
}