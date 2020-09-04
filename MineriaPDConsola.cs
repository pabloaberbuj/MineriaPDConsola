using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using VMS.CA.Scripting;
using VMS.DV.PD.Scripting;
using System.Diagnostics;
using System.Globalization;

namespace MineriaPDConsola
{

    static class Program
    {

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                using (Application app = Application.CreateApplication())
                {
                    Execute(app);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
            }
        }

        static string Patologia(PlanSetup planSetup)
        {
            if (planSetup.UniqueFractionation.NumberOfFractions == 39)
            {
                return "Prostata";
            }
            else if (planSetup.UniqueFractionation.NumberOfFractions == 33)
            {
                return "CyC";
            }
            else if (planSetup.StructureSet != null)
            {
                if (planSetup.StructureSet.Structures.Any(s => s.Id.Contains("Bladder-PTV")))
                {
                    return "Prostata";
                }
                if (planSetup.StructureSet.Structures.Any(s => s.Id.Contains("CTV_Vasos")))
                {
                    return "Páncreas";
                }
                else if (planSetup.StructureSet.Structures.Any(s => s.Id.Contains("Lung") || s.Id.Contains("Heart")))
                {
                    return "Tórax";
                }
                else if (planSetup.StructureSet.Structures.Any(s => s.Id.Contains("Kidney") || s.Id.Contains("Liver")))
                {
                    return "Abdomen";
                }
                else if (planSetup.StructureSet.Structures.Any(s => s.Id.Contains("Bowel") || s.Id.Contains("Rectum")))
                {
                    return "Pelvis";
                }
                else
                {
                    return "Indefinido";
                }
            }
            else
            {
                return "Indefinido";
            }

        }

        static void Execute(Application app)
        {
            VMS.DV.PD.UI.Base.VTransientImageDataMgr.CreateInstance(true);
            Stopwatch sw = new Stopwatch();
            sw.Start();
            List<EvaluationTestDesc> evaluationTestDescs = new List<EvaluationTestDesc>();
            EvaluationTestDesc evaluationTestDesc = new EvaluationTestDesc(EvaluationTestKind.GammaAreaLessThanOne, double.NaN, 0.95, true);
            evaluationTestDescs.Add(evaluationTestDesc);
            PDTemplate pDTemplate = new PDTemplate(false, false, false, false, AnalysisMode.CU, NormalizationMethod.MaxEachDose, false, 0.1, ROIType.Field, 10, 0.04, 4, false, evaluationTestDescs);
            List<string> salida = new List<string>();
            List<string> log = new List<string>();
            string encabezado = ("ID;Nombre;Plan;Patología;Campo;Fecha;Fecha Inicio;Resultado");
            File.WriteAllText(@"\\Ariadb-cdt\va_transfer\03_PABLO\Registro.txt", encabezado);

            Console.WriteLine("Iniciando...");
            int i = 1;
            List<string> excluidas = new List<string> { "CAP-0001", "1-01861-0", "Prueba", "19-24419-0" };
            string patologia = "";

            foreach (PatientSummary PacienteSum in app.PatientSummaries)
            {
                if (!excluidas.Contains(PacienteSum.Id))
                {
                    int j = 0;
                    Console.WriteLine(i.ToString() + ". Paciente: " + PacienteSum.Id);
                    Patient Paciente = app.OpenPatient(PacienteSum);
                    if (Paciente.PDPlanSetups.Count() > 0)
                    {
                        foreach (PDPlanSetup pDPlanSetup in Paciente.PDPlanSetups)
                        {
                            if (!pDPlanSetup.PlanSetup.Course.Id.Contains("QA") && pDPlanSetup.Beams.Count > 0)
                            {
                                patologia = Patologia(pDPlanSetup.PlanSetup);
                                foreach (PDBeam campo in pDPlanSetup.Beams)
                                {
                                    if (campo.PortalDoseImages.Count > 1)
                                    {
                                        foreach (PortalDoseImage imagen in campo.PortalDoseImages.Skip(1)) //saltea el primero
                                        {
                                            try
                                            {
                                                //double diasDesdeInicio = Math.Ceiling((imagen.Session.SessionDate - campo.PortalDoseImages.First().Session.SessionDate).TotalDays);
                                                PDAnalysis pDAnalysis = imagen.CreateTransientAnalysis(pDTemplate, campo.PortalDoseImages.First());
                                                string aux = Paciente.Id + ";" + Paciente.Name + ";" + pDPlanSetup.Id + ";" + patologia + ";" + campo.Id + ";" + imagen.Session.SessionDate.ToShortDateString() + ";" + campo.PortalDoseImages.First().Session.SessionDate.ToShortDateString() + ";" + Math.Round(pDAnalysis.EvaluationTests.First().TestValue, 4).ToString();
                                                salida.Add(aux);
                                                j++;
                                            }
                                            catch (Exception e)
                                            {
                                                Console.WriteLine("Error: ");
                                                log.Add(Paciente.Id + " " + imagen.Id + " " + campo.PortalDoseImages.First() + e.ToString());
                                            }
                                        }
                                    }

                                }

                            }

                        }


                    }
                    Console.WriteLine("Se analizaron " + j.ToString() + " imágenes");
                    app.ClosePatient();
                    i++;
                    if (i%40==0)
                    {
                        Console.WriteLine("Se escribieron " + (salida.Count - 1).ToString() + " registros");
                        Console.WriteLine("Demoró " + sw.Elapsed.ToString());
                        File.AppendAllLines(@"\\Ariadb-cdt\va_transfer\03_PABLO\Registro.txt", salida.ToArray());
                        File.WriteAllLines(@"\\Ariadb-cdt\va_transfer\03_PABLO\log.txt", log.ToArray());
                        salida.Clear();
                    }
                }

            }
            sw.Stop();
            Console.WriteLine("Se escribieron " + (salida.Count - 1).ToString() + " registros");
            Console.WriteLine("Demoró " + sw.Elapsed.ToString());
            salida.Add(sw.Elapsed.ToString());
            File.AppendAllLines(@"\\Ariadb-cdt\va_transfer\03_PABLO\Registro.txt", salida.ToArray());
            File.WriteAllLines(@"\\Ariadb-cdt\va_transfer\03_PABLO\log.txt", log.ToArray());
            Console.Read();

        }

    }

}
