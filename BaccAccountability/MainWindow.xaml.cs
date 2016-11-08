using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using Forms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace BaccAccountability
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void run(object sender, RoutedEventArgs e)
        {
            SqlConnection conn = new SqlConnection("Server=vulcan;database=Adhoc;Trusted_Connection=yes");
            List<String> inStateStudents = new List<string>();
            List<String> outOfStateStudents = new List<string>();
            int inStateHours = 0;
            int outOfStateHours = 0;
            Dictionary<String, int> inStateProgramHours = new Dictionary<string, int>();
            Dictionary<String, int> outOfStateProgramHours = new Dictionary<string, int>();
            Dictionary<String, List<String>> inStateProgramEnrollments = new Dictionary<string, List<string>>();
            Dictionary<String, List<String>> outOfStateProgramEnrollments = new Dictionary<string, List<string>>();

            try
            {
                conn.Open();
            }
            catch (Exception)
            {
                
                throw;
            }

            SqlCommand comm = new SqlCommand("SELECT                                                                                        "
	                                         +"       r6.DE1021,                                                                            "
	                                         +"       r1.[TERM-ID],                                                                         "
	                                         +"       aa1a.CIP,                                                                             "
	                                         +"       SUM(CAST(LEFT(r6.DE3012, 4) AS INT)) AS Hours,                                        "
	                                         +"       r1.[DE1004-RESIDENCE-FEE],                                                            "
	                                         +"       prog.PGM_CD,                                                                          "
                                             +"       ROW_NUMBER() OVER (PARTITION BY r6.DE1021, r1.[TERM-ID] ORDER BY prog.PGM_CD) RN      "
                                             +"   FROM                                                                                      "
	                                         +"       StateSubmission.sdb.RecordType6 r6                                                    "
	                                         +"       LEFT JOIN StateSubmission.sdb.RecordType1 r1 ON r1.[STUDENT-ID] = r6.DE1021           "
											 +"	                                                  AND r1.[TERM-ID] = r6.DE1028              "
	                                         +"       LEFT JOIN Adhoc.dbo.AA1AENRL aa1a ON aa1a.[Student Identification] = r6.DE1021        "
									         +"                                         AND aa1a.Term = LEFT(r6.DE1028, 1)                  "
	                                         +"       LEFT JOIN MIS.dbo.ST_STATE_FED_CIP_CODES_393 xwalk ON xwalk.FEDERAL_CIP_CD = aa1a.CIP "
	                                         +"       LEFT JOIN MIS.dbo.ST_PROGRAMS_A_136 prog ON prog.CIP_CD = xwalk.STATE_CIP_CD          "
                                             +"   WHERE                                                                                     "
	                                         +"       SUBSTRING(r6.DE3008, 4, 1) IN ('3','4')                                               "
	                                         +"       AND r6.submissionType = 'E'                                                           "
	                                         +"       AND r1.submissionType = 'E'                                                           "
	                                         +"       AND r6.DE1028 IN ('115','215','316')                                                  "
	                                         +"       AND r1.[TERM-ID] IN ('115','215','316')	                                            "
	                                         +"       AND (aa1a.[Program of Study Level] = 'C' OR aa1a.[Program of Study Level] IS NULL)    "
	                                         +"       AND (prog.EFF_TRM_D <> '' OR prog.EFF_TRM_D IS NULL)                                  "
	                                         +"       AND (prog.END_TRM = '' OR prog.END_TRM IS NULL)                                       "
	                                         +"       AND (LEFT(prog.AWD_TY, 1) = 'B' OR prog.AWD_TY IS NULL)                               "
                                             +"   GROUP BY                                                                                  "
	                                         +"       r6.DE1021                                                                             "
	                                         +"       ,r1.[DE1004-RESIDENCE-FEE]                                                            "
	                                         +"       ,r1.[TERM-ID]                                                                         "
	                                         +"       ,aa1a.CIP                                                                             "
	                                         +"       ,prog.PGM_CD                                                                          "
                                             +"   ORDER BY                                                                                  "
	                                         +"       r6.DE1021                                                                             "
	                                         +"       ,r1.[TERM-ID] DESC                                                                    "
	                                         +"       ,aa1a.CIP", conn);

            SqlDataReader reader = comm.ExecuteReader();

            String curStudent;
            String curProgram;
            String curResidency;
            String curTerm;

            String effectiveStudent = "";
            String effectiveProgram = "";
            String effectiveSecondaryProgram = "";
            String effectiveTerm = "";
            String effectiveResidency = "F";

            int effectiveProgramHours = 0;
            int totalHours = 0;

            while (reader.Read())
            {
                curStudent = reader["DE1021"].ToString();
                curProgram = reader["PGM_CD"].ToString();
                curTerm = reader["TERM-ID"].ToString();
                curResidency = reader["DE1004-RESIDENCE-FEE"].ToString();
                int rowNumber = int.Parse(reader["RN"].ToString());
                int hours = int.Parse(reader["Hours"].ToString());

                if (curProgram != "" && !inStateProgramEnrollments.ContainsKey(curProgram))
                {
                    inStateProgramEnrollments.Add(curProgram, new List<string>());
                    outOfStateProgramEnrollments.Add(curProgram, new List<string>());
                }
                                
                if (curStudent != effectiveStudent || (curProgram != "" && curProgram != effectiveProgram && rowNumber != 2))
                {
                    if (curStudent != effectiveStudent)
                    {
                        if (effectiveResidency == "F")
                        {
                            inStateStudents.Add(effectiveStudent);
                        }
                        else
                        {
                            outOfStateStudents.Add(effectiveStudent);
                        }
                    }

                    effectiveStudent = curStudent;
                                        
                    if (effectiveProgramHours != 0)
                    {
                        totalHours += effectiveProgramHours;
                                                
                        Dictionary<String, int> correctProgramHourDictionaryForResidency = (effectiveResidency == "F" ? inStateProgramHours : outOfStateProgramHours);
                        Dictionary<String, List<String>> correctEnrollmentDictionaryForResidency = (effectiveResidency == "F" ? inStateProgramEnrollments : outOfStateProgramEnrollments);

                        if (!correctProgramHourDictionaryForResidency.ContainsKey(effectiveProgram))
                        {
                            correctProgramHourDictionaryForResidency.Add(effectiveProgram, effectiveProgramHours);
                        }
                        else
                        {
                            correctProgramHourDictionaryForResidency[effectiveProgram] += effectiveProgramHours;
                        }

                        if (!correctEnrollmentDictionaryForResidency[effectiveProgram].Contains(curStudent))
                        {
                            correctEnrollmentDictionaryForResidency[effectiveProgram].Add(curStudent);
                        }

                        if (effectiveSecondaryProgram != "")
                        {
                            if (!correctProgramHourDictionaryForResidency.ContainsKey(effectiveSecondaryProgram))
                            {
                                correctProgramHourDictionaryForResidency.Add(effectiveSecondaryProgram, effectiveProgramHours);
                            }
                            else
                            {
                                correctProgramHourDictionaryForResidency[effectiveSecondaryProgram] += effectiveProgramHours;
                            }
                        }

                        if (effectiveResidency == "F")
                        {
                            inStateHours += effectiveProgramHours;
                        }
                        else
                        {
                            outOfStateHours += effectiveProgramHours;
                        }

                        effectiveProgramHours = 0;
                        effectiveResidency = "F";
                    }

                    if (rowNumber == 2)
                    {
                        effectiveSecondaryProgram = curProgram;
                    }
                    else if (curProgram != "")
                    {
                        effectiveSecondaryProgram = "";
                    }

                    effectiveTerm = curTerm;

                    if (curProgram != "" && effectiveSecondaryProgram == "")
                    {
                        effectiveProgram = curProgram;
                    }
                }
                if (curResidency != "F")
                {
                    effectiveResidency = "N";
                }
                if (effectiveProgram != "")
                {
                    effectiveProgramHours += hours;
                }               
            }

            reader.Close();
            conn.Close();

            int totalInStateHours = 0;
            int totalOutOfStateHours = 0;

            foreach (String program in inStateProgramHours.Keys)
            {
                totalInStateHours += inStateProgramHours[program];
            }

            foreach (String program in outOfStateProgramHours.Keys)
            {
                totalOutOfStateHours += outOfStateProgramHours[program];
            }

            int inStateStudentHeadCount = inStateStudents.Count;
            int outOfStateStudentHeadCount = outOfStateStudents.Count;

            buildFile(inStateStudentHeadCount, outOfStateStudentHeadCount, totalInStateHours, totalOutOfStateHours, inStateProgramHours, outOfStateProgramHours,
                inStateProgramEnrollments, outOfStateProgramEnrollments);
        }

        private void select(object sender, RoutedEventArgs e)
        {
            Forms.FolderBrowserDialog fold = new Forms.FolderBrowserDialog();

            if (fold.ShowDialog() == Forms.DialogResult.OK)
            {
                outputDirectory.Text = fold.SelectedPath.ToString();
            }
  
        }

        public void buildFile(int totalInStateHeadCount, int totalOutOfStateHeadCount, int totalInStateProgramHours, int totalOutOfStateProgramHours,
            Dictionary<String, int> inStateProgramHours, Dictionary<String, int> outOfStateProgramHours,
            Dictionary<String, List<String>> inStateProgramEnrollments, Dictionary<String, List<String>> outOfStateProgramEnrollments)
        {
            Excel.Application xlApp = new Excel.Application();
            xlApp.Workbooks.Add();

            xlApp.Range["A1"].Value = "Program";
            xlApp.Range["B1"].Value = "Total In-State HeadCount";
            xlApp.Range["C1"].Value = "Total In-State Enrollment Hours";
            xlApp.Range["D1"].Value = "Total Out-of-State HeadCount";
            xlApp.Range["E1"].Value = "Total Out-of-State Enrollment Hours";

            int rowNum = 3;

            String[] programs = new String[inStateProgramEnrollments.Keys.Count];
            inStateProgramEnrollments.Keys.CopyTo(programs, 0);

            xlApp.Range["A2"].Value = "All Programs";
            xlApp.Range["B2"].Value = totalInStateHeadCount;
            xlApp.Range["C2"].Value = totalInStateProgramHours;
            xlApp.Range["D2"].Value = totalInStateHeadCount;
            xlApp.Range["E2"].Value = totalOutOfStateHeadCount;
            
            foreach (String program in programs)
            {
                xlApp.Range["A" + rowNum].Value = program;
                xlApp.Range["B" + rowNum].Value = inStateProgramEnrollments.ContainsKey(program) ? inStateProgramEnrollments[program].Count : 0;
                xlApp.Range["C" + rowNum].Value = inStateProgramHours.ContainsKey(program) ? inStateProgramHours[program] : 0;
                xlApp.Range["D" + rowNum].Value = outOfStateProgramEnrollments.ContainsKey(program) ? outOfStateProgramEnrollments[program].Count : 0;
                xlApp.Range["E" + rowNum].Value = outOfStateProgramHours.ContainsKey(program) ? outOfStateProgramHours[program] : 0;

                rowNum++;
            }

            xlApp.ActiveWorkbook.SaveAs(outputDirectory.Text + "\\Enrollment Hours By Bacc Program And State Residency.xlsx",
                Type.Missing, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            xlApp.ActiveWorkbook.Close();
        }                   
    }
}
