using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ResourceScheduler.Classes
{
   public class ExcelBuilder
    {
        private static List<XLColor> colors = new List<XLColor>() { XLColor.DollarBill,
      XLColor.Iceberg,
      XLColor.Moccasin,
      XLColor.Bittersweet,
      XLColor.Jonquil,
      XLColor.NaplesYellow,
      XLColor.Sandstorm,
      XLColor.Alizarin,
    };

        public static void Start(List<Operation> operacije)
        {

            try
            {
                string filename = "test_" + Guid.NewGuid().ToString() + ".xlsx";
                XLWorkbook xLWorkbook = new XLWorkbook();
                IXLWorksheet s1 = xLWorkbook.Worksheets.Add("Operacije");
                IXLWorksheet s2 = xLWorkbook.Worksheets.Add("Radni nalozi");
                IXLWorksheet s3; // = xLWorkbook.Worksheets.Add("Kapaciteti");



                //SqlDataReader dr = null;

                //List<SqlParameter> parameters = new List<SqlParameter>();
                //{
                //    SqlParameter p = new SqlParameter("@p_OrganizacijaId", SqlDbType.SmallInt);
                //    p.Precision = 5;
                //    p.Scale = 0;
                //    p.Value = (object)organizacijaid ?? DBNull.Value;
                //    p.Direction = ParameterDirection.Input;
                //    parameters.Add(p);
                //};
                //{
                //    SqlParameter p = new SqlParameter("@p_TerminiranjeId", SqlDbType.Int);
                //    p.Precision = 5;
                //    p.Scale = 0;
                //    p.Value = (object)terminiranjeid ?? DBNull.Value;
                //    p.Direction = ParameterDirection.Input;
                //    parameters.Add(p);
                //};
                //{
                //    SqlParameter p = new SqlParameter("@p_Godina", SqlDbType.SmallInt);
                //    p.Precision = 5;
                //    p.Scale = 0;
                //    p.Value = (object)godina ?? DBNull.Value;
                //    p.Direction = ParameterDirection.Input;
                //    parameters.Add(p);
                //};
                //{
                //    SqlParameter p = new SqlParameter("@p_Mjesec", SqlDbType.SmallInt);
                //    p.Precision = 5;
                //    p.Scale = 0;
                //    p.Value = (object)mjesec ?? DBNull.Value;
                //    p.Direction = ParameterDirection.Input;
                //    parameters.Add(p);
                //};
                //{
                //    SqlParameter p = new SqlParameter("@p_OrganizacijaIdNosilac", SqlDbType.SmallInt);
                //    p.Precision = 5;
                //    p.Scale = 0;
                //    p.Value = (object)organizacijaidnosilac ?? DBNull.Value;
                //    p.Direction = ParameterDirection.Input;
                //    parameters.Add(p);
                //};
                //{
                //    SqlParameter p = new SqlParameter("@p_PoslovnaGodinaOznaka", SqlDbType.SmallInt);
                //    p.Precision = 5;
                //    p.Scale = 0;
                //    p.Value = (object)poslovnagodinaoznaka ?? DBNull.Value;
                //    p.Direction = ParameterDirection.Input;
                //    parameters.Add(p);
                //};
                //{
                //    SqlParameter p = new SqlParameter("@p_WorkOrderId", SqlDbType.Int);
                //    p.Precision = 5;
                //    p.Scale = 0;
                //    p.Value = DBNull.Value;
                //    p.Direction = ParameterDirection.Input;
                //    parameters.Add(p);
                //};
                //{
                //    SqlParameter p = new SqlParameter("@p_DatumOd", SqlDbType.DateTime);
                //    p.Precision = 5;
                //    p.Scale = 0;
                //    p.Value = DBNull.Value;
                //    p.Direction = ParameterDirection.Input;
                //    parameters.Add(p);
                //};
                //{
                //    SqlParameter p = new SqlParameter("@p_DatumDo", SqlDbType.DateTime);
                //    p.Precision = 5;
                //    p.Scale = 0;
                //    p.Value = DBNull.Value;
                //    p.Direction = ParameterDirection.Input;
                //    parameters.Add(p);
                //};
                //{
                //    SqlParameter p = new SqlParameter("@p_PrikaziZavrseneOperacije", SqlDbType.Bit);
                //    p.Precision = 5;
                //    p.Scale = 0;
                //    p.Value = false;
                //    p.Direction = ParameterDirection.Input;
                //    parameters.Add(p);
                //};

                //DBManager.ExecuteProcedureAsDataReader("Proizvodnja.rptTerminiranjePregled", out dr, parameters.ToArray());


                //var operacije = new List<Operation>();
                //// var dr = comm.ExecuteReader();

                //while (dr.Read())
                //{
                //    Operation op = new Operation();
                //    op.Artikal = (string)dr["Artikal"];
                //    op.ArtikalId = (int)dr["ArtikalId"];
                //    op.VerzijaId = (short)dr["VerzijaId"];
                //    op.Verzija = (string)dr["Verzija"];
                //    op.Sifra = (string)dr["Sifra"];
                //    //op.ArtikalRoot = (string)dr["ArtikalRoot"];
                //    //op.ArtikalIdRoot = (int)dr["ArtikalIdRoot"];
                //    //op.VerzijaIdRoot = (short)dr["VerzijaIdRoot"];
                //    //op.VerzijaRoot = (string)dr["VerzijaRoot"];
                //    //op.SifraRoot = (string)dr["SifraRoot"];

                //    op.ArtikalNad = (string)dr["ArtikalNad"];
                //    op.ArtikalIdNad = (int)dr["ArtikalIdNad"];
                //    op.VerzijaIdNad = (short)dr["VerzijaIdNad"];
                //    op.VerzijaNad = (string)dr["VerzijaNad"];
                //    op.SifraNad = (string)dr["SifraNad"];


                //    op.Level = (int)dr["Level"];
                //    op.RN = (string)dr["RN"];
                //    op.WorkOrderId = (int)dr["WorkOrderId"];
                //    op.RadniNalogNadredjeniId = (int?)dr["WorkOrderIdNadredjeni"];
                //    op.Id = (int)dr["Id"];
                //    op.ResourceId = (int)dr["ResourceId"];
                //    op.Description = dr["Description"] != DBNull.Value ? (string)dr["Description"] : null;
                //    op.ResourceName = dr["ResourceName"] != DBNull.Value ? (string)dr["ResourceName"] : "";
                //    op.Kolicina = (decimal)dr["Kolicina"];
                //    op.PripremnoVreme = (decimal)dr["PripremnoVreme"];
                //    op.VremeZaSeriju = (decimal)dr["VremeZaSeriju"];
                //    op.StartDateTime = dr["StartDateTime"] != DBNull.Value ? DateTime.Parse(dr["StartDateTime"].ToString()) : (DateTime?)null;
                //    op.EndDateTime = dr["EndDateTime"] != DBNull.Value ? DateTime.Parse(dr["EndDateTime"].ToString()) : (DateTime?)null;
                //    operacije.Add(op);
                //}
                int i = 0;
                int? prethodniRN = null;
                #region S1
                //SHEET 1
                int counter = 0;
                foreach (Operation op in operacije)
                {
                    s1.Cell(i + 1, 1).Value = op.WorkOrderId;
                    //if (op.ArtikalIdNad != null)
                    //  s1.Cell(i + 1, 2).Value = "[" + op.ArtikalIdNad.ToString() + "/" + op.VerzijaIdNad.ToString() + "]" + Environment.NewLine;
                    //s1.Cell(i + 1, 2).Value += "[" + op.ArtikalId.ToString() + "/" + op.VerzijaId.ToString() + "]" + op.Artikal;
                    //s1.Cell(i + 1, 3).Value = op.RadniNalogNadredjeniId;
                    s1.Cell(i + 1, 4).Value = op.Id;
                    s1.Cell(i + 1, 5).Value = op.ResourceId;
                    s1.Cell(i + 1, 6).Value = op.Description;
                    s1.Cell(i + 1, 7).Value = op.ResourceName;
                    //s1.Cell(i + 1, 8).Value = op.Kolicina;
                    //s1.Cell(i + 1, 9).Value = op.PripremnoVreme;
                    //s1.Cell(i + 1, 10).Value = op.VremeZaSeriju;
                    s1.Cell(i + 1, 11).Value = op.StartDateTime;
                    s1.Cell(i + 1, 12).Value = op.EndDateTime;
                    s1.Cell(i + 1, 13).Style.DateFormat.Format = "dd.MM.yyyy HH:mm:ss";
                    s1.Cell(i + 1, 14).Style.DateFormat.Format = "dd.MM.yyyy HH:mm:ss";

                    if (prethodniRN != op.WorkOrderId)
                    {
                        s1.Cell(i + 1, 11).Style.Font.Bold = true;

                    }


                    if (operacije.Count == i + 1 || operacije[i + 1].WorkOrderId != op.WorkOrderId)
                    {
                        s1.Cell(i + 1, 12).Style.Font.Bold = true;
                        s1.Row(i + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                        s1.Cell(i - counter + 1, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        s1.Range(i - counter + 1, 2, i + 1, 2).Merge();
                        s1.Cell(i - counter + 1, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        s1.Range(i - counter + 1, 2, i + 1, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        counter = 0;
                    }
                    else
                        counter++;
                    s1.Row(i + 1).Style.Fill.BackgroundColor = colors[op.Level];

                    //if (op.Kolicina == 0)
                    //    s1.Row(i + 1).Style.Font.Bold = true;


                    prethodniRN = op.WorkOrderId;

                    i++;
                }
                s1.Columns(1, 12).AdjustToContents();

                #endregion S1

                #region S2

                DateTime? minDatum = operacije.Min(ope => ope.StartDateTime);
                DateTime? maxDatum = operacije.Max(ope => ope.EndDateTime);
                int diff = (int)maxDatum.Value.Date.Subtract(minDatum.Value.Date).TotalDays;

                //SHEET 2
                i = 0;
                int rowIndex = 1;
                for (int k = 0; k <= diff; k++)
                {
                    s2.Cell(1, 2 + k).Value = minDatum.Value.Date.AddDays(k);
                    s2.Style.DateFormat.Format = "dd.MM.yyyy";
                    s2.Style.Font.Bold = true;
                    s2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
                while (1 == 1)
                {
                    if (operacije.Count == i)
                        break;
                    var op = operacije[i];
                    int j = 0;


                    while (1 == 1)
                    {
                        var opp = operacije[i + j];
                        if (opp.StartDateTime != null && opp.EndDateTime != null)
                        {
                            if (s2.Cell(rowIndex + 1, 2 + (int)opp.StartDateTime.Value.Date.Subtract(minDatum.Value.Date).TotalDays).Value.ToString() != "")
                                s2.Cell(rowIndex + 1, 2 + (int)opp.StartDateTime.Value.Date.Subtract(minDatum.Value.Date).TotalDays).Value += Environment.NewLine +
                                  opp.StartDateTime.Value.ToString("HH:mm") + "-" + opp.EndDateTime.Value.ToString("HH:mm") + " (" + opp.Id.ToString() + ")" + opp.Description;
                            else
                                s2.Cell(rowIndex + 1, 2 + (int)opp.StartDateTime.Value.Date.Subtract(minDatum.Value.Date).TotalDays).Value =
                                  opp.StartDateTime.Value.ToString("HH:mm") + "-" + opp.EndDateTime.Value.ToString("HH:mm") + " (" + opp.Id.ToString() + ")" + opp.Description;

                            //s2.Cell(rowIndex + 1, 2 + (int)opp.StartDateTime.Value.Date.Subtract(minDatum.Value.Date).TotalDays).Style.Fill.BackgroundColor = colors[op.Level];

                            s2.Cell(rowIndex + 1, 2 + (int)opp.StartDateTime.Value.Date.Subtract(minDatum.Value.Date).TotalDays).Style.Alignment.WrapText = true;
                            s2.Cell(rowIndex + 1, 2 + (int)opp.StartDateTime.Value.Date.Subtract(minDatum.Value.Date).TotalDays).DataType = XLCellValues.Text;
                            s2.Cell(rowIndex + 1, 2 + (int)opp.StartDateTime.Value.Date.Subtract(minDatum.Value.Date).TotalDays).Style.Font.FontSize = 7;
                            s2.Cell(rowIndex + 1, 2 + (int)opp.StartDateTime.Value.Date.Subtract(minDatum.Value.Date).TotalDays).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            s2.Cell(rowIndex + 1, 2 + (int)opp.StartDateTime.Value.Date.Subtract(minDatum.Value.Date).TotalDays).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            s2.Cell(rowIndex + 1, 2 + (int)opp.StartDateTime.Value.Date.Subtract(minDatum.Value.Date).TotalDays).Style.Font.Bold = false;
                        }
                        j++;
                        if (operacije.Count == i + j || opp.WorkOrderId != operacije[i + j].WorkOrderId)
                            break;

                    }

                    i = i + j - 1;

                    var minope = operacije.Where(ope => ope.WorkOrderId == op.WorkOrderId && ope.StartDateTime != null).ToList();

                    int rnStart = 0;
                    if (minope.Count != 0)
                        rnStart = (int)minope.Min(ope => ope.StartDateTime).Value.Date.Subtract(minDatum.Value.Date).TotalDays;
                    var maxope = operacije.Where(ope => ope.WorkOrderId == op.WorkOrderId && ope.EndDateTime != null).ToList();
                    int rnEnd = 0;
                    if (maxope.Count != 0)
                        rnEnd = (int)maxope.Max(ope => ope.EndDateTime).Value.Date.Subtract(minDatum.Value.Date).TotalDays;

                    s2.Cell(rowIndex + 1, 1).Value = op.WorkOrderId;
                    //if (op.WorkOrderId == 91100)
                    //{
                    //  int ss = 0;
                    //}
                    if (rnStart != 0 || rnEnd != 0)
                        s2.Range(rowIndex + 1, 2 + rnStart, rowIndex + 1, 2 + rnEnd).Style.Fill.BackgroundColor = colors[op.Level];

                    //s2.Cell(rowIndex + 1, 1).Style.Fill.BackgroundColor = colors[op.Level];
                    //s2.Range(rowIndex + 1, 2 + rnStart, rowIndex + 1, 2 + rnEnd).Merge();



                    //s2.Row(rowIndex + 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    i++;
                    rowIndex++;
                }
                s2.Rows(1, rowIndex).AdjustToContents();
                s2.Columns(1, diff + 2).AdjustToContents();

                #endregion S2



                #region s3

                //DateTime? minDatum = operacije.Min(ope => ope.StartDateTime);
                //DateTime? maxDatum = operacije.Max(ope => ope.EndDateTime);
                //int diff = (int)maxDatum.Value.Date.Subtract(minDatum.Value.Date).TotalDays;

                //SHEET 3

                List<Operation> masine = operacije
                .GroupBy(p => p.ResourceId)
                .Select(g => g.First()).OrderBy(m => m.ResourceId)
                .ToList();





                for (int l = 0; l < masine.Count; l++)
                {
                    minDatum = operacije.Where(ope => ope.ResourceId == masine[l].ResourceId).Min(ope => ope.StartDateTime);
                    maxDatum = operacije.Where(ope => ope.ResourceId == masine[l].ResourceId).Max(ope => ope.EndDateTime);
                    if (minDatum != null && maxDatum != null)
                    {
                        diff = (int)maxDatum.Value.Date.Subtract(minDatum.Value.Date).TotalDays;
                        var ts = minDatum.Value.Date.Add(TimeSpan.FromMinutes(420));

                        s3 = xLWorkbook.Worksheets.Add(masine[l].ResourceId.ToString() + "-" + (masine[l].ResourceName != null ? masine[l].ResourceName : "").Replace("[", "").Replace("]", "").Replace("?", "").Replace("\\", "").Replace("/", ""));

                        for (int s = 1; s < 190; s++)
                        {
                            DateTime val = ts.Add(TimeSpan.FromMinutes(s * 5));
                            s3.Cell(s + 2, 1 + 0 * (diff + 1)).Value = val;
                            s3.Cell(s + 2, 1 + 0 * (diff + 1)).DataType = XLCellValues.DateTime;
                            s3.Column(1 + 0 * (diff + 1)).Style.DateFormat.Format = "HH:mm";
                            s3.Column(1 + 0 * (diff + 1)).Style.Font.Bold = true;
                        }

                        var masina = masine[l];
                        s3.Cell(1, 1 + 0 * (diff + 1)).Value = "[" + masina.ResourceId.ToString() + "]" + masina.ResourceName;
                        s3.Cell(1, 1 + 0 * (diff + 1)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        i = 0;
                        rowIndex = 1;
                        for (int k = 0; k <= diff; k++)
                        {
                            s3.Cell(2, 2 + k + 0 * (diff + 1)).Value = minDatum.Value.Date.AddDays(k);
                            s3.Cell(2, 2 + k + 0 * (diff + 1)).Style.DateFormat.Format = "dd.MM.yyyy";
                            s3.Cell(2, 2 + k + 0 * (diff + 1)).Style.Font.Bold = true;
                            s3.Cell(2, 2 + k + 0 * (diff + 1)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            ts = minDatum.Value.Date.AddDays(k).Date.Add(TimeSpan.FromMinutes(420));


                            var operacijeMasine = operacije.Where(ope => ope.EndDateTime != null).Where(o => o.ResourceId == masina.ResourceId && ((o.StartDateTime.Value.Date < minDatum.Value.Date.AddDays(k).Date && o.EndDateTime.Value.Date > minDatum.Value.Date.AddDays(k).Date) || o.StartDateTime.Value.Date == minDatum.Value.Date.AddDays(k).Date || o.EndDateTime.Value.Date == minDatum.Value.Date.AddDays(k).Date)).OrderBy(op => op.StartDateTime).ToList();
                            //if (operacijeMasine.Count > 0)
                            //{
                            //  int jj = 0;
                            //}


                            for (int s = 1; s < 190; s++)
                            {

                                DateTime val = ts.Add(TimeSpan.FromMinutes(s * 5));
                                var opp = operacijeMasine.Where(o => o.StartDateTime <= val && o.EndDateTime >= val).ToList();
                                if (opp.Count() > 0)
                                {
                                    s3.Cell(s + 2, 2 + k + 0 * (diff + 1)).Style.Fill.BackgroundColor = colors[1];
                                    s3.Cell(s + 2, 2 + k + 0 * (diff + 1)).Value = opp.First().WorkOrderId.ToString();
                                    //s3.Cell(s + 2, 2 + k + 0 * (diff + 1)).Comment.AddText(val.ToString("dd.MM.yyyy HH:mm:ss"));
                                    s3.Cell(s + 2, 2 + k + 0 * (diff + 1)).Style.Alignment.WrapText = true;
                                }


                            }



                            //s3.Column(2 + k + l * (diff + 1)).Style.Fill.BackgroundColor = colors[l % 2];
                            s3.Column(2 + k + 0 * (diff + 1)).AdjustToContents();
                        }



                        s3.Range(1, 1 + (diff + 1) * 0, 1, (diff + 1) * (0 + 1)).Merge();
                        ////////s3.Range(1, 1 + (diff + 1) * l, 1, (diff + 1) * (l + 1)).Style.Fill.BackgroundColor = colors[1];
                        s3.Range(1, 1 + (diff + 1) * 0, 1, (diff + 1) * (0 + 1)).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        //s3.Range(1, 15, 1, 27).Merge();
                        //s3.Range(1, 28, 1, 40).Merge();


                        while (1 == 1)
                        {
                            if (operacije.Count == i)
                                break;
                            var op = operacije[i];
                            i = i + 1;
                        }
                    }
                    //dr.Close();



                }

                #endregion s3

                MemoryStream ms = new MemoryStream();
                xLWorkbook.SaveAs(ms);
                FileStream file = new FileStream(filename, FileMode.Create, FileAccess.Write);
                ms.WriteTo(file);
                file.Close();
                ms.Close();
                System.Diagnostics.Process.Start(filename);
                //conn.Close();
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
            }

        }
    }
}
