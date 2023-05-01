using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Tabula;

namespace Streamlining_Customer_Verification
{
    internal class Validation
    {
        public static DataTable DataValid (Table tbl, DataTable dataTable)
        {
            var newRow = dataTable.NewRow();
            dataTable.Rows.Add(newRow);
            bool qtDt = false;
            int qtNbr = 0;
            foreach (var cell in tbl.Cells)
            {
                string st = cell.ToString();
                string temp = null;

                if (st.Contains("Quote Date"))
                    qtDt = true;
                else if (qtDt)
                {
                    newRow["Quote Date"] = st;
                    qtDt = false;
                }
                else if (st.Contains("Quote Number") || qtNbr == 1)
                    qtNbr++;
                else if (qtNbr == 2)
                {
                    newRow["Quote Number"] = st;
                    qtNbr = 0;
                }
                else if (st.Contains("Customer Name"))
                    newRow["Customer Name"] = st.Substring(13);
                else if (st.Contains("Phone Number"))
                {
                    newRow["Country Code"] = st.Substring(12, 3);
                    newRow["Phone No."] = Regex.Replace(st.Substring(16), @"\D", "");
                }                    
                else if (st.Contains("Address Line 1"))
                    newRow["Address Line 1"] = st.Substring(14);
                else if (st.Contains("Address Line 2"))
                    newRow["Address Line 2"] = st.Substring(14); // 14
                else if (st.Contains("Postal Code"))
                    newRow["Postal Code"] = Regex.Replace(st.Substring(11), @"\s+", ""); // 11
                else if (st.Contains("Locality"))
                {
                    string[] parts1 = st.Split('\r');
                    string[] parts2 = parts1[1].Split(',');
                    newRow["Locality"] = parts2[0].Trim();
                    newRow["Country"] = parts2[1].Trim();                    
                }                    
                // newRow["Country"] = st.Substring();
                else if (st.Contains("Account Number"))
                    newRow["IBAN"] = Regex.Replace(st.Substring(15), @"\s+", "");
                else if (st.Contains("Product Id"))
                    newRow["Product ID"] = st.Substring(10);
                else if (st.Contains("Product Description"))
                    newRow["Product Description"] = st.Substring(19);
                else if (st.Contains("Invoice Number"))
                    newRow["Invoice Number"] = Regex.Replace(st.Substring(14), @"[\W_]+", "");
                else if (st.Contains("Reason for Refund"))
                    newRow["Reason for Refund"] = st.Substring(17);
                else if (st.Contains("Suggested Amount"))
                {
                    temp = st.Substring(16);
                    if (!temp.Contains("EUR"))
                        temp = CurrencyConversion.ConvertToEuros(temp);
                    newRow["Suggested Amount"] = temp;
                }
                else if (st.Contains("Salesman Name"))
                    newRow["Salesman Name"] = st.Substring(13);
                else if (st.Contains("Salesman Id"))
                    newRow["Salesman Id"] = Regex.Replace(st.Substring(11), @"[\W_]+", "");
            }            
            return dataTable;
        }

        public static DataTable CheckElig (DataTable dtFile, DataTable billDt, DataTable eligDt)
        {
            for (int i = 0; i < dtFile.Rows.Count; i++)
            {
                // Check if the string is found in the specified column of the current row
                if (dtFile.Rows[i]["Product IDs"].ToString().IndexOf(searchString, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Add the row number to the list of rows with data
                    fndRow = i;
                }
            }
            return dtFile;
        }
    }
}
