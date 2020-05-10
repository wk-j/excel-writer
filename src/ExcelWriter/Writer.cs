using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelWriter {
    public class Writer {

        private void CreateHeader(IEnumerable<string> heads, IXLWorksheet ws) {
            var headRange = Enumerable.Range('A', heads.Count());
            foreach (var (x, i) in headRange.Select((x, i) => (x, i))) {
                var cellChar = Char.ToString((char)x);
                var value = heads.ElementAt(i);
                var cell = $"{cellChar}1";
                ws.Cell(cell).Value = value;
            }
        }

        private void CreateBody(IEnumerable<string> heads, IEnumerable<Dictionary<string, object>> records, IXLWorksheet ws) {
            var headRange = Enumerable.Range('A', heads.Count());
            foreach (var (row, i) in records.Select((x, i) => (x, i + 2))) {
                foreach (var (column, ci) in heads.Select((x, i) => (x, i))) {
                    var cellChar = (char)headRange.ElementAt(ci);
                    var value = row[heads.ElementAt(ci)];
                    var cell = $"{cellChar}{i}";
                    Console.WriteLine(cell);
                    ws.Cell(cell).Value = value;
                }
            }
        }

        public void Write(IEnumerable<string> heads, IEnumerable<Dictionary<string, object>> records, string file) {
            using var wb = new XLWorkbook();

            var ws = wb.AddWorksheet("Data");

            CreateHeader(heads, ws);
            CreateBody(heads, records, ws);

            wb.SaveAs(file);
        }
    }
}
