using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelWriter.Tests {
    public class UnitTest1 {
        [Fact]
        public void Test1() {
            var head = new[] {
                "A",
                "B",
                "C",
                "D"
            };
            var writer = new Writer();
            var data = new List<Dictionary<string, object>> {
                new Dictionary<string, object> {
                    {"A", "A100"},
                    {"B", "B100"},
                    {"C", "C100"},
                    {"D", "D100"}
                },
                new Dictionary<string, object> {
                    {"A", "A200"},
                    {"B", "B200"},
                    {"C", "C200"},
                    {"D", "D200"}
                }
            };

            writer.Write(head, data, "../../../../../outputs/A.xlsx");
        }
    }
}
