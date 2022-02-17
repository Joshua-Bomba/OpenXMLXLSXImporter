using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLXLXSImporter.Builders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLXSImporter
{
    public interface ISheetProperties
    {
        /// <summary>
        /// This is the Sheet name
        /// </summary>
        string Sheet { get; }

        void LoadConfig(ISpreadSheetInstructionBuilder builder);

        Task ResultsProcessed(ISpreadSheetQueryResults query);
    }
}
