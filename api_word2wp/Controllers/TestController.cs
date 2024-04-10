using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

[ApiController]
[Route("[controller]")]
public class WordController : ControllerBase
{
    [HttpPost("changeNumberingAlignment")]
    public async Task<IActionResult> ChangeNumberingAlignment( IFormFile file)
    {
        try
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            // Modify numbering alignment
            byte[] modifiedFile = await ModifyNumberingAlignmentAsync(file);

            // Return the modified Word document
            return File(modifiedFile, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "modified_document.docx");
        }
        catch (Exception ex)
        {
            return StatusCode(StatusCodes.Status500InternalServerError, $"An error occurred: {ex.Message}");
        }
    }

    private async Task<byte[]> ModifyNumberingAlignmentAsync(IFormFile inputFile)
    {
        using (MemoryStream memoryStream = new MemoryStream())
        {
            // Read the Word document into a MemoryStream
            await inputFile.CopyToAsync(memoryStream);

            // Open the Word document using Open XML SDK
            using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
            {
                var numberingDefinitionsPart = doc.MainDocumentPart.NumberingDefinitionsPart;
                if (numberingDefinitionsPart != null)
                {
                    // Retrieve the numbering definitions
                    Numbering numbering = numberingDefinitionsPart.Numbering;

                    if (numbering != null)
                    {
                        // Loop through all the numbering instances
                        foreach (var num in numbering.Descendants<NumberingInstance>())
                        {
                            // Get the abstract numbering definition ID from the numbering instance
                            string abstractNumId = num.AbstractNumId.Val;

                            // Find the corresponding abstract numbering definition
                            AbstractNum abstractNum = numbering.Descendants<AbstractNum>().FirstOrDefault(an => an.AbstractNumberId == abstractNumId);

                            if (abstractNum != null)
                            {
                                // Modify the numbering level justification
                                foreach (var level in abstractNum.Descendants<Level>())
                                {
                                    level.LevelJustification = new LevelJustification { Val = LevelJustificationValues.Left };
                                }
                            }
                        }
                    }
                }
            }

            // Convert the MemoryStream to a byte array and return it
            return memoryStream.ToArray();
        }
    }

}
