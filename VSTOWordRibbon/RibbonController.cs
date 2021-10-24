using Microsoft.Office.Core;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using VSTOWordRibbon.Properties;

namespace VSTOWordRibbon
{
    [ComVisible(true)]
    public class RibbonController : Microsoft.Office.Core.IRibbonExtensibility
    {
        private Microsoft.Office.Core.IRibbonUI _ribbonUi;

        public string GetCustomUI(string ribbonID) =>
            @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                        <ribbon>
                           <tabs>
                                <tab id='sample_tab' label='ALASDAIR REID'>
                                    <group id='sample_group' label='Operations'>
                                        <button id='do_1' label='Insert Details' size='large' getImage='OnDo1GetImage' onAction='OnDo1Click'/>
                                    </group>
                                </tab>
                            </tabs>
                        </ribbon>
                    </customUI>";

        public void OnLoad(Microsoft.Office.Core.IRibbonUI ribbonUI)
        {
            _ribbonUi = ribbonUI;
        }

        public void OnDo1Click(Microsoft.Office.Core.IRibbonControl control)
        {
            string baseUrl = "http://www.alasdair-reid.com";

            var document = Globals.ThisAddIn.Application.ActiveDocument;

            // first paragraph heading
            var firstPara = document.Paragraphs.Add();
            firstPara.Format.SpaceAfter = 10f;
            firstPara.Range.Text = "AR Auto Inserted Heading from VSTO C# .NET Controls";
            firstPara.Range.InsertParagraphAfter();

            // Set formatting
            var rng = document.Paragraphs[1].Range;
            rng.Font.Size = 14;
            rng.Font.Name = "Arial";
            rng.Font.Bold = 1;

            // Second paragraph
            var secondPara = document.Paragraphs.Add();
            secondPara.Format.SpaceAfter = 10f;
            secondPara.Range.Text = baseUrl;
            secondPara.Range.InsertParagraphAfter();

            // add image
            var thirdPara = document.Paragraphs.Add();
            thirdPara.Format.SpaceAfter = 10f;
            thirdPara.Range.InsertParagraphAfter();

            string imagesurl = string.Format("{0}/img/logo/logo_dark.png", baseUrl);
            thirdPara.Range.InlineShapes.AddPicture(imagesurl);
        }


        public Bitmap OnDo1GetImage(Microsoft.Office.Core.IRibbonControl control) => Resources.Do1_128px;
    }
}
