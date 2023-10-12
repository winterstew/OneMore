
namespace River.OneMoreAddIn.Commands
{
    using River.OneMoreAddIn.Models;
    using River.OneMoreAddIn.Styles;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using System.Xml.Linq;
    using Resx = Properties.Resources;

    /// <summary>
    /// Merge multiple consecutive paragraphs into a single paragraph. 
    /// </summary>
    internal class MergeParagraphCommand : Command
    {
        private Page page;
        private XNamespace ns;
        private IEnumerable<XElement> range;
        private bool all;


        public MergeParagraphCommand()
        {
        }

        public override async Task Execute(params object[] args)
        {
            var result = UIHelper.ShowQuestion(Resx.MergeParagraphCommand_option, false, true);
            if (result == DialogResult.Cancel)
            {
                return;
            }

            all = result == DialogResult.Yes;

            logger.StartClock();

            using var one = new OneNote();
            page = one.GetPage(OneNote.PageDetail.Selection);
            ns = page.Namespace;

            range = page.GetSelectedElements();
            logger.WriteLine($"found {range.Count()} runs, scope={page.SelectionScope}");

            var modified = MergeConsecutive();

            logger.WriteTime("saving", true);

            if (modified)
            {
                await one.Update(page);
                logger.WriteTime("saved");
            }
        }


        public bool MergeConsecutive()
        {
            
            var elements = range
                .Select(e => e.Parent)
                .Distinct()
                .ToList();

            if (elements?.Any() != true)
            {
                logger.WriteLine("no mergable paragraphs found");
                return false;
            }

            //System.Diagnostics.Debugger.Launch();
            //logger.WriteLine($"found {elements.Count} mergable paragraphs");

            var modified = false;

            var currentCDATA = new XCData("");
            var currentT = new XElement(ns + "T", currentCDATA);
            var currentOE = new XElement(ns + "OE", currentT);
            var thisCDATA = new XCData("");
            var thisT = new XElement(ns + "T", thisCDATA);
            var thisOE = new XElement(ns + "OE", thisT);

            var noOE = true;

            //grab the first paragraph with a T element
            foreach (var element in elements)
            {
                if (element is XElement coe && coe.Name.LocalName == "OE")
                {
                    if (coe.LastNode is XElement ct && ct.Name.LocalName == "T")
                    {
                        if (ct.LastNode is XCData cd && cd.ToString().Trim().Length > 0)
                        {
                            currentCDATA = cd;
                        }
                        else
                        {
                            currentCDATA = new XCData("");
                            ct.Add(currentCDATA);
                        }
                        currentT = ct;
                        currentOE = coe;
                        noOE = false;
                        break;
                    }
                }
            }
            // if there are no paragraphs in this range just return
            if (noOE) { return modified; }

            foreach (var element in elements)
            {
                if (element == currentOE) { continue; }
                if (element is XElement oe && oe.Name.LocalName == "OE")
                {
                    if (oe.LastNode is XElement t && t.Name.LocalName == "T")
                    {
                        // assign "this" variables or skip line if it is empty
                        if (t.LastNode is XCData d && d.Value.Trim().Length > 0)
                        {
                            thisOE = oe;

                            //I could cycle through all of the T elements and merge them here
                            //however in my use case each OE element has only one T element and 
                            // each T has only one CDATA
                            thisT = t;

                            //I could cycle through all the CDATA elements and merge them here
                            //however in my use case each T element has only one CDATA
                            //merge contents of this paragraph into the current one
                            thisCDATA = d;
                        }
                        else { continue; }
                    }

                    // if this paragraph is preceded by an empty paragraph, make it the new current
                    if (thisOE.PreviousNode is XElement poe && poe.Name.LocalName == "OE")
                    {
                        // does previous paragraph end with an empty run?
                        if (poe.LastNode is XElement pt && pt.Name.LocalName == "T")
                        {
                            if (pt.LastNode is XCData pd && pd.Value.Trim().Length == 0)
                            {
                                // make this the new current paragraph
                                currentOE = thisOE;
                                currentT = thisT;
                                currentCDATA = thisCDATA;
                                continue;
                            }
                        }
                    }
                    // append this elements text to the current paragraph
                    currentCDATA.Value = currentCDATA.Value + " " + thisCDATA.Value;
                    //remove this element
                    thisOE.Remove();
                    modified = true;  
                }
                //there is a case where if an element is not an "OE" element it will be 
                //ignored and maybe shunted to the end?  
            }
            return modified;
        }
    }
}
