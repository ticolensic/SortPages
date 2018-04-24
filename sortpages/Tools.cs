using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace SortPages
{
    public static class Tools
    {
        public static void SortByAttribute(this XElement source, string attribute, Boolean ascending = true)
        {
            if (source == null) throw new ArgumentNullException("source");

            if (source.HasElements && ascending)
            {
                List<XElement> sortedChildren = source.Elements().OrderBy(c => (string)c.Attribute(attribute)).ToList();
                source.RemoveNodes();
                sortedChildren.ForEach(c => source.Add(c));
            }
            if (source.HasElements && !ascending)
            {
                List<XElement> sortedChildren = source.Elements().OrderByDescending(c => (string)c.Attribute(attribute)).ToList();
                source.RemoveNodes();
                sortedChildren.ForEach(c => source.Add(c));
            }
        }
    }
}