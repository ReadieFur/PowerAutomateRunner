using System;
using System.Linq;
using System.Windows.Automation;

#nullable enable
namespace PowerAutomateRunner
{
    public static class Helpers
    {
        public static AutomationElement? FindElement(this AutomationElement parent,
            Func<AutomationElement.AutomationElementInformation, bool> predicate, bool recursive = false)
        {
            AutomationElementCollection elements = parent.FindAll(TreeScope.Children, Condition.TrueCondition);
            AutomationElement? element = elements.Cast<AutomationElement>().FirstOrDefault(e => predicate(e.Current));
            if (element != null)
                return element;

            if (!recursive)
                return null;

            foreach (AutomationElement child in elements)
            {
                element = FindElement(child, predicate, recursive);
                if (element != null)
                    return element;
            }

            return null;
        }

        public static string[] Split(this string self, string separator, StringSplitOptions options = StringSplitOptions.None) =>
            self.Split(new[] { separator }, options);
    }
}
