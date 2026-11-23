using System;
using System.Xml.Linq;

namespace ZsuTools
{
    /// <summary>
    /// Fluent-білдер XML для Excel Ribbon (customUI).
    /// Кнопка додається до останньої доданої групи, група — до останньої доданої вкладки.
    /// Екранування спецсимволів у атрибутах виконує XElement.
    /// </summary>
    public class CustomUIBuilder
    {
        private static readonly XNamespace Ns = "http://schemas.microsoft.com/office/2009/07/customui";

        private readonly XElement _tabs = new XElement(Ns + "tabs");
        private XElement _currentTab;
        private XElement _currentGroup;

        public CustomUIBuilder AddTab(string id, string label)
        {
            _currentTab = new XElement(Ns + "tab",
                new XAttribute("id", id),
                new XAttribute("label", label));
            _tabs.Add(_currentTab);
            _currentGroup = null;
            return this;
        }

        public CustomUIBuilder AddGroup(string id, string label)
        {
            if (_currentTab == null)
                throw new InvalidOperationException("Спочатку викличте AddTab().");

            _currentGroup = new XElement(Ns + "group",
                new XAttribute("id", id),
                new XAttribute("label", label));
            _currentTab.Add(_currentGroup);
            return this;
        }

        public CustomUIBuilder AddButton(string id, string label, string onAction, string imageMso,
            bool large = true)
        {
            if (_currentGroup == null)
                throw new InvalidOperationException("Спочатку викличте AddGroup().");

            _currentGroup.Add(new XElement(Ns + "button",
                new XAttribute("id", id),
                new XAttribute("label", label),
                new XAttribute("size", large ? "large" : "normal"),
                new XAttribute("imageMso", imageMso),
                new XAttribute("showImage", "true"),
                new XAttribute("onAction", onAction)));
            return this;
        }

        /// <summary>
        /// Те саме, що AddButton, але додає кнопку лише якщо <paramref name="condition"/> == true.
        /// Зручно для #if-подібних умов без розриву ланцюжка викликів.
        /// </summary>
        public CustomUIBuilder AddButtonIf(bool condition, string id, string label, string onAction,
            string imageMso, bool large = true)
        {
            return condition ? AddButton(id, label, onAction, imageMso, large) : this;
        }

        public string Build()
        {
            var doc = new XElement(Ns + "customUI",
                new XElement(Ns + "ribbon", _tabs));
            return doc.ToString(SaveOptions.None);
        }
    }
}
