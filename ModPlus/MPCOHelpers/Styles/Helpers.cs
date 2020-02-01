#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Xml.Linq;
using Autodesk.AutoCAD.Runtime;

namespace ModPlus.MPCOHelpers.Styles
{
    public static class Helpers
    {
        /// <summary>
        /// Получение целочисленного параметра из стиля
        /// </summary>
        /// <param name="style">Стиль</param>
        /// <param name="name">Имя параметра</param>
        /// <param name="defaultValue">Значение по умолчанию, возвращаемое в случае неудачного парсинга</param>
        /// <returns>Целочисленное значение параметра</returns>
        public static int GetIntegerPropery(IMPCOStyle style, string name, int defaultValue)
        {
            return int.TryParse(GetProperty(style, name).ToString(), out int i) ? i : defaultValue;
        }
        private static object GetProperty(IMPCOStyle style, string name)
        {
            return style.Properties.FirstOrDefault(x => x.Name.Equals(name))?.Value;
        }




        /// <summary>
        /// Сохранение системного стиля в специальный файл
        /// </summary>
        /// <param name="style">Системный стиль</param>
        public static void SaveSystemStyleToFile(IMPCOStyle style)
        {
            // Загружаем файл системных стилей
            var xFile = XElement.Load(MPCOInitialize.SystemStylesFile);
            if (!xFile.Elements("Style").Any(x =>
            {
                var xAttribute = x.Attribute("Name");
                return xAttribute != null && xAttribute.Value.Equals(style.Name);
            }))
            {
                XElement newStyle = new XElement("Style");
                newStyle.SetAttributeValue("Name", style.Name);
                newStyle.SetAttributeValue("FunctionName", style.FunctionName);
                newStyle.SetAttributeValue("Description", style.Description);
                newStyle.SetAttributeValue("Guid", style.Guid);
                newStyle.SetAttributeValue("StyleType", style.StyleType);
                foreach (MPCOStyleProperty property in style.Properties)
                {
                    var propXel = new XElement("Property");
                    propXel.SetElementValue("Name", property.Name);
                    propXel.SetElementValue("Value", property.Value);
                    propXel.SetElementValue("DisplayName", property.DisplayName);
                    propXel.SetElementValue("Description", property.Description);
                    propXel.SetElementValue("StylePropertyType", property.StylePropertyType);
                    if (property.MinInt != null)
                        propXel.SetElementValue("MinInt", property.MinInt.Value);
                    if (property.MaxInt != null)
                        propXel.SetElementValue("MaxInt", property.MaxInt.Value);
                    if (property.MinDouble != null)
                        propXel.SetElementValue("MinDouble", property.MinDouble.Value);
                    if (property.MaxDouble != null)
                        propXel.SetElementValue("MaxDouble", property.MaxDouble.Value);
                    newStyle.Add(propXel);
                }
                xFile.Add(newStyle);
                xFile.Save(MPCOInitialize.SystemStylesFile);
            }
        }
        /// <summary>
        /// Получение списка системных стилей
        /// </summary>
        /// <returns></returns>
        public static List<MPCOstyle> GetSystemStyle()
        {
            var systemStyles = new List<MPCOstyle>();
            // Загружаем файл системных стилей
            var xFile = XElement.Load(MPCOInitialize.SystemStylesFile);
            foreach (XElement styleXel in xFile.Elements("Style"))
            {
                var style = GetStyleFromXElement(styleXel);
                if(style != null)
                    systemStyles.Add(style);
            }
            return systemStyles;
        }

        private static MPCOstyle GetStyleFromXElement(XElement xElement)
        {
            var style = new MPCOstyle
            {
                Name = xElement.Attribute("Name")?.Value,
                FunctionName = xElement.Attribute("FunctionName")?.Value,
                Description = xElement.Attribute("Description")?.Value,
                Guid = xElement.Attribute("Guid")?.Value,
                StyleType = GetStyleTypeByString(xElement.Attribute("StyleType")?.Value)
            };
            foreach (var propertyXElement in xElement.Elements("Property"))
            {
                MPCOStyleProperty property = new MPCOStyleProperty();
                var element = propertyXElement.Element("Name");
                if (element != null) property.Name = element.Value;
                element = propertyXElement.Element("Value");
                if (element != null) property.Value = element.Value;
                element = propertyXElement.Element("DisplayName");
                if (element != null) property.DisplayName = element.Value;
                element = propertyXElement.Element("Description");
                if (element != null) property.Description = element.Value;
                element = propertyXElement.Element("StylePropertyType");
                if (element != null) property.StylePropertyType = GetPropertyTypeByString(element.Value);
                element = propertyXElement.Element("MinInt");
                if (element != null) property.MinInt = int.Parse(element.Value);
                element = propertyXElement.Element("MaxInt");
                if (element != null) property.MaxInt = int.Parse(element.Value);
                element = propertyXElement.Element("MinDouble");
                if (element != null) property.MinDouble = double.Parse(element.Value);
                element = propertyXElement.Element("MaxDouble");
                if (element != null) property.MaxDouble = double.Parse(element.Value);

                if(!string.IsNullOrEmpty(property.Name) & property.Value != null)
                    style.Properties.Add(property);
            }
            if (!string.IsNullOrEmpty(style.Name) & !string.IsNullOrEmpty(style.FunctionName))
            {
                return style;
            }
            return null;
        }

        private static MPCOStyleType GetStyleTypeByString(string type)
        {
            if (type == "System")
                return MPCOStyleType.System;
            if (type == "User")
                return MPCOStyleType.User;
            return MPCOStyleType.System;
        }

        private static MPCOStylePropertyType GetPropertyTypeByString(string type)
        {
            if(type == "Int")
                return MPCOStylePropertyType.Int;
            if(type == "Double")
                return MPCOStylePropertyType.Double;

            return MPCOStylePropertyType.String;
        }
    }

    public class MPCOstyle : IMPCOStyle
    {
        public MPCOstyle()
        {
            Properties = new List<MPCOStyleProperty>();
        }
        [Category("Основные")]
        [DisplayName("Название стиля")]
        [Description("Description")]
        [ReadOnly(true)]
        public string Name { get; set; }
        [Browsable(false)]
        public string FunctionName { get; set; }
        public string Description { get; set; }
        public string Guid { get; set; }
        [Browsable(false)]
        public MPCOStyleType StyleType { get; set; }
        [Browsable(false)]
        public List<MPCOStyleProperty> Properties { get; set; }
    }

    public class StyleEditorWork
    {
        private StyleEditor styleEditor;
        [CommandMethod("ModPlus", "mpStyleEditor", CommandFlags.Modal)]
        public void OpenStyleEditor()
        {
            if (styleEditor == null)
            {
                styleEditor = new StyleEditor();
                styleEditor.Closed += styleEditor_Closed;
            }
            if (styleEditor.IsLoaded) styleEditor.Activate();
            else AcApp.ShowModalWindow(AcApp.MainWindow.Handle, styleEditor, false);
        }

        void styleEditor_Closed(object sender, EventArgs e)
        {
            styleEditor = null;
        }
    }
}
