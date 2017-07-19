using System.Collections.Generic;
using System.ComponentModel;

namespace ModPlus.MPCOHelpers.Styles
{
    public interface IMPCOStyle
    {
        string Name { get; set; }
        string FunctionName { get; set; }
        string Description { get; set; }
        string Guid { get; }
        MPCOStyleType StyleType { get; set; }
        List<MPCOStyleProperty> Properties { get; set; }
    }

    public enum MPCOStyleType
    {
        System = 1,
        User = 2
    }

    public class MPCOStyleStringProperty
    {
        public string Description { get; set; }
        public string Name { get; set; }
    
    }

    public class MPCOStyleIntProperty
    {
        
    }
    public class MPCOStyleProperty
    {
        [Category("Property")]
        public object Value { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public MPCOStylePropertyType StylePropertyType { get; set; }
        public int? MinInt { get; set; }
        public int? MaxInt { get; set; }
        public double? MinDouble { get; set; }
        public double? MaxDouble { get; set; }
    }

    public enum MPCOStylePropertyType
    {
        String = 1,
        Int = 2,
        Double = 3
    }
}
