//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace OLECalApplication
{
    using System;
    using System.Collections.Generic;
    
    public partial class tblmimic
    {
        public int mid { get; set; }
        public string MachineOnTime { get; set; }
        public string OperatingTime { get; set; }
        public string SetupTime { get; set; }
        public string IdleTime { get; set; }
        public string MachineOffTime { get; set; }
        public string BreakdownTime { get; set; }
        public int MachineID { get; set; }
        public string Shift { get; set; }
        public string CorrectedDate { get; set; }
    
        public virtual tblmachinedetail tblmachinedetail { get; set; }
    }
}
