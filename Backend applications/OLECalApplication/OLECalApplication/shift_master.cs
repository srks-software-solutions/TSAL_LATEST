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
    
    public partial class shift_master
    {
        public int ShiftID { get; set; }
        public string ShiftName { get; set; }
        public System.TimeSpan StartTime { get; set; }
        public System.TimeSpan EndTime { get; set; }
        public System.DateTime InsertedOn { get; set; }
        public string InsertedBy { get; set; }
        public Nullable<System.DateTime> ModifiedOn { get; set; }
        public string ModifiedBy { get; set; }
        public int IsDeleted { get; set; }
        public Nullable<int> Duration { get; set; }
    }
}
