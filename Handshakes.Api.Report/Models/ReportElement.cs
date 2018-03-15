using System;
using System.Collections.Generic;
using System.Text;

namespace Handshakes.Api.Report.Models
{
    public abstract class ReportElement
    {
		public string Key { get; set; }

		protected object value { get; set; }

		protected T GetValue<T>()
		{
			return (T)value;
		}

		protected void SetValue<T>(T value)
		{
			this.value = value;
		}
	}
}
