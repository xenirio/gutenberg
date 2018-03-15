using System;
using System.Collections.Generic;
using System.Text;

namespace Handshakes.Api.Report.Models
{
    public abstract class ReportElement
    {
		private string _key;
		public string Key {
			get { return String.Format("{{{0}}}", _key); }
			set { _key = value; }
		}
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
