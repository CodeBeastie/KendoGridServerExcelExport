using System;
using System.Linq;
using Kendo.Mvc;
using Kendo.Mvc.UI;
using System.Collections.Generic;
using System.Collections;
using Kendo.Mvc.Infrastructure;

namespace TelerikMvcApp1.Web {
	public class KendoHelper {
		const string BLANKTEXT = "NULL";

		/// <summary>
		/// Handle blank searches in the kendo grid filters
		/// </summary>
		/// <param name="request"></param>
		public static void BlankSearch(DataSourceRequest request) {
			if (request.Filters != null) {
				for (int i = 0; i < request.Filters.Count; i++) {
					if (request.Filters[i] is FilterDescriptor) {
						BlankSearchFilterDescriptor(request.Filters[i] as FilterDescriptor);
					} else if (request.Filters[i] is CompositeFilterDescriptor) {
						BlankSearchCompositeFilterDescriptor(request.Filters[i] as CompositeFilterDescriptor);
					}
				}
			}
		}

		public static void BlankSearchFilterDescriptor(FilterDescriptor fd) {
			if (fd.Value != null) {
				if (fd.Value.ToString() == BLANKTEXT) {
					fd.Value = "";
				}
			}
		}

		public static void BlankSearchCompositeFilterDescriptor(CompositeFilterDescriptor cfd) {
			if (cfd == null) return;
			for (int i = 0; i < cfd.FilterDescriptors.Count(); i++) {
				if (cfd.FilterDescriptors[i] is FilterDescriptor) {
					BlankSearchFilterDescriptor(cfd.FilterDescriptors[i] as FilterDescriptor);
				} else if (cfd.FilterDescriptors[i] is CompositeFilterDescriptor) {
					BlankSearchCompositeFilterDescriptor(cfd.FilterDescriptors[i] as CompositeFilterDescriptor);
				}
			}
		}

		/// <summary>
		/// Extract the list of data (CB viewmodel's) from the KEndo data source result.
		/// Handles the grouping being enabled.
		/// </summary>
		/// <param name="res"></param>
		/// <returns></returns>
		public static IEnumerable GetDataList(DataSourceResult res) {
			IEnumerable data = null;

			if (res == null) return null;
			if (res.Data == null) return null;

			//check if grouping is in use and if so the data is deeper in the object.
			if (res.Data is IEnumerable<Kendo.Mvc.Infrastructure.AggregateFunctionsGroup>) {
				var en = res.Data.GetEnumerator();
				if (en != null) {
					if (en.MoveNext()) {
						AggregateFunctionsGroup m = en.Current as AggregateFunctionsGroup;
						if (m != null) {
							data = m.Items;
						}
					}
				}

			} else {
				data = res.Data;
			}

			return data;
		}


	}
}