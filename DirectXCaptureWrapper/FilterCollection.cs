/*
 * Cool Stream, Bro
 * 
 * DirectXCaptureWrapper
 * 
 * This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
 * If a copy of the MPL was not distributed with this file, 
 * You can obtain one at http://mozilla.org/MPL/2.0/. 
 */

using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using DirectShowWrapper;
using DirectShowWrapper.Device;

namespace DirectXCaptureWrapper.Capture
{

  public class FilterCollection : CollectionBase
  {
    /// <summary> Populate the collection with a list of filters from a particular category. </summary>
    internal FilterCollection(Guid category)
    {
      getFilters(category);
    }

    /// <summary> Populate the InnerList with a list of filters from a particular category </summary>
    protected void getFilters(Guid category)
    {
      int hr;
      object comObj = null;
      ICreateDevEnum enumDev = null;
      IEnumMoniker enumMon = null;
      IMoniker[] mon = new IMoniker[1];

      try
      {
        // Get the system device enumerator
        Type srvType = Type.GetTypeFromCLSID(Clsid.SystemDeviceEnum);
        if (srvType == null)
          throw new NotImplementedException("System Device Enumerator");
        comObj = Activator.CreateInstance(srvType);
        enumDev = (ICreateDevEnum)comObj;

        // Create an enumerator to find filters in category
        hr = enumDev.CreateClassEnumerator(ref category, out enumMon, 0);
        if (hr != 0)
          throw new NotSupportedException("No devices of the category");

        // Loop through the enumerator
        do
        {
          // Next filter
          hr = enumMon.Next(1, mon, IntPtr.Zero);
          if ((hr != 0) || (mon[0] == null))
            break;

          // Add the filter
          Filter filter = new Filter(mon[0]);
          InnerList.Add(filter);

          // Release resources
          Marshal.ReleaseComObject(mon[0]);
          mon[0] = null;
        }
        while (true);

        // Sort
        InnerList.Sort();
      }
      finally
      {
        enumDev = null;
        if (mon[0] != null)
          Marshal.ReleaseComObject(mon[0]);
        mon[0] = null;
        if (enumMon != null)
          Marshal.ReleaseComObject(enumMon);
        enumMon = null;
        if (comObj != null)
          Marshal.ReleaseComObject(comObj);
        comObj = null;
      }
    }

    /// <summary> Get the filter at the specified index. </summary>
    public Filter this[int index]
    {
      get
      {
        return ((Filter)InnerList[index]);
      }
    }
  }
}
