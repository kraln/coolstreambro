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
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DirectShowWrapper;

namespace DirectXCaptureWrapper.Capture
{
  public class DirectShowPropertyPage : PropertyPage
  {
    protected ISpecifyPropertyPages specifyPropertyPages;

    public DirectShowPropertyPage(string name, ISpecifyPropertyPages specifyPropertyPages)
    {
      Name = name;
      SupportsPersisting = false;
      this.specifyPropertyPages = specifyPropertyPages;
    }

    public override void Show(Control owner)
    {
      DsCAUUID cauuid = new DsCAUUID();
      try
      {
        int hr = specifyPropertyPages.GetPages(out cauuid);
        if (hr != 0)
          Marshal.ThrowExceptionForHR(hr);

        object o = specifyPropertyPages;
        hr = OleCreatePropertyFrame(owner.Handle, 30, 30, null, 1,
          ref o, cauuid.cElems, cauuid.pElems, 0, 0, IntPtr.Zero);
      }
      finally
      {
        if (cauuid.pElems != IntPtr.Zero)
          Marshal.FreeCoTaskMem(cauuid.pElems);
      }
    }

    public new void Dispose()
    {
      if (specifyPropertyPages != null)
        Marshal.ReleaseComObject(specifyPropertyPages);
      specifyPropertyPages = null;
    }

    /* This DLL sometimes isn't there, should we ship it? */
    [DllImport("olepro32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    private static extern int OleCreatePropertyFrame(
      IntPtr hwndOwner, int x, int y,
      string lpszCaption, int cObjects,
      [In, MarshalAs(UnmanagedType.Interface)] ref object ppUnk,
      int cPages, IntPtr pPageClsID, int lcid, int dwReserved, IntPtr pvReserved);

  }
}
