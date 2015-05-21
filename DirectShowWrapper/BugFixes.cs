/*
 * Cool Stream, Bro
 * 
 * DirectShowWrapper
 * 
 * This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
 * If a copy of the MPL was not distributed with this file, 
 * You can obtain one at http://mozilla.org/MPL/2.0/. 
 */

using System;
using System.Text;
using System.Runtime.InteropServices;

namespace DirectShowWrapper
{

  public class BugFixes
  {
    /*
     * CoCreateInstance( CLSID_CaptureGraphBuilder2, ..., IID_IUnknown, ...); 
     * 
     * Fails with an E_NOTIMPL, so we do this instead:
     * 
     * CoCreateInstance( CLSID_CaptureGraphBuilder2, ..., IID_ICaptureGraphBuilder2, ...);
     * 
     * Sigh.
     */

    public static object InstantiateDirectShow(ref Guid clsid, ref Guid riid)
    {
      IntPtr ptrIf;
      int hr = CoCreateInstance(ref clsid, IntPtr.Zero, CLSCTX.Inproc, ref riid, out ptrIf);
      if ((hr != 0) || (ptrIf == IntPtr.Zero))
      {
        Marshal.ThrowExceptionForHR(hr);
      }

      Guid iu = new Guid("00000000-0000-0000-C000-000000000046");
      IntPtr ptrXX;
      hr = Marshal.QueryInterface(ptrIf, ref iu, out ptrXX);

      object ooo = System.Runtime.Remoting.Services.EnterpriseServicesHelper.WrapIUnknownWithComObject(ptrIf);
      int ct = Marshal.Release(ptrIf);
      return ooo;
    }

    [DllImport("ole32.dll")]
    private static extern int CoCreateInstance(ref Guid clsid, IntPtr pUnkOuter, CLSCTX dwClsContext, ref Guid iid, out IntPtr ptrIf);
  }

  [Flags]
  internal enum CLSCTX
  {
    Inproc = 0x03,
    Server = 0x15,
    All = 0x17,
  }
}
