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

  public class VfwCompressorPropertyPage : PropertyPage
  {
    protected IAMVfwCompressDialogs vfwCompressDialogs = null;

    public override byte[] State
    {
      get
      {
        byte[] data = null;
        int size = 0;

        int hr = vfwCompressDialogs.GetState(null, ref size);
        if ((hr == 0) && (size > 0))
        {
          data = new byte[size];
          hr = vfwCompressDialogs.GetState(data, ref size);
          if (hr != 0)
            data = null;
        }
        return (data);
      }
      set
      {
        int hr = vfwCompressDialogs.SetState(value, value.Length);
        if (hr != 0)
          Marshal.ThrowExceptionForHR(hr);
      }
    }
    public VfwCompressorPropertyPage(string name, IAMVfwCompressDialogs compressDialogs)
    {
      Name = name;
      SupportsPersisting = true;
      this.vfwCompressDialogs = compressDialogs;
    }

    public override void Show(Control owner)
    {
      vfwCompressDialogs.ShowDialog(VfwCompressDialogs.Config, owner.Handle);
    }

  }
}
