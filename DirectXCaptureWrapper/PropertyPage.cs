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
  public class PropertyPage : IDisposable, DirectXCaptureWrapper.Capture.IPropertyPage
  {

    public string Name;
    public bool SupportsPersisting = false;

    public virtual byte[] State
    {
      get
      {
        throw new NotSupportedException("This property page does not support persisting state.");
      }
      set
      {
        throw new NotSupportedException("This property page does not support persisting state.");
      }

    }
    public virtual void Show(Control owner)
    {
      throw new NotSupportedException("Not implemented. Use a derived class. ");
    }

    public void Dispose()
    {
      throw new NotImplementedException();
    }
  }
}
