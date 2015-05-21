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

namespace DirectXCaptureWrapper.Capture
{
  public class DeviceInUseException : SystemException
  {
    // Initializes a new instance with the specified HRESULT
    public DeviceInUseException(string deviceName, int hResult)
      : base(deviceName + " is in use or cannot be rendered. (" + hResult + ")")
    {
    }
  }
}
