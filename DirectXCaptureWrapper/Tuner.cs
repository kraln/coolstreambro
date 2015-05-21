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
using DirectShowWrapper;

namespace DirectXCaptureWrapper.Capture
{

  public enum TunerInputType
  {
    Cable,
    Antenna
  }


  public class Tuner : IDisposable
  {
    protected IAMTVTuner tvTuner = null;

    public Tuner(IAMTVTuner tuner)
    {
      tvTuner = tuner;
    }


    public int Channel
    {
      get
      {
        int channel;
        int v, a;
        int hr = tvTuner.get_Channel(out channel, out v, out a);
        if (hr != 0)
          Marshal.ThrowExceptionForHR(hr);
        return (channel);
      }

      set
      {
        int hr = tvTuner.put_Channel(value, AMTunerSubChannel.Default, AMTunerSubChannel.Default);
        if (hr != 0)
          Marshal.ThrowExceptionForHR(hr);
      }
    }

    public TunerInputType InputType
    {
      get
      {
        DirectShowWrapper.TunerInputType t;
        int hr = tvTuner.get_InputType(0, out t);
        if (hr != 0)
          Marshal.ThrowExceptionForHR(hr);
        return ((TunerInputType)t);
      }
      set
      {
        DirectShowWrapper.TunerInputType t = (DirectShowWrapper.TunerInputType)value;
        int hr = tvTuner.put_InputType(0, t);
        if (hr != 0)
          Marshal.ThrowExceptionForHR(hr);
      }
    }


    public bool SignalPresent
    {
      get
      {
        AMTunerSignalStrength sig;
        int hr = tvTuner.SignalPresent(out sig);
        if (hr != 0)
          Marshal.ThrowExceptionForHR(hr);
        if (sig == AMTunerSignalStrength.NA)
          throw new NotSupportedException("Signal strength not available.");
        return (sig == AMTunerSignalStrength.SignalPresent);
      }
    }

    public void Dispose()
    {
      if (tvTuner != null)
        Marshal.ReleaseComObject(tvTuner);
      tvTuner = null;
    }
  }
}
