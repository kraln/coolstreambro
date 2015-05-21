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
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using DirectShowWrapper;

namespace DirectXCaptureWrapper.Capture
{
  public class VideoCapabilities
  {


    public Size InputSize;

    public Size MinFrameSize;

    public Size MaxFrameSize;

    public int FrameSizeGranularityX;

    public int FrameSizeGranularityY;

    public double MinFrameRate;

    public double MaxFrameRate;



    internal VideoCapabilities(IAMStreamConfig videoStreamConfig)
    {
      if (videoStreamConfig == null)
        throw new ArgumentNullException("videoStreamConfig");

      AMMediaType mediaType = null;
      VideoStreamConfigCaps caps = null;
      IntPtr pCaps = IntPtr.Zero;
      IntPtr pMediaType;
      try
      {
        // Ensure this device reports capabilities
        int c, size;
        int hr = videoStreamConfig.GetNumberOfCapabilities(out c, out size);
        if (hr != 0)
          Marshal.ThrowExceptionForHR(hr);
        if (c <= 0)
          throw new NotSupportedException("This video device does not report capabilities.");
        if (size > Marshal.SizeOf(typeof(VideoStreamConfigCaps)))
          throw new NotSupportedException("Unable to retrieve video device capabilities. This video device requires a larger VideoStreamConfigCaps structure.");
        if (c > 1)
          Debug.WriteLine("This video device supports " + c + " capability structures. Only the first structure will be used.");

        // Alloc memory for structure
        pCaps = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(VideoStreamConfigCaps)));

        // Retrieve first (and hopefully only) capabilities struct
        // XXX Maybe there is more than one?
        hr = videoStreamConfig.GetStreamCaps(0, out pMediaType, pCaps);
        if (hr != 0)
          Marshal.ThrowExceptionForHR(hr);

        // Convert pointers to managed structures
        mediaType = (AMMediaType)Marshal.PtrToStructure(pMediaType, typeof(AMMediaType));
        caps = (VideoStreamConfigCaps)Marshal.PtrToStructure(pCaps, typeof(VideoStreamConfigCaps));

        // Extract info
        InputSize = caps.InputSize;
        MinFrameSize = caps.MinOutputSize;
        MaxFrameSize = caps.MaxOutputSize;
        FrameSizeGranularityX = caps.OutputGranularityX;
        FrameSizeGranularityY = caps.OutputGranularityY;
        MinFrameRate = (double)10000000 / caps.MaxFrameInterval;
        MaxFrameRate = (double)10000000 / caps.MinFrameInterval;
      }
      finally
      {
        if (pCaps != IntPtr.Zero)
          Marshal.FreeCoTaskMem(pCaps);
        pCaps = IntPtr.Zero;
        if (mediaType != null)
          CommonUtils.FreeAMMediaType(mediaType);
        mediaType = null;
      }
    }
  }
}
