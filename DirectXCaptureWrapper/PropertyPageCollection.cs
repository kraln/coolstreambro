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
using System.Collections;
using System.Runtime.InteropServices;
using DirectShowWrapper;

namespace DirectXCaptureWrapper.Capture
{

  public class PropertyPageCollection : CollectionBase, IDisposable
  {

    internal PropertyPageCollection()
    {
      InnerList.Capacity = 1;
    }

    internal PropertyPageCollection(
      ICaptureGraphBuilder2 graphBuilder,
      IBaseFilter videoDeviceFilter, IBaseFilter audioDeviceFilter,
      IBaseFilter videoCompressorFilter, IBaseFilter audioCompressorFilter,
      SourceCollection videoSources, SourceCollection audioSources)
    {
      addFromGraph(graphBuilder,
        videoDeviceFilter, audioDeviceFilter,
        videoCompressorFilter, audioCompressorFilter,
        videoSources, audioSources);
    }

    ~PropertyPageCollection()
    {
      Dispose();
    }

    public new void Clear()
    {
      for (int c = 0; c < InnerList.Count; c++)
        this[c].Dispose();
      InnerList.Clear();
    }

    public void Dispose()
    {
      Clear();
      InnerList.Capacity = 1;
    }


    public PropertyPage this[int index]
    {
      get
      {
        return ((PropertyPage)InnerList[index]);
      }
    }

    protected void addFromGraph(
      ICaptureGraphBuilder2 graphBuilder,
      IBaseFilter videoDeviceFilter, IBaseFilter audioDeviceFilter,
      IBaseFilter videoCompressorFilter, IBaseFilter audioCompressorFilter,
      SourceCollection videoSources, SourceCollection audioSources)
    {
      object filter = null;
      Guid cat;
      Guid med;
      Guid iid;
      int hr;

      Trace.Assert(graphBuilder != null);

      // 1. the video capture filter
      addIfSupported(videoDeviceFilter, "Video Capture Device");

      // 2. the video capture pin
      cat = PinCategory.Capture;
      med = MediaType.Interleaved;
      iid = typeof(IAMStreamConfig).GUID;
      hr = graphBuilder.FindInterface(
        ref cat, ref med, videoDeviceFilter, ref iid, out filter);
      if (hr != 0)
      {
        med = MediaType.Video;
        hr = graphBuilder.FindInterface(
          ref cat, ref med, videoDeviceFilter, ref iid, out filter);
        if (hr != 0)
          filter = null;
      }
      addIfSupported(filter, "Video Capture Pin");

      // 3. the video preview pin
      cat = PinCategory.Preview;
      med = MediaType.Interleaved;
      iid = typeof(IAMStreamConfig).GUID;
      hr = graphBuilder.FindInterface(
        ref cat, ref med, videoDeviceFilter, ref iid, out filter);
      if (hr != 0)
      {
        med = MediaType.Video;
        hr = graphBuilder.FindInterface(
          ref cat, ref med, videoDeviceFilter, ref iid, out filter);
        if (hr != 0)
          filter = null;
      }
      addIfSupported(filter, "Video Preview Pin");

      // 4. the video crossbar(s)
      ArrayList crossbars = new ArrayList();
      int num = 1;
      for (int c = 0; c < videoSources.Count; c++)
      {
        CrossbarSource s = videoSources[c] as CrossbarSource;
        if (s != null)
        {
          if (crossbars.IndexOf(s.Crossbar) < 0)
          {
            crossbars.Add(s.Crossbar);
            if (addIfSupported(s.Crossbar, "Video Crossbar " + (num == 1 ? "" : num.ToString())))
              num++;
          }
        }
      }
      crossbars.Clear();

      // 5. the video compressor
      addIfSupported(videoCompressorFilter, "Video Compressor");

      // 6. the video TV tuner
      cat = PinCategory.Capture;
      med = MediaType.Interleaved;
      iid = typeof(IAMTVTuner).GUID;
      hr = graphBuilder.FindInterface(
        ref cat, ref med, videoDeviceFilter, ref iid, out filter);
      if (hr != 0)
      {
        med = MediaType.Video;
        hr = graphBuilder.FindInterface(
          ref cat, ref med, videoDeviceFilter, ref iid, out filter);
        if (hr != 0)
          filter = null;
      }
      addIfSupported(filter, "TV Tuner");

      // 7. the video compressor (VFW)
      IAMVfwCompressDialogs compressDialog = videoCompressorFilter as IAMVfwCompressDialogs;
      if (compressDialog != null)
      {
        VfwCompressorPropertyPage page = new VfwCompressorPropertyPage("Video Compressor", compressDialog);
        InnerList.Add(page);
      }

      // 8. the audio capture filter
      addIfSupported(audioDeviceFilter, "Audio Capture Device");

      // 9. the audio capture pin
      cat = PinCategory.Capture;
      med = MediaType.Audio;
      iid = typeof(IAMStreamConfig).GUID;
      hr = graphBuilder.FindInterface(
        ref cat, ref med, audioDeviceFilter, ref iid, out filter);
      if (hr != 0)
      {
        filter = null;
      }
      addIfSupported(filter, "Audio Capture Pin");

      // 9. the audio preview pin
      cat = PinCategory.Preview;
      med = MediaType.Audio;
      iid = typeof(IAMStreamConfig).GUID;
      hr = graphBuilder.FindInterface(
        ref cat, ref med, audioDeviceFilter, ref iid, out filter);
      if (hr != 0)
      {
        filter = null;
      }
      addIfSupported(filter, "Audio Preview Pin");

      // 10. the audio crossbar(s)
      num = 1;
      for (int c = 0; c < audioSources.Count; c++)
      {
        CrossbarSource s = audioSources[c] as CrossbarSource;
        if (s != null)
        {
          if (crossbars.IndexOf(s.Crossbar) < 0)
          {
            crossbars.Add(s.Crossbar);
            if (addIfSupported(s.Crossbar, "Audio Crossbar " + (num == 1 ? "" : num.ToString())))
              num++;
          }
        }
      }
      crossbars.Clear();

      // 11. the audio compressor
      addIfSupported(audioCompressorFilter, "Audio Compressor");

    }

    protected bool addIfSupported(object o, string name)
    {
      ISpecifyPropertyPages specifyPropertyPages = null;
      DsCAUUID cauuid = new DsCAUUID();
      bool wasAdded = false;

      // Determine if the object supports the interface
      // and has at least 1 property page
      try
      {
        specifyPropertyPages = o as ISpecifyPropertyPages;
        if (specifyPropertyPages != null)
        {
          int hr = specifyPropertyPages.GetPages(out cauuid);
          if ((hr != 0) || (cauuid.cElems <= 0))
            specifyPropertyPages = null;
        }
      }
      finally
      {
        if (cauuid.pElems != IntPtr.Zero)
          Marshal.FreeCoTaskMem(cauuid.pElems);
      }

      // Add the page to the internal collection
      if (specifyPropertyPages != null)
      {
        DirectShowPropertyPage p = new DirectShowPropertyPage(name, specifyPropertyPages);
        InnerList.Add(p);
        wasAdded = true;
      }
      return (wasAdded);
    }
  }
}
