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
using DirectShowWrapper;

namespace DirectXCaptureWrapper.Capture
{
  public class Filters
  {
    public FilterCollection VideoInputDevices = new FilterCollection(FilterCategory.VideoInputDevice);
    public FilterCollection AudioInputDevices = new FilterCollection(FilterCategory.AudioInputDevice);
    public FilterCollection VideoCompressors = new FilterCollection(FilterCategory.VideoCompressorCategory);
    public FilterCollection AudioCompressors = new FilterCollection(FilterCategory.AudioCompressorCategory);
  }
}
