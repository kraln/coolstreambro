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

  public class Source : IDisposable
  {

    protected string name;				// Name of the source

    public string Name
    {
      get
      {
        return (name);
      }
    }

    public override string ToString()
    {
      return (Name);
    }


    public virtual bool Enabled
    {
      get
      {
        throw new NotSupportedException("This method should be overriden in derrived classes.");
      }
      set
      {
        throw new NotSupportedException("This method should be overriden in derrived classes.");
      }
    }

    ~Source()
    {
      Dispose();
    }



    public virtual void Dispose()
    {
      name = null;
    }

  }
}
