using System;
namespace DirectXCaptureWrapper.Capture
{
  interface IPropertyPage
  {
    void Show(System.Windows.Forms.Control owner);
    byte[] State
    {
      get;
      set;
    }
  }
}
