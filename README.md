#Cool Stream, Bro

I was watching PJ's stream, and he had suddenly a lot of problems with Amarec.
I figured I could throw something together quickly that can record audio/video from inputs (and preview it simultaneously).

So I did.

Cool story, bro. 

Cheers,
Kraln

##Roughness

* You'll need the .NET framework 4.5 ( get it here http://go.microsoft.com/fwlink/p/?LinkId=245484 )
* If you try to write to a file where you don't have permission, it will fail in ugly ways (try your documents folder)
* It doesn't have a proper file picker dialog
* It doesn't remember any of your settings
* It probably won't crash? 

Any problems? Hatred? Let me know coolstreambro@kraln.com

## Todo
 * Support A/V Crossbars (Capture Cards with Audio Intermux'd)
 * Support rebroadcasting

## Acknowledgements
This is based on a lot of work that Brian Low did in 2003 and released into the public domain. I cleaned up a lot of it, and modernized it, but at its core it's still very strongly influenced by Brian and the DirectX SDK examples. Hopefully the MPL2 license keeps this very useful code open.
