using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

internal static class WordOleObjectAccessor
{
    public static object GetRunningObject(OLEFormat format)
    {
        if (format is null) throw new ArgumentNullException(nameof(format));
        try
        {
            var runningObject = format.Object;
            if (runningObject is not null) return runningObject;
        }
        catch (Exception error) when (error is COMException or InvalidCastException)
        {
            // A freshly pasted embedded OLE object can be dormant until Word
            // activates it. Fall through to the same activation path below.
        }

        object showVerb = (int)WdOLEVerb.wdOLEVerbShow;
        format.DoVerb(ref showVerb);
        return format.Object
            ?? throw new COMException(
                "Word activated the VisualTeX OLE object but did not expose its running COM object.");
    }
}
