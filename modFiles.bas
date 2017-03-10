Attribute VB_Name = "modFiles"
'When working on the source in a virtual machine, VB6 cannot ChDir to Shared Folders (UNC, Network Drives) https://msdn.microsoft.com/en-us/library/aa263345(v=vs.60).aspx
'These functions build paths based on the App.Path which is still valid

Function GetGfxPath()
    GetGfxPath = App.Path + "\" + GFXPATH
End Function

Function GetSoundPath()
    GetSoundPath = App.Path + "\" + SOUNDPATH
End Function

Function GetMusicPath()
    GetMusicPath = App.Path + "\" + MUSICPATH
End Function
