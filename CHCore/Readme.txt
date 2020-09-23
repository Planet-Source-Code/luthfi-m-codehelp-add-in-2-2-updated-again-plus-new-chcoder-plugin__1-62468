Version 2.2


To compile the project you need to have WinAPIForVB type library already registered 
in your machine. 
You can find the typelib at 
http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=62060&lngWId=1 
(thanks MH)

You can also download a convenient setup file from my forum at
http://codehelp.cjb.net

IMPORTANT!!!
If you decide to compile the add in yourself, please follow these steps:

- Register CHLib.tlb found in Interfaces sub folder
- Register WinAPIForVB.tlb
- Build CHGlobal.vbp
- Build CHCore.vbp
- Build all the vbp in plugins sub folder, place the compiled dll of each plugin in "plugins"
  sub folder where the CHCore.dll resides.

Note:
If you got warning message saying can not find compatible dll just ignore it and set binary
compatibility to no compatibility, after you build each dll, set it back to binary compatibility
and set it to your newly built dll.


Important changes since ver 2.0:
- Moved interface definition from VB.dll to typelib, so the plugin GUID now truly constant.
- Added new plugin (CHTabIdx.dll)
- Added new ShowHelp, ShowPropertyDialog interface
- Added Enable/disable plugin at runtime
- Added features to existing plugin (grouped panels, keyboard shortcut for activating visible tab
  with on-screen hint
- bug fix


Thanks,

Luthfi
