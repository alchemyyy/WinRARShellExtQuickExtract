WinRAR Dual Context Menu
========================

This Windows 11 shell extension adds a top level context menu item for WinRAR's "Extract to <foldername>".

This differs from the "legacy" context menu setting in WinRAR since this just adds the one thing I use 99.9% of the time. I don't like how cluttered the menu gets with every option I have enabled put at the top level menu.

But I wanted get the bonus of having WinRAR's current complete cascading context menu as I occasionally use some of those other functions.

Notes:
* The dll goes into WinRAR's program folder
* The "supported file types" are grabbed from WinRAR's registry *at launch*. So if you want this to update, you have to restart explorer.
* The positioning in the context menu is about as good as it is gonna get. Can't go higher without registering it as a "verb", which means no dynamic entry naming. I prefer having the output folder name visible for the extra context clue over moving the entry up a couple slots. I also haven't investigated moving WinRAR's own menu down to be with it yet.