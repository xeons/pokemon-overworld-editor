Pokémon Overworld Sprite Editor
===============================

Sprite Navigation
-----------------
This is where you can pick the index of the sprite
to edit, there are over 200 sprites on all the games.

Right now there is no way of determining how many
frames each sprite has, so if you go over how many
frames are actually in a sprite, it will start reading
from the sprite # after it.

Sprite Header #1 Info
---------------------
Sprite #: Current sprite you are on.
Starter Bytes: First two bytes in the header, always FFFF.
Pallete #: This is just the numerical index of the Pallete. There are around 20+ sprite palletes in the game.
Unknown Data: Self-Explanitory, its in the header, and I have no idea what it does?
Sprite Data Size: The number of bytes a the sprite takes
Width and Height: Self-explanitory.
Unknown Data 2: More data I have no clue what it does.
Unknown Pointer 1, 2, 3, and 4 I have no idea what kind of data these pointers point to. I believe 2 has to do with tile arrangment.
Sprite Pointer: Points to Sprite Header #2

Sprite Header #2 Info
---------------------
Sprite Pointer: Actual pointer to the sprite image data.
Data Size: another specifier of datasize
Unknown 1: As the name says!


Drawing Canvas Features
-----------------------
- You can drag the current color, its like paintbrush.
- Right click on a pixel and it will make that the selected
  color.
- Displays the color your mouse is over, and the currently
  selected color.

Games Supported
===============
- Pokemon Ruby/Sapphire/Emerald (English Versions)
- Pokemon FireRed/LeafGreen (English and Japanese Versions)

Features Coming Soon
====================
- Import/Export Bitmaps
- Undo/Redo
- German and More japanese ROM support.
- Repointing and Pallete Editing