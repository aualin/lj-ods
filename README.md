# lj-ods
Barebones OpenDocument Spreadsheet generator.
Done with some light spec reading, and heavy
libreoffice generated sheets reading. Goal is to
make a 1.1 file that opens without errors in
Office 2007 and ODF addin for Office 2003.
Not quite there yet...
Depends on [genx](https://www.github.com/luapower/genx) and [minizip](https://www.github.com/luapower/minizip)
from [luapower](https://www.luapower.com/)

Example:
```lua
local ods = require("ods")
local sheet = ods()
sheet['A1'] = "Hello"
sheet['B1'] = "world!"
sheet:save("test.ods")
```
