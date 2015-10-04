 local genx = require("genx")
local ffi = require("ffi")
local minizip = require("minizip")
local zlib = require("zlib")

local odFuncs = {}

function odFuncs:save(filename)
	local zip = minizip.open(filename,"w")
	local xml = genx.new()
	local function writeToZip(s)
		if s then
			zip:write(ffi.string(s))
		end
	end
	-- Write the mimetype
	local mime = "application/vnd.oasis.opendocument.spreadsheet"
	zip:add_file({['filename'] = "mimetype",raw=true,method=0})
	zip:write(mime)
	zip:close_file_raw(#mime,zlib.crc32(mime))

	-- Write meta.xml
	zip:add_file("meta.xml")
	zip:write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
	xml:start_doc(writeToZip)
	local dcNS = xml:ns("http://purl.org/dc/elements/1.1/", "dc")
	local officeNS = xml:ns("urn:oasis:names:tc:opendocument:xmlns:office:1.0", "office")
	local metaNS = xml:ns("urn:oasis:names:tc:opendocument:xmlns:meta:1.0", "meta")
	local styleNS = xml:ns("urn:oasis:names:tc:opendocument:xmlns:style:1.0", "style")
	local svgNS = xml:ns("urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0", "svg")
	local tableNS = xml:ns("urn:oasis:names:tc:opendocument:xmlns:table:1.0", "table")
	local textNS = xml:ns("urn:oasis:names:tc:opendocument:xmlns:text:1.0","text")
	local configNS = xml:ns("urn:oasis:names:tc:opendocument:xmlns:config:1.0")
	xml:start_element("document-meta", officeNS)
	xml:start_element("meta", officeNS)
	xml:start_element("editing-duration", metaNS)
	xml:text("PT0H0M0S")
	xml:end_element()
	xml:start_element("editing-cycles", metaNS)
	xml:text("1")
	xml:end_element()
	xml:start_element("document-statistic", metaNS)
	xml:end_element()
	xml:start_element("generator", metaNS)
	xml:text("lj-ods")
	xml:end_element()
	xml:end_element()
	xml:end_element()
	xml:end_doc()
	zip:close_file()

	-- Write content.xml
	zip:add_file("content.xml")
	zip:write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
	xml:start_doc(writeToZip)
	xml:start_element("document-content", officeNS)
	xml:add_attr("version", "1.1", officeNS)
	xml:start_element("scripts", officeNS)
	xml:end_element()
	xml:start_element("font-face-decls", officeNS)
	xml:start_element("font-face", styleNS)
	xml:add_attr("name", "Liberation Sans", styleNS)
	xml:add_attr("font-family", "'Liberation Sans'", svgNS)
	xml:add_attr("font-family-generic", "swiss", styleNS)
	xml:add_attr("font-pitch", "variable", styleNS)
	xml:end_element()
	xml:end_element()
	xml:start_element("automatic-styles", officeNS)
	xml:end_element()
	xml:start_element("body", officeNS)
	xml:start_element("spreadsheet", officeNS)
	xml:start_element("table", tableNS)
	xml:add_attr("name", "Sheet1", tableNS)
	xml:start_element("table-column", tableNS)
	xml:add_attr("number-columns-repeated", "1", tableNS)
	xml:end_element()
	local lastK = 0
	for k,v in pairs(self.spreadsheet) do
		-- Fill up empty rows
		local distance = k-lastK-1
		if distance > 0 then
			xml:start_element("table-row",tableNS)
			xml:add_attr("number-rows-repeated", tostring(distance), tableNS)
			xml:end_element()
		end
		lastK = k
		xml:start_element("table-row", tableNS)
		local sortedColumns = {}
		for a,b in pairs(v) do
			sortedColumns[#sortedColumns+1] = {a,b}
		end
		table.sort(sortedColumns, function(lh,rh) return lh[1] < rh[1] end)
		local lastColumn = 0
		for a,b in ipairs(sortedColumns) do
			-- Fill up empty columns
			local distance = b[1]-lastColumn-1
			if distance > 0 then
				xml:start_element("table-cell", tableNS)
				xml:add_attr("number-columns-repeated", tostring(distance), tableNS)
				xml:end_element()
			end
			lastColumn = b[1]
			xml:start_element("table-cell", tableNS)
			-- TODO: Fix...
--			if type(b) == "number" then
--				xml:add_attr("value-type", "float", tableNS)
--				xml:add_attr("value", tostring(b[2]), officeNS)
--			else
				xml:add_attr("value-type", "string", tableNS)
--				xml:add_attr("string-value", tostring(b[2]), officeNS)
--			end
			xml:start_element("p", textNS)
			xml:text(tostring(b[2]))
			xml:end_element()
			xml:end_element()
		end
		xml:end_element()
	end
	xml:end_element()
	xml:end_element()
	xml:end_element()
	xml:end_element()
	xml:end_doc()
	zip:close_file()

	zip:add_file("styles.xml")
	xml:start_doc(writeToZip)
	xml:start_element("document-styles", officeNS)
	xml:start_element("font-face-decls", officeNS)
	xml:start_element("font-face", styleNS)
	xml:add_attr("name", "Arial", styleNS)
	xml:add_attr("font-family", "Arial", svgNS)
	xml:end_element()
	xml:end_element()
	xml:end_element()
	xml:end_doc()
	zip:close_file()

	zip:add_file("settings.xml")
	xml:start_doc(writeToZip)
	xml:start_element("settings", officeNS)
	xml:end_element()
	xml:end_doc()
	zip:close_file()

	-- Write manifest file last...
	zip:add_file("META-INF/")
	zip:close_file()
	zip:add_file("META-INF/manifest.xml")

	xml:start_doc(writeToZip)
	local manifestNS = xml:ns("urn:oasis:names:tc:opendocument:xmlns:manifest:1.0", "manifest")
	xml:start_element("manifest", manifestNS)
	xml:add_attr("version", "1.1", manifestNS)
	xml:start_element("file-entry", manifestNS)
	xml:add_attr("full-path", "/", manifestNS)
	xml:add_attr("version", "1.1", manifestNS)
	xml:add_attr("media-type", "application/vnd.oasis.opendocument.spreadsheet", manifestNS)
	xml:end_element()
	for k,v in pairs({'meta.xml','content.xml','styles.xml','settings.xml'}) do
		xml:start_element("file-entry", manifestNS)
		xml:add_attr("full-path", v, manifestNS)
		xml:add_attr("media-type", "text/xml", manifestNS)
		xml:end_element()
	end
	xml:end_element()
	xml:end_doc()
	xml:free()
	zip:close_file()
	zip:close()
end

local function parseColumn(column)
	local columnChars = {string.byte(column,1,#column)}
	local column = 1
	for k,v in pairs(columnChars) do
		column = column+(v-65)+((k-1)*26)
	end
	return column
end

local function new()
	return setmetatable({spreadsheet = {}}, {
		__newindex = function(self,k,v)
			local column,row = string.match(k,"([^%d]+)(%d+)")
			row = tonumber(row)
			column = parseColumn(column)
			if not self.spreadsheet[row] then
				self.spreadsheet[row] = {}
			end
			self.spreadsheet[row][column] = v
		end,
		__index = function(self,k)
			local column,row = string.match(k, "([^%d]+)(%d+)")
			if not (column and row) then
				return odFuncs[k]
			end
			row = tonumber(row)
			column = parseColumn(column)
			if not self.spreadsheet[row] then return nil end
			return self.spreadsheet[row][column]
		end
	})
end
return new
