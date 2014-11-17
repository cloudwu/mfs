local zip = require "zip"
local base64 = require "base64"

------------- convert xlsx to mfs

local function x2m(filename, mfs)

	assert(filename, "need filename")
	assert(mfs, "need mfs filename")

	local xml_ext = {}

	for _, v in ipairs { "xml", "rels", "vml" } do
		xml_ext[v] = true
	end

	local function is_xml(filename)
		return xml_ext[filename:match("%.(%a+)$")]
	end

	local function xml_split(s)
		s = s:gsub("[\n\r]","")
		local tbl = {}
		for v in s:gmatch "<[^>]*>[^<]*" do
			if v:find "^</" and tbl[#tbl]:sub(-1) ~= ">" then
				tbl[#tbl] = tbl[#tbl] .. v
			else
				table.insert(tbl, v)
			end
		end
		return tbl
	end

	local xml_remove = {}
	for _, v in ipairs { "bookViews" , "sheetViews" } do
		xml_remove["<"..v..">"] = "</"..v..">"
	end

	local function xml_filter(s)
		local closetag
		for i, line in ipairs(s) do
			if line:find "^<dcterms:modified" then
				s[i] = ""
			elseif closetag then
				s[i] = ""
				if line == closetag then
					closetag = nil
				end
			else
				closetag = xml_remove[line]
				if closetag then
					s[i] = ""
				end
			end
		end
		return s
	end

	local u = zip.unzip(filename)
	local list = u:list()
	table.sort(list)

	local output = io.open(mfs, "wb")

	for _, filename in ipairs(list) do
		local blob = u:read(filename)
		if is_xml(filename) then
			output:write("*" .. filename .. "\n")
			local xml = xml_split(blob)
			xml_filter(xml)
			for _, v in ipairs(xml) do
				if v ~= "" then
					output:write(v)
					output:write("\n")
				end
			end
		else
			output:write("+" .. filename .. "\n")
			output:write(base64.encode(blob))
			output:write("\n")
		end
	end
	u:close()
	output:close()
end

--------------------Convert mfs to xlsx

local function m2x(mfs, xlsx)
	assert(mfs, "Need mfs filename")
	assert(xlsx, "Need xlsx filename")

	local source = assert(io.open(mfs, "rb") ,"Can't open " .. mfs)
	local zipfile = assert(zip.zip(xlsx), "Can't write to " .. xlsx)

	local content
	local blob
	local filename

	for line in source:lines() do
		local t = line:sub(1,1)
		if t == "*" or t == "+" then
			if content then
				zipfile:write(filename, table.concat(content))
			end
			filename = line:sub(2)
			if t == "*" then
				content = {}
			else
				blob = true
				content = nil
			end
		elseif blob then
			zipfile:write(filename, base64.decode(line))
			blob = nil
		else
			assert(t == "<")
			table.insert(content, line)
		end
	end

	if blob then
		zipfile:write(filename, base64.decode(line))
	elseif content then
		zipfile:write(filename, table.concat(content))
	end

	source:close()
	zipfile:close()
end

--------------------------------------------------

x2m(...)
