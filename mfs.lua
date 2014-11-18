local zip = require "zip"
local base64 = require "base64"
local xlsx = require "xlsxobj"

local function alert(err)
	local ok, winapi = pcall(require, "winapi")
	if ok then
		winapi.MessageBox(err)
	else
		print(err)
	end
end

------------- convert xlsx to mfs

local MFSVERSION = "MFS1"

local function x2m(filename, mfs)
	assert(filename, "need filename")
	mfs = mfs or filename:match("(.*%.)xlsx$") .. "mfs"
	assert(mfs, "need mfs filename")

	local xml_ext = {}

	for _, v in ipairs { "xml", "rels", "vml" } do
		xml_ext[v] = true
	end

	local function is_xml(filename)
		return xml_ext[filename:match("%.(%a+)$")]
	end

	local com_ext = {}

	for _,v in ipairs { "bin" } do
		com_ext[v] = true
	end

	local function is_com(filename)
		return com_ext[filename:match("%.(%a+)$")]
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
		xml_remove["<"..v] = "</"..v..">"
	end

	local xml_sort = {}
	for _, v in ipairs { "numFmts" , "Relationships", "cellStyles" } do
		xml_sort["<"..v] = "</"..v..">"
	end


	local function xml_filter(s)
		local closetag
		local filter_type
		local from
		for i, line in ipairs(s) do
			if line:find "^<dcterms:modified" then
				s[i] = ""
			elseif closetag then
				if filter_type == "REMOVE" then
					s[i] = ""
				end
				if line == closetag then
					if filter_type == "SORT" then
						if from+1 < i-1 then
							local tmp = { table.unpack(s, from+1, i-1) }
							table.sort(tmp)
							for j = from+1, i-1 do
								s[j] = tmp[j-from]
							end
						end
					end
					closetag = nil
					filter_type = nil
				end
			else
				line = line:match "^(<%w+)"
				closetag = xml_remove[line]
				if closetag then
					filter_type = "REMOVE"
					s[i] = ""
				else
					closetag = xml_sort[line]
					if closetag then
						filter_type = "SORT"
						from = i
					end
				end
			end
		end
		return s
	end

	local u = assert(zip.unzip(filename) , "Can't open "..filename)
	local list = u:list()
	table.sort(list)

	local output = assert(io.open(mfs, "wb") , "Can't write to " .. mfs)
	output:write(MFSVERSION.."\n")

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
			if is_com(filename) then
				blob = xlsx.bin(blob) or blob
			end
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
	xlsx = xlsx or mfs:match("(.*%.)mfs$") .. "xlsx"
	assert(xlsx, "Need xlsx filename")

	local source = assert(io.open(mfs, "rb") ,"Can't open " .. mfs)
	local zipfile = assert(zip.zip(xlsx), "Can't write to " .. xlsx)

	local content
	local blob
	local filename

	local version = source:read "*l"
	assert(version == MFSVERSION, "Unsupport mfs version")

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

local function log(...)
	print(os.date(), ...)
end

local function monitor(name)
	local winapi = require "winapi"
	local ti = winapi.LastWriteTime(name)
	coroutine.yield(ti)
	while true do
		winapi.Sleep(5000)	-- sleep 5 s
		local f = io.open(name, "r+")
		if f then
			f:close()
			return
		end

		local newt = winapi.LastWriteTime(name)
		if newt ~= ti then
			ti = newt
			coroutine.yield(true)
		end
	end
end

local M = {}

local TMP_PATH = os.getenv "mfs_tmp" or ""

function M.monitor(filename)
	local winapi = require "winapi"
	assert(filename, "Need an xlsx filename")
	local fullpath, name_noext, ext = filename:match("(.*)\\(.*%.)(%w+)$")
	assert(ext:lower() == "mfs", "Only support .mfs file")
	local tmpdir = base64.hex(base64.hashkey(filename))
	tmpdir = TMP_PATH .. tmpdir
	log("Create Dir", tmpdir)
	winapi.CreateDirectory(tmpdir)
	local tmpname = tmpdir .. "\\" .. name_noext .. "xlsx"
	log(string.format("Convert %s to %s", filename, tmpname))
	m2x(filename, tmpname)
	assert(winapi.ShellExecute(tmpname), "open " .. tmpname .. " Failed")
	local mo = coroutine.wrap(monitor)
	local ft = mo(tmpname)
	while mo() do
		log(string.format("Convert back %s to %s", tmpname, filename))
		x2m(tmpname, filename)
	end
	if winapi.LastWriteTime(tmpname) ~= ft then
		log(string.format("Convert back %s to %s", tmpname, filename))
		x2m(tmpname, filename)
	end
	log("Remove ", tmpname)
	assert(winapi.DeleteFile(tmpname), "Remove " .. tmpname .. "failed")
	log("Remove Dir", tmpdir)
	assert(winapi.RemoveDirectory(tmpdir), "Remove tmp dir failed")
end

M.m2x = m2x
M.x2m = x2m

local function main(method, ...)
	assert(method, [[

x2m convert xlsx to mfs
m2x convert mfs to xlsx
monitor monitor a mfs
]])
	local f = assert(M[method] , "Undefined method")
	return f(...)
end

local ok, err = pcall(main, ...)
if not ok then
	alert(err)
end
