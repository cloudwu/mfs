local zip = require "zip"
local base64 = require "base64"

local mfs, xlsx = ...
assert(mfs, "Need mfs filename")
assert(xlsx, "Need xlsx filename")

local source = assert(io.open(mfs), "rb")
local zip = assert(zip.zip(xlsx))

local content
local blob
local filename

for line in source:lines() do
	local t = line:sub(1,1)
	if t == "*" or t == "+" then
		if content then
			zip:write(filename, table.concat(content))
		end
		filename = line:sub(2)
		if t == "*" then
			content = {}
		else
			blob = true
			content = nil
		end
	elseif blob then
		zip:write(filename, base64.decode(line))
		blob = nil
	else
		assert(t == "<")
		table.insert(content, line)
	end
end

if blob then
	zip:write(filename, base64.decode(line))
elseif content then
	zip:write(filename, table.concat(content))
end

source:close()
zip:close()
