-- Keep native Pandoc Figure nodes intact so their caption, identifier, and
-- attributes survive the LaTeX -> Markdown stage.  The only conversion this
-- filter performs is for a standalone raw LaTeX \includegraphics command that
-- the LaTeX reader did not already turn into an Image.

local function trim(value)
  return value:match("^%s*(.-)%s*$")
end

local function normalize_dimension(value)
  value = trim(value)

  local factor, relative_to = value:match("^([%d]*%.?[%d]+)%s*\\([%a]+)$")
  if factor and (
      relative_to == "textwidth" or
      relative_to == "linewidth" or
      relative_to == "columnwidth" or
      relative_to == "textheight") then
    return string.format("%.6g%%", tonumber(factor) * 100)
  end

  local relative_only = value:match("^\\([%a]+)$")
  if relative_only == "textwidth" or
      relative_only == "linewidth" or
      relative_only == "columnwidth" or
      relative_only == "textheight" then
    return "100%"
  end

  if value:match("^[%d]*%.?[%d]+%%$") or
      value:match("^[%d]*%.?[%d]+%s*[A-Za-z]+$") then
    return value:gsub("%s+", "")
  end

  return nil
end

local function parse_options(option_group)
  local attributes = {}
  if not option_group then
    return attributes
  end

  local options = option_group:sub(2, -2)
  for option in options:gmatch("[^,]+") do
    local key, value = option:match("^%s*([%a]+)%s*=%s*(.-)%s*$")
    if key == "width" or key == "height" then
      local dimension = normalize_dimension(value)
      if dimension then
        attributes[key] = dimension
      end
    end
  end
  return attributes
end

local function includegraphics_to_image(raw)
  if raw.format ~= "latex" and raw.format ~= "tex" then
    return nil
  end

  local option_group, target_group = raw.text:match(
    "^%s*\\includegraphics%*?%s*(%b[])%s*(%b{})%s*$"
  )
  if not target_group then
    target_group = raw.text:match(
      "^%s*\\includegraphics%*?%s*(%b{})%s*$"
    )
  end
  if not target_group then
    return nil
  end

  local image = pandoc.Image({}, target_group:sub(2, -2))
  for key, value in pairs(parse_options(option_group)) do
    image.attributes[key] = value
  end
  return image
end

function RawInline(raw)
  return includegraphics_to_image(raw)
end

function RawBlock(raw)
  local image = includegraphics_to_image(raw)
  if image then
    return pandoc.Plain({image})
  end
  return nil
end
