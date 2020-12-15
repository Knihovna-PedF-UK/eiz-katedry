local lustache = require "lustache"
kpse.set_program_name "luatex"
local xlsx = require "spreadsheet.spreadsheet-xlsx-reader"
local log = require "spreadsheet.spreadsheet-log"
log.level = "warn"
local input = arg[1]
local output_dir = arg[2] or "out"

-- sloupec, kde začínají katedry
local katedry_start = 6

local category_names = {
"Licencovaný zdroj",
"Volně dostupný zdroj",
"Zkušební přístup"
}

local template = [[
---
title: {{name}} – elektronické informační zdroje
img: /img/eiz-{{shortcut}}.jpg
alt: "{{fullname}}"
---
<h1>{{fullname}}</h1>
{{#sources}}
<h2>{{name}}</h2>
<ul>
{{#entries}}
<li><a href="{{link}}">{{name}}</a> – {{description}}</li>
{{/entries}}
</ul>
{{/sources}}
]]

local index_tpl = [[
---
title: Elektronické informační zdroje pro katedry
---
<h1>Elektronické informační zdroje pro katedry</h1>
<ul>
{{#katedry}}
<li><a href="{{filename}}">{{fullname}}</a></li>
{{/katedry}}
</ul>
]]




local function load_table(input)
  local lo,msg  = xlsx.load(input)
  if not lo then 
    print(msg)
    os.exit()
  end
  local sheet = lo:get_sheet(1)
  return sheet.table, lo
end

local function get_cell_value(cell)
  -- vrať první objekt v buňce
  local cell = cell or {}
  local first = cell[1] or {}
  return first.value
end

local function get_cell_link(cell)
  -- najdi odkaz
  local cell = cell or {}
  local first = cell[1] or {}
  local style = first.style or {}
  return style.link
end


local function get_katedry(sheet)
  -- vrať tabulku s katedrami
  local katedry = {}
  local first_row = sheet[1]
  -- uložíme objekt pro katedru na pozici ve sloupci
  for i= katedry_start, #first_row do
    local value = get_cell_value(first_row[i])
    if value then
      katedry[i] = {name = value, databases = {}}
    end
  end
  return katedry
end

local function get_katedry_names(katedry, xlsx_file)
  -- načíst jména kateder z druhého listu
  local katedry_sheet = xlsx_file:get_sheet(2)
  local t = {}
  for k, v in ipairs(katedry_sheet.table) do
    -- ulož mapování mezi zkratkou katedry a jejím jménem
    t[get_cell_value(v[1])] = get_cell_value(v[2])
  end
  for _, katedra in pairs(katedry) do
    katedra.fullname =  t[katedra.name]
    -- ulož zkratku katerdry v lowercase. využijeme ve jménech obrázků
    katedra.shortcut  = string.lower(katedra.name)
  end
  return katedry
end

local function get_eiz(sheet)
  local zdroje = {}
  for i= 2, #sheet do 
    local row = sheet[i]
    local name = get_cell_value(row[1])
    local link = get_cell_value(row[2])
    local description = get_cell_value(row[3])
    local category = get_cell_value(row[4])
    local comment = get_cell_value(row[5])
    local object = {name = name, link = link, description = description, category = category, comment = comment, katedry = {}}
    for i = katedry_start, #row do 
      local value = get_cell_value(row[i])
      if value and value ~= "" then object.katedry[i] = true end
    end
    table.insert(zdroje, object)
  end
  return zdroje
end

local function sort_eiz(zdroje)
  local categories = {}
  -- nejdřív zdroje setřídíme podle abecedy
  table.sort(zdroje, function(a,b) return a.name < b.name end)
  for _, zdroj in ipairs(zdroje) do
    local cat_name = zdroj.category
    category = categories[cat_name] or {}
    table.insert(category, zdroj)
    categories[cat_name] = category
  end
  return categories
end

local lower = unicode.utf8.lower
local gmatch = unicode.utf8.gmatch
local function detox(str)
  local replace = {["č"]="c", ["š"] = "s"}
  local str = lower(str)
  local t = {}
  for c in gmatch(str, ".") do 
    t[#t+1] = replace[c] or c
  end
  return table.concat(t)
end

local used_names = {}

local function make_output_name(name)
  local filename = output_dir .. "/eiz-" .. detox(name) .. ".html"
  return filename:gsub("/+", "/") -- normalizovat cesty
end

local function save_katedra(katedra, text)
  local filename = make_output_name(katedra.name)
  -- ulozit jmeno HTML souboru, ale bez adresare 
  katedra.filename = filename:gsub("^[^%/]+/", "")
  -- testovat, jestli neexistuje kolize jmen
  if used_names[filename] then
    print("error: filename " .. filename .. "already exists")
  end
  used_names[filename] = true
  local f = io.open(filename, "w")
  print("saving", filename)
  f:write(text)
  f:close()
end

local function save_eiz(katedry, categories)
  for i, katedra in pairs(katedry) do
    -- ziskat zdroje dostupne pro katedru
    katedra.sources = {}
    -- zpracovat jednotlive kategorie zdroju
    for _, name in ipairs(category_names) do
      -- ulozit dostupne zdroje do nove tabulky
      local current = categories[name]
      local t = {}
      for _, zdroj in ipairs(current) do
        if zdroj.katedry[i] then
          t[#t+1] = zdroj
        end
      end
      -- pokud ma katedra nejake zdroje dane kategorie, ulozit 
      if #t > 0 then
        table.insert(katedra.sources, {name = name, entries = t})
      end
    end
    -- zobrazit zdroj
    save_katedra(katedra, lustache:render(template, katedra))
  end
end

local function save_index(katedry)
  local filename = make_output_name("katedry")
  local f = io.open(filename, "w")
  local t = {katedry = {}}
  -- katedry nezacinaji od nuly, musime to pole preindexovat
  for k,v in pairs(katedry.katedry) do
    table.insert(t.katedry, v)
  end
  local text = lustache:render(index_tpl, t)
  f:write(text)
  f:close()
end

local sheet, xlsx_file = load_table(input)
local katedry = get_katedry(sheet)
get_katedry_names(katedry, xlsx_file)
local zdroje = get_eiz(sheet)
local categories = sort_eiz(zdroje)
save_eiz(katedry, categories)
save_index({katedry = katedry}) -- trik aby fungovala šablona. potřebuje asociativní pole

