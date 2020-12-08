kpse.set_program_name "luatex"
local xlsx = require "spreadsheet.spreadsheet-xlsx-reader"
local log = require "spreadsheet.spreadsheet-log"
log.level = "warn"
local input = arg[1]
local output_dir = arg[2]

-- sloupec, kde začínají katedry
local katedry_start = 5

local category_names = {
"Licencovaný zdroj",
"Volně dostupný zdroj",
"Zkušební přístup"
}



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
  end
  return katedry
end

local function get_eiz(sheet)
  local zdroje = {}
  for i= 2, #sheet do 
    local row = sheet[i]
    local name = get_cell_value(row[1])
    local link = get_cell_link(row[1])
    local description = get_cell_value(row[2])
    local category = get_cell_value(row[3])
    local comment = get_cell_value(row[4])
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

local function save_eiz(katedry, categories)
  for i, katedra in pairs(katedry) do
    print("*******************")
    print(katedra.fullname)
    for _, name in ipairs(category_names) do
      print("====================")
      print(name)
      local current = categories[name]
      for _, zdroj in ipairs(current) do
        if zdroj.katedry[i] then
          print("",zdroj.name)
        end
      end
    end
  end
end

local sheet, xlsx_file = load_table(input)
local katedry = get_katedry(sheet)
get_katedry_names(katedry, xlsx_file)
local zdroje = get_eiz(sheet)
local categories = sort_eiz(zdroje)
save_eiz(katedry, categories)

