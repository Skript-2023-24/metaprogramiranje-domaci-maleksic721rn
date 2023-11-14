require './maleksic721rn_table.rb'

LOAD_FILE = "maleksic721rn_example.xlsx"
LOAD_FILE_2 = "maleksic721rn_example.xls"

tbl = Worksheet::Load(LOAD_FILE)
tbl2 = Worksheet::Load(LOAD_FILE_2)

p tbl.cells_copy
p
p tbl.map(&:inspect)
p
p tbl.row(0)
p tbl.row(3)
p tbl.row(4)
p
p tbl['ABC'].to_s
p tbl['ABC'][1]
tbl['ABC'][1] = tbl['ABC'][1] + 1
p tbl['ABC'][1]
p
p tbl.dobra_kolona.to_s
p tbl.dobra_kolona.sum
p tbl.dobra_kolona.avg
p tbl.dobra_kolona.test
p tbl.dobra_kolona.map(&:inspect)
p
p tbl.row(12)

tbl.set_raw(15, 1, 0)

tbl.save("_" + LOAD_FILE)
(tbl + tbl2).save("_union_" + LOAD_FILE)
(tbl - tbl2).save("_diff_" + LOAD_FILE)