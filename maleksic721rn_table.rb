require 'roo'
require 'roo-xls'
require 'rubyXL'
require 'spreadsheet'
require 'matrix'

class Column
	include Enumerable
	
	def initialize(tbl, i)
		@CTable = tbl
		@ColumnIndex = i
	end
	
	def col
		@CTable.col(@ColumnIndex)
	end
	
	def row_count
		@CTable.row_count
	end
	
	def at_raw(i)
		@CTable.cell_raw(i, @ColumnIndex)
	end
	
	def [](i)
		col[@CTable.eff_row(i)]
	end
	
	def []=(i, value)
		@CTable.set(i, @ColumnIndex, value)
	end
	
	def to_s
		col.to_s
	end
	
	def each
		@CTable.row_count_raw.times { |i| yield self[i] if @CTable.counted? i }
	end
	
	def num_values
		col.select.with_index { |e, i| e.is_a?(Numeric) and @CTable.counted? i }
	end
	
	def sum
		num_values.reduce(0, :+)
	end
	
	def method_missing(m)
		@CTable.row(col.find_index(m.to_s) || @CTable.row_count + 1)
	end
	
	def avg
		1.0 * sum / num_values.size
	end
end

class Table
	include Enumerable
	
	def initialize(tbl)
		@InnerTable = tbl
		@Columns = Array.new(col_count) { |i| Column.new(self, i) }
	end
	
	def row_count_raw
		@InnerTable.last_row - @InnerTable.first_row + 1
	end
	
	def row_count
		0.upto(row_count_raw - 1).count { |i| counted?(i) }
	end
	
	def col_count
		@InnerTable.last_column - @InnerTable.first_column + 1
	end
	
	def cell(r, c)
		cell_raw(eff_row(r), c)
	end
	
	def cell_raw(r, c)
		@InnerTable.cell(r + @InnerTable.first_row, c + @InnerTable.first_column)
	end
	
	def set(r, c, value)
		set_raw(eff_row(r), c, value)
	end
	
	def set_raw(r, c, value)
		@InnerTable.set(r + @InnerTable.first_row, c + @InnerTable.first_column, value)
	end
	
	def cells_copy
		mat = Matrix.build(row_count_raw, col_count) { nil }
		row_count_raw.times { |r| col_count.times { |c| mat[r, c] = cell(r, c) } }
		return mat
	end
	
	def formula(r, c)
		formula_raw(eff_row(r), c)
	end
	
	def formula_raw(r, c)
		return nil if @InnerTable.is_a? Roo::Excel
		@InnerTable.formula(r + @InnerTable.first_row, c + @InnerTable.first_column)
	end
	
	def row(r)
		row_raw(eff_row(r))
	end
	
	def row_raw(r)
		@InnerTable.row(r + @InnerTable.first_row)
	end
	
	def col(c)
		@InnerTable.column(c + @InnerTable.first_column)
	end
	
	def [](c)
		@Columns[row(0).find_index(c) || (@Columns.size + 1)]
	end
	
	def each
		row_count_raw.times { |r| col_count.times { |c| yield cell(r, c) if counted? r } }
	end
	
	def method_missing(m)
		self[m.to_s.gsub("_", " ")]
	end
	
	def eff_row(r)
		[((r.upto(row_count_raw - 1).to_a.take_while { |i| not counted? i }).last || r - 1) + 1, row_count_raw - 1].min
	end
	
	def counted?(r)
		not (subtotal?(r) or empty?(r))
	end
	
	def subtotal?(r)
		@InnerTable.row(r + @InnerTable.first_row).grep(String).map(&:downcase).any? { |v| v == "total" or v == "subtotal" }
	end
	
	def empty?(r)
		@InnerTable.row(r + @InnerTable.first_row).all? nil
	end
	
	def +(t)
		return self if row(0).union(t.row(0)).size != row(0).size
		res = Table.new(@InnerTable.clone)
		crow = row_count_raw
		row_count_raw.upto(row_count_raw + t.row_count_raw - 2).each do |r|
			next if 1.upto(row_count_raw - 1).any? { |i| row_raw(i) == t.row_raw(r - row_count_raw) }
			res.row(0).each.with_index do |c, ci|
				res.set_raw(crow, ci, t.formula_raw(r, ci) || t.cell_raw(r, ci) || "")
			end
			crow += 1
		end
		return res
	end
	
	def -(t)
		return self if row(0).union(t.row(0)).size != row(0).size
		res = Table.new(@InnerTable.clone)
		crow = 1
		1.upto(row_count_raw - 1).each do |r|
			next if 1.upto(t.row_count_raw - 1).any? { |i| t.row_raw(i) == row_raw(r) }
			0.upto(col_count).each do |c|
				res.set_raw(crow, c, res.cell_raw(r, c) || "")
			end
			crow += 1
		end
		crow.upto(row_count_raw - 1).each do |r|
			0.upto(col_count).each do |c|
				res.set_raw(r, c, "")
			end
		end
		return res
	end
	
	def save(f)
		book = nil
		if @InnerTable.is_a? Roo::Excelx
			book = RubyXL::Workbook.new
			row_count_raw.times { |r| col_count.times do |c|
				book[0].add_cell(r + @InnerTable.first_row - 1, c + @InnerTable.first_column - 1, formula_raw(r, c) ? cell_raw(r, c) : '', formula_raw(r, c))
			end }
		else
			book = Spreadsheet::Workbook.new
			sheet = book.create_worksheet
			row_count_raw.times { |r| col_count.times { |c| sheet[r + @InnerTable.first_row, c + @InnerTable.first_column] = cell_raw(r, c) } }
		end
		book.write f
	end
end

module Worksheet
	def Load(path)
		Table.new(File.extname(path) == ".xlsx" ? Roo::Excelx.new(path, {:expand_merged_ranges => false}) : Roo::Excel.new(path, {:expand_merged_ranges => false}))
	end
	
	module_function :Load
end
