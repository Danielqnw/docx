require 'docx/containers/table_row'
require 'docx/containers/table_column'
require 'docx/containers/table_grid'
require 'docx/containers/table_merge'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Table
        include Container
        include Elements::Element
        include TableMerge

        @grid_generations = Hash.new(0)

        class << self
          attr_reader :grid_generations
        end

        def self.tag
          'tbl'
        end

        def initialize(node)
          @node = node
          @properties_tag = 'tblGrid'
          @grid_generation = nil
        end

        # Array of row
        def rows
          @node.xpath('w:tr').map {|r_node| Containers::TableRow.new(r_node) }
        end

        def row_count
          @node.xpath('w:tr').count
        end

        # Array of column
        def columns
          (0...column_count).map do |col|
            cells = (0...row_count).map { |row| cell_at(row, col) }
            Containers::TableColumn.new(cells)
          end
        end

        def column_count
          @node.xpath('w:tblGrid/w:gridCol').count
        end

        def cell_at(row, col)
          slot = grid.cell_at(row, col)
          return nil if slot.nil?

          Containers::TableCell.new(slot.node, grid_slot: slot, logical_row: row, logical_col: col)
        end

        def each_cell
          return enum_for(:each_cell) unless block_given?

          grid.each_anchor do |anchor|
            yield(Containers::TableCell.new(anchor.node, grid_slot: anchor), anchor.row, anchor.col)
          end
        end

        def merged?(row, col)
          slot = grid.cell_at(row, col)
          return false if slot.nil?

          slot.colspan > 1 || slot.rowspan > 1
        end

        def invalidate_grid!
          @grid = nil
          self.class.grid_generations[@node.object_id] += 1
        end

        # Iterate over each row within a table
        def each_rows
          rows.each { |r| yield(r) }
        end

        private

        def grid
          current_generation = self.class.grid_generations[@node.object_id]
          if @grid_generation != current_generation
            @grid = nil
            @grid_generation = current_generation
          end

          @grid ||= Containers::TableGrid.new(@node)
        end
      end
    end
  end
end
