require 'docx/containers/table_row'
require 'docx/containers/table_column'
require 'docx/containers/table_grid'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Table
        include Container
        include Elements::Element

        def self.tag
          'tbl'
        end

        def initialize(node)
          @node = node
          @properties_tag = 'tblGrid'
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

          Containers::TableCell.new(slot.node)
        end

        def each_cell
          return enum_for(:each_cell) unless block_given?

          grid.each_anchor do |anchor|
            yield(Containers::TableCell.new(anchor.node), anchor.row, anchor.col)
          end
        end

        def invalidate_grid!
          @grid = nil
        end

        # Iterate over each row within a table
        def each_rows
          rows.each { |r| yield(r) }
        end

        private

        def grid
          @grid ||= Containers::TableGrid.new(@node)
        end
      end
    end
  end
end
